//! Bounded-memoization policy for long-lived processes.
//!
//! typst/comemo memoize layout and introspection work process-wide. Without
//! eviction those caches grow unboundedly, which is a slow memory leak for a
//! long-lived preview server. This module implements the typst-cli watch-mode
//! pattern (`comemo::evict` on a cadence) instead of typst-cli's default
//! "compile once and exit" behavior, which never evicts.
//!
//! Policy:
//! - Every finished compile (single-shot or session, success or failure)
//!   increments a process-wide counter.
//! - Every [`EVICT_INTERVAL_COMPILES`] compiles, one `comemo::evict` pass runs
//!   with `max_age = [`EVICT_MAX_AGE`]. Entries not touched within the last
//!   `EVICT_MAX_AGE` eviction passes are reclaimed, so memoized data unused
//!   for roughly `EVICT_INTERVAL_COMPILES * EVICT_MAX_AGE` compiles is dropped
//!   while actively-used session state survives.
//! - `typst_bridge_evict_cache(max_age)` lets the host force a pass on demand
//!   (e.g. on memory pressure or idle transitions); `max_age = 0` evicts
//!   everything not currently referenced.
//!
//! Eviction only drops memoization entries; recomputation is deterministic
//! (`World::today` is pinned, font collection is sorted), so output bytes are
//! unaffected by eviction boundaries.

use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::Mutex;

/// Number of finished compiles between automatic eviction passes.
pub const EVICT_INTERVAL_COMPILES: u64 = 10;

/// Entries unused for more than this many eviction passes are reclaimed.
pub const EVICT_MAX_AGE: u32 = 30;

static COMPILES_FINISHED: AtomicU64 = AtomicU64::new(0);
static EVICT_LOCK: Mutex<()> = Mutex::new(());

/// Records a finished compile and runs an eviction pass when the cadence
/// boundary is reached.
pub fn note_compile_finished() {
    let finished = COMPILES_FINISHED.fetch_add(1, Ordering::Relaxed) + 1;
    if finished % EVICT_INTERVAL_COMPILES == 0 {
        evict(EVICT_MAX_AGE);
    }
}

/// Runs one `comemo::evict` pass, serialized against concurrent passes.
pub fn evict(max_age: u32) {
    let _guard = EVICT_LOCK
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner());
    comemo::evict(max_age as usize);
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn eviction_passes_do_not_fail_across_boundaries() {
        for _ in 0..EVICT_INTERVAL_COMPILES {
            note_compile_finished();
        }
        evict(0);
        evict(EVICT_MAX_AGE);
    }
}
