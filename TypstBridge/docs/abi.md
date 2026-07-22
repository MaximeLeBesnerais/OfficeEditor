# TypstBridge ABI

This document defines the C ABI between the Rust Typst bridge and .NET managed wrapper.

The ABI is intentionally C-compatible. Rust-owned memory must always be released by Rust-provided free functions. C# must never free Rust allocations directly.

## ABI versions

- **v1/v2** — single-shot compile entry point (`typst_bridge_compile`).
- **v3 (current)** — adds persistent compile sessions (`typst_bridge_session_*`), explicit memoization eviction (`typst_bridge_evict_cache`), and an internal native font cache. `typst_bridge_abi_version` returns `3`.

The v2 single-shot entry point remains supported in v3: the `typst_bridge_compile_request` layout is unchanged, and the native bridge accepts requests carrying `abi_version` 2 or 3. The managed wrapper enforces exact ABI equality (`TypstBridgeCompiler.SupportedAbiVersion`), so managed and native builds must ship in lockstep.

## Design Rules

- All string inputs are UTF-8 with explicit lengths.
- All returned buffers and strings are owned by the bridge.
- The caller must release returned results with `typst_bridge_free_result`.
- Session handles are released with `typst_bridge_session_free`.
- Panics must never cross the FFI boundary.
- Compile results must support one output item for PDF and one output item per page for PNG/SVG.

## Thread-safety guarantees

- `typst_bridge_compile` is reentrant: all per-call state is local, the parsed-font cache is `Mutex`-guarded, and comemo's global memoization caches are internally synchronized.
- A session handle serializes its own calls behind an internal mutex, so one handle may be shared across threads (calls are serialized). Distinct sessions compile fully independently and in parallel.
- The last-error string is `thread_local`, so concurrent calls on different threads (including concurrent sessions) never clobber each other's error messages.

## Native font cache

Since ABI v3 the bridge memoizes the parsed font set (`FontBook` + `Vec<Font>`) process-wide, keyed by the sorted caller-supplied font-path list plus the working directory and the system/embedded-fonts flag state. A cache hit skips directory scanning, `fs::read` of font files, and `Font` parsing. Entries are shared by `Arc`.

Tradeoff: the cache is never evicted. Its size is bounded in practice by the number of distinct font-path lists a process compiles with (path-keyed; one entry per distinct font directory set). Font files changed on disk are not re-read until process restart.

## Memoization eviction policy

typst/comemo memoize layout work process-wide; without eviction those caches grow unboundedly in a long-lived server. The bridge applies the typst-cli watch-mode pattern:

- Every finished compile (single-shot or session, success or failure) increments a process-wide counter.
- Every 10 finished compiles, one `comemo::evict(30)` pass runs: entries not touched within the last 30 eviction passes are reclaimed. Memoized data unused for roughly 300 compiles is dropped; actively-used session state survives.
- `typst_bridge_evict_cache(max_age)` runs a pass on demand (memory pressure, idle transitions). `max_age = 0` evicts everything not currently referenced.

Eviction only drops memoization entries. Recomputation is deterministic (`World::today` is pinned to 1970-01-01 and font collection order is sorted), so output bytes are identical across eviction boundaries.

## Implemented compile surface

The current native bridge uses this ABI for source-string compilation with working-directory asset resolution, explicit font paths, PDF output, per-page SVG/PNG outputs, PNG PPI, and diagnostics. OfficeEditor.Core's `TypstCompilerService` now uses TypstBridge as its primary backend, and the external Typst CLI remains the safety-net fallback.

## Status Codes

```c
typedef enum typst_bridge_status {
    TYPST_BRIDGE_OK = 0,
    TYPST_BRIDGE_ERR_INVALID_ARGUMENT = 1,
    TYPST_BRIDGE_ERR_COMPILE = 2,
    TYPST_BRIDGE_ERR_RENDER = 3,
    TYPST_BRIDGE_ERR_IO = 4,
    TYPST_BRIDGE_ERR_PANIC = 5,
    TYPST_BRIDGE_ERR_UNSUPPORTED = 6,
    TYPST_BRIDGE_ERR_INTERNAL = 255
} typst_bridge_status;
```

## Output Format

```c
typedef enum typst_bridge_output_format {
    TYPST_BRIDGE_OUTPUT_PDF = 1,
    TYPST_BRIDGE_OUTPUT_PNG = 2,
    TYPST_BRIDGE_OUTPUT_SVG = 3
} typst_bridge_output_format;
```

## Input String

```c
typedef struct typst_bridge_string {
    const char* value_utf8;
    uintptr_t value_len;
} typst_bridge_string;
```

## Compile Request

```c
typedef struct typst_bridge_compile_request {
    uint32_t abi_version;

    const char* source_utf8;
    uintptr_t source_len;

    const char* working_dir_utf8;
    uintptr_t working_dir_len;

    const char* root_file_name_utf8;
    uintptr_t root_file_name_len;

    const typst_bridge_string* font_paths;
    uintptr_t font_paths_count;

    typst_bridge_output_format output_format;
    double ppi;
    uint32_t flags;
} typst_bridge_compile_request;
```

`working_dir_utf8` is required for PPTX-generated sources because the converter emits relative asset paths such as `assets/...`.

Font paths use the same explicit pointer + length string rule as scalar strings.
When `font_paths_count > 0`, `font_paths` must point to an array of `font_paths_count` items. Each `value_utf8`/`value_len` pair must be valid UTF-8 and must identify a file or directory path. Do not supply zero-length font path entries; they are reserved/invalid until the implementation gives them explicit semantics. Current implementations may treat an empty path like `working_dir`, so callers should omit such entries instead.

## Output Item

```c
typedef struct typst_bridge_output_item {
    uint32_t page_index;
    char* file_name_utf8;
    uintptr_t file_name_len;
    uint8_t* data;
    uintptr_t data_len;
} typst_bridge_output_item;
```

Expected behavior:

- PDF returns one item with `page_index = 0`.
- PNG returns one item per page, ordered by page index.
- SVG returns one item per page, ordered by page index.

## Diagnostic Item

```c
typedef struct typst_bridge_diagnostic {
    uint32_t severity;
    char* message_utf8;
    uintptr_t message_len;
    char* file_utf8;
    uintptr_t file_len;
    uint32_t line;
    uint32_t column;
} typst_bridge_diagnostic;
```

Severity values:

```c
#define TYPST_BRIDGE_DIAG_ERROR 1
#define TYPST_BRIDGE_DIAG_WARNING 2
#define TYPST_BRIDGE_DIAG_INFO 3
```

## Compile Result

```c
typedef struct typst_bridge_compile_result {
    typst_bridge_status status;

    typst_bridge_output_item* outputs;
    uintptr_t outputs_count;

    typst_bridge_diagnostic* diagnostics;
    uintptr_t diagnostics_count;

    char* message_utf8;
    uintptr_t message_len;
} typst_bridge_compile_result;
```

## Functions

```c
uint32_t typst_bridge_abi_version(void);
const char* typst_bridge_version(void);
typst_bridge_status typst_bridge_probe(void);

typst_bridge_compile_result* typst_bridge_compile(
    const typst_bridge_compile_request* request
);

void typst_bridge_free_result(
    typst_bridge_compile_result* result
);

void typst_bridge_free_string(
    char* value
);

const char* typst_bridge_last_error_message(void);
```

## Session functions (ABI v3)

A session keeps a warm Typst world — library, parsed font set, and working-directory resolution — alive across compiles, so comemo-memoized layout work survives an edited-source recompile.

```c
typedef struct typst_bridge_session typst_bridge_session; // opaque handle

typst_bridge_status typst_bridge_session_create(
    const char* working_dir_utf8,
    uintptr_t working_dir_len,
    const typst_bridge_string* font_paths,
    uintptr_t font_paths_count,
    typst_bridge_session** out_session
);

typst_bridge_status typst_bridge_session_update_source(
    typst_bridge_session* session,
    const char* source_utf8,
    uintptr_t source_len,
    const char* root_file_name_utf8,
    uintptr_t root_file_name_len
);

typst_bridge_compile_result* typst_bridge_session_compile(
    typst_bridge_session* session,
    uint32_t output_format,
    double ppi
);

void typst_bridge_session_free(
    typst_bridge_session* session
);

typst_bridge_status typst_bridge_evict_cache(
    uint32_t max_age
);
```

Session lifecycle and rules:

- `typst_bridge_session_create` validates the working directory and font paths (same rules as the compile request; an empty `working_dir` falls back to the process current directory) and writes an opaque handle to `out_session`. On failure it returns a non-OK status, leaves `*out_session` null, and sets the thread-local error string.
- `typst_bridge_session_update_source` replaces the main source in place. An empty `root_file_name` behaves like the compile request (defaults to `main.typ`). Compiling before the first update returns `TYPST_BRIDGE_ERR_INVALID_ARGUMENT`.
- `typst_bridge_session_compile` accepts the raw `typst_bridge_output_format` value and `ppi` (same validation as the compile request) and returns the exact same result layout and ownership rules as `typst_bridge_compile` — release it with `typst_bridge_free_result`.
- `typst_bridge_session_free` releases the handle and is null-safe. Freeing the same handle twice is a double-free, as with any Rust-owned pointer.
- Session output is byte-identical to a cold `typst_bridge_compile` of the same source, working directory, font paths, format, and ppi.

## Acceptance Criteria

- Invalid requests return structured errors where possible.
- Panics are caught and returned as `TYPST_BRIDGE_ERR_PANIC`.
- Repeated compile/free cycles do not crash.
- Session create/update/compile/free cycles do not crash, and recompiles of unchanged sources are byte-deterministic — including across memoization eviction boundaries.
- Output signatures are correct: `%PDF`, PNG magic bytes, and SVG/XML text.
