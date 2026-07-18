mod abi;
mod compiler;
mod diagnostics;
mod evict;
mod fonts;
mod memory;
mod render_pdf;
mod render_png;
mod render_svg;
mod session;
mod world;

use std::cell::RefCell;
use std::ffi::{c_char, CString};
use std::panic;

pub use abi::{TypstBridgeCompileRequest, TypstBridgeCompileResult, TypstBridgeStatus};

use abi::{TypstBridgeString, TYPST_BRIDGE_ABI_VERSION};
use session::BridgeSession;

const VERSION: &[u8] = concat!(env!("CARGO_PKG_VERSION"), "\0").as_bytes();
const PANIC_MESSAGE: &str = "TypstBridge native bridge panicked";

thread_local! {
    static LAST_ERROR_MESSAGE: RefCell<CString> = RefCell::new(CString::new("").expect("empty string is valid CString"));
}

#[no_mangle]
pub extern "C" fn typst_bridge_abi_version() -> u32 {
    TYPST_BRIDGE_ABI_VERSION
}

#[no_mangle]
pub extern "C" fn typst_bridge_version() -> *const c_char {
    VERSION.as_ptr().cast::<c_char>()
}

#[no_mangle]
pub extern "C" fn typst_bridge_probe() -> TypstBridgeStatus {
    match panic::catch_unwind(|| TypstBridgeStatus::Ok) {
        Ok(status) => status,
        Err(_) => {
            set_last_error(PANIC_MESSAGE);
            TypstBridgeStatus::Panic
        }
    }
}

#[no_mangle]
pub extern "C" fn typst_bridge_compile(
    request: *const TypstBridgeCompileRequest,
) -> *mut TypstBridgeCompileResult {
    match panic::catch_unwind(|| compiler::compile(request)) {
        Ok(result) => result,
        Err(_) => {
            set_last_error(PANIC_MESSAGE);
            compiler::panic_result(PANIC_MESSAGE)
        }
    }
}

#[no_mangle]
pub extern "C" fn typst_bridge_free_result(result: *mut TypstBridgeCompileResult) {
    let free = || unsafe { memory::free_result(result) };
    if panic::catch_unwind(free).is_err() {
        set_last_error(PANIC_MESSAGE);
    }
}

#[no_mangle]
pub extern "C" fn typst_bridge_free_string(value: *mut c_char) {
    let free = || unsafe { memory::free_string(value) };
    if panic::catch_unwind(free).is_err() {
        set_last_error(PANIC_MESSAGE);
    }
}

#[no_mangle]
pub extern "C" fn typst_bridge_last_error_message() -> *const c_char {
    LAST_ERROR_MESSAGE.with(|message| message.borrow().as_ptr())
}

// --- ABI v3: persistent compile sessions -----------------------------------

/// Creates a session holding a warm world (library + font set + working
/// directory). On success writes an opaque handle to `out_session`; the caller
/// must release it with `typst_bridge_session_free`.
#[no_mangle]
pub extern "C" fn typst_bridge_session_create(
    working_dir_utf8: *const c_char,
    working_dir_len: usize,
    font_paths: *const TypstBridgeString,
    font_paths_count: usize,
    out_session: *mut *mut BridgeSession,
) -> TypstBridgeStatus {
    match panic::catch_unwind(|| {
        session::create(
            working_dir_utf8,
            working_dir_len,
            font_paths,
            font_paths_count,
            out_session,
        )
    }) {
        Ok(status) => status,
        Err(_) => {
            set_last_error(PANIC_MESSAGE);
            TypstBridgeStatus::Panic
        }
    }
}

/// Replaces the session's main source in place, keeping the world hot.
#[no_mangle]
pub extern "C" fn typst_bridge_session_update_source(
    session: *mut BridgeSession,
    source_utf8: *const c_char,
    source_len: usize,
    root_file_name_utf8: *const c_char,
    root_file_name_len: usize,
) -> TypstBridgeStatus {
    match panic::catch_unwind(|| {
        session::update_source(
            session,
            source_utf8,
            source_len,
            root_file_name_utf8,
            root_file_name_len,
        )
    }) {
        Ok(status) => status,
        Err(_) => {
            set_last_error(PANIC_MESSAGE);
            TypstBridgeStatus::Panic
        }
    }
}

/// Compiles the session's current source. Same result layout and ownership
/// rules as `typst_bridge_compile`.
#[no_mangle]
pub extern "C" fn typst_bridge_session_compile(
    session: *mut BridgeSession,
    output_format: u32,
    ppi: f64,
) -> *mut TypstBridgeCompileResult {
    match panic::catch_unwind(|| session::compile(session, output_format, ppi)) {
        Ok(result) => result,
        Err(_) => {
            set_last_error(PANIC_MESSAGE);
            compiler::panic_result(PANIC_MESSAGE)
        }
    }
}

/// Releases a session handle. Null-safe; calling it more than once on the same
/// handle is a double-free and forbidden, as with any Rust-owned pointer.
#[no_mangle]
pub extern "C" fn typst_bridge_session_free(session: *mut BridgeSession) {
    let free = || session::free(session);
    if panic::catch_unwind(free).is_err() {
        set_last_error(PANIC_MESSAGE);
    }
}

/// Runs one comemo eviction pass. See `evict.rs` for the automatic cadence
/// policy that complements this explicit hook.
#[no_mangle]
pub extern "C" fn typst_bridge_evict_cache(max_age: u32) -> TypstBridgeStatus {
    match panic::catch_unwind(|| evict::evict(max_age)) {
        Ok(()) => TypstBridgeStatus::Ok,
        Err(_) => {
            set_last_error(PANIC_MESSAGE);
            TypstBridgeStatus::Panic
        }
    }
}

pub(crate) fn set_last_error(message: &str) {
    let sanitized = message.replace('\0', "�");
    LAST_ERROR_MESSAGE.with(|last_error| {
        *last_error.borrow_mut() = CString::new(sanitized).expect("interior NULs were removed");
    });
}

#[cfg(test)]
mod tests {
    use std::ffi::CStr;

    use super::*;

    #[test]
    fn version_and_probe_are_available() {
        assert_eq!(typst_bridge_abi_version(), 3);
        assert_eq!(typst_bridge_probe(), TypstBridgeStatus::Ok);

        let version = unsafe { CStr::from_ptr(typst_bridge_version()) };
        assert!(!version.to_str().unwrap().is_empty());
    }

    #[test]
    fn exported_compile_handles_null_request() {
        let result = typst_bridge_compile(std::ptr::null());
        assert!(!result.is_null());
        unsafe {
            assert_eq!((*result).status, TypstBridgeStatus::InvalidArgument);
        }
        typst_bridge_free_result(result);
        typst_bridge_free_result(std::ptr::null_mut());
    }
}
