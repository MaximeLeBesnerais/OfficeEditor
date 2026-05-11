mod abi;
mod compiler;
mod diagnostics;
mod memory;

use std::cell::RefCell;
use std::ffi::{c_char, CString};
use std::panic;

pub use abi::{TypstBridgeCompileRequest, TypstBridgeCompileResult, TypstBridgeStatus};

use abi::TYPST_BRIDGE_ABI_VERSION;

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

fn set_last_error(message: &str) {
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
        assert_eq!(typst_bridge_abi_version(), 2);
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
