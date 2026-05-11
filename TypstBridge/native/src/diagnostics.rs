use std::ptr;

use crate::abi::{TypstBridgeDiagnostic, TYPST_BRIDGE_DIAG_ERROR};
use crate::memory::into_raw_string;

pub fn error(message: &str) -> TypstBridgeDiagnostic {
    let (message_utf8, message_len) = into_raw_string(message);

    TypstBridgeDiagnostic {
        severity: TYPST_BRIDGE_DIAG_ERROR,
        message_utf8,
        message_len,
        file_utf8: ptr::null_mut(),
        file_len: 0,
        line: 0,
        column: 0,
    }
}
