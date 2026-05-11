use std::ffi::{c_char, c_uchar};

pub const TYPST_BRIDGE_ABI_VERSION: u32 = 2;

pub const TYPST_BRIDGE_DIAG_ERROR: u32 = 1;
pub const TYPST_BRIDGE_DIAG_WARNING: u32 = 2;
#[allow(dead_code)]
pub const TYPST_BRIDGE_DIAG_INFO: u32 = 3;

#[repr(u32)]
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum TypstBridgeStatus {
    Ok = 0,
    InvalidArgument = 1,
    Compile = 2,
    Render = 3,
    Io = 4,
    Panic = 5,
    Unsupported = 6,
    Internal = 255,
}

#[repr(u32)]
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum TypstBridgeOutputFormat {
    Pdf = 1,
    Png = 2,
    Svg = 3,
}

impl TypstBridgeOutputFormat {
    pub fn from_raw(value: u32) -> Option<Self> {
        match value {
            1 => Some(Self::Pdf),
            2 => Some(Self::Png),
            3 => Some(Self::Svg),
            _ => None,
        }
    }
}

#[repr(C)]
#[derive(Clone, Copy)]
pub struct TypstBridgeString {
    pub value_utf8: *const c_char,
    pub value_len: usize,
}

#[repr(C)]
#[derive(Clone, Copy)]
pub struct TypstBridgeCompileRequest {
    pub abi_version: u32,
    pub source_utf8: *const c_char,
    pub source_len: usize,
    pub working_dir_utf8: *const c_char,
    pub working_dir_len: usize,
    pub root_file_name_utf8: *const c_char,
    pub root_file_name_len: usize,
    pub font_paths: *const TypstBridgeString,
    pub font_paths_count: usize,
    pub output_format: u32,
    pub ppi: f64,
    pub flags: u32,
}

#[repr(C)]
pub struct TypstBridgeOutputItem {
    pub page_index: u32,
    pub file_name_utf8: *mut c_char,
    pub file_name_len: usize,
    pub data: *mut c_uchar,
    pub data_len: usize,
}

#[repr(C)]
pub struct TypstBridgeDiagnostic {
    pub severity: u32,
    pub message_utf8: *mut c_char,
    pub message_len: usize,
    pub file_utf8: *mut c_char,
    pub file_len: usize,
    pub line: u32,
    pub column: u32,
}

#[repr(C)]
pub struct TypstBridgeCompileResult {
    pub status: TypstBridgeStatus,
    pub outputs: *mut TypstBridgeOutputItem,
    pub outputs_count: usize,
    pub diagnostics: *mut TypstBridgeDiagnostic,
    pub diagnostics_count: usize,
    pub message_utf8: *mut c_char,
    pub message_len: usize,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn abi_values_match_documentation() {
        assert_eq!(TypstBridgeStatus::Ok as u32, 0);
        assert_eq!(TypstBridgeStatus::Unsupported as u32, 6);
        assert_eq!(TypstBridgeStatus::Internal as u32, 255);
        assert_eq!(TYPST_BRIDGE_DIAG_ERROR, 1);
        assert_eq!(TYPST_BRIDGE_DIAG_WARNING, 2);
        assert_eq!(TYPST_BRIDGE_DIAG_INFO, 3);
        assert_eq!(TypstBridgeOutputFormat::Pdf as u32, 1);
        assert_eq!(TypstBridgeOutputFormat::Png as u32, 2);
        assert_eq!(TypstBridgeOutputFormat::Svg as u32, 3);
    }
}
