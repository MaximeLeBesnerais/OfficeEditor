use std::ffi::c_char;
use std::slice;
use std::str;

use crate::abi::{
    TypstBridgeCompileRequest, TypstBridgeCompileResult, TypstBridgeOutputFormat,
    TypstBridgeStatus, TYPST_BRIDGE_ABI_VERSION,
};
use crate::diagnostics;
use crate::memory::{into_raw_string, into_raw_vec};

const NOT_IMPLEMENTED_MESSAGE: &str = "Typst rendering is not implemented yet";

pub fn compile(request: *const TypstBridgeCompileRequest) -> *mut TypstBridgeCompileResult {
    let result = match validate_request(request) {
        Ok(()) => result(
            TypstBridgeStatus::Unsupported,
            NOT_IMPLEMENTED_MESSAGE,
            true,
        ),
        Err(message) => result(TypstBridgeStatus::InvalidArgument, &message, true),
    };

    Box::into_raw(Box::new(result))
}

pub fn panic_result(message: &str) -> *mut TypstBridgeCompileResult {
    Box::into_raw(Box::new(result(TypstBridgeStatus::Panic, message, true)))
}

fn result(
    status: TypstBridgeStatus,
    message: &str,
    include_diagnostic: bool,
) -> TypstBridgeCompileResult {
    let (message_utf8, message_len) = into_raw_string(message);
    let diagnostics = if include_diagnostic {
        vec![diagnostics::error(message)]
    } else {
        Vec::new()
    };
    let (diagnostics, diagnostics_count) = into_raw_vec(diagnostics);

    TypstBridgeCompileResult {
        status,
        outputs: std::ptr::null_mut(),
        outputs_count: 0,
        diagnostics,
        diagnostics_count,
        message_utf8,
        message_len,
    }
}

fn validate_request(request: *const TypstBridgeCompileRequest) -> Result<(), String> {
    if request.is_null() {
        return Err("Compile request pointer must not be null".to_owned());
    }

    let request = unsafe { &*request };

    if request.abi_version != TYPST_BRIDGE_ABI_VERSION {
        return Err(format!(
            "Unsupported ABI version {}; expected {}",
            request.abi_version, TYPST_BRIDGE_ABI_VERSION
        ));
    }

    read_utf8(request.source_utf8, request.source_len, "source")?;
    read_utf8(
        request.working_dir_utf8,
        request.working_dir_len,
        "working_dir",
    )?;
    read_utf8(
        request.root_file_name_utf8,
        request.root_file_name_len,
        "root_file_name",
    )?;

    let output_format = TypstBridgeOutputFormat::from_raw(request.output_format)
        .ok_or_else(|| format!("Unsupported output format {}", request.output_format))?;

    validate_ppi(output_format, request.ppi)?;

    if request.font_paths_count > 0 {
        if request.font_paths.is_null() {
            return Err(
                "font_paths must not be null when font_paths_count is greater than zero".to_owned(),
            );
        }

        let font_paths =
            unsafe { slice::from_raw_parts(request.font_paths, request.font_paths_count) };
        for (index, font_path) in font_paths.iter().enumerate() {
            read_utf8(
                font_path.value_utf8,
                font_path.value_len,
                &format!("font_paths[{index}].value"),
            )?;
        }
    }

    Ok(())
}

fn read_utf8<'a>(ptr: *const c_char, len: usize, name: &str) -> Result<&'a str, String> {
    if len == 0 {
        return Ok("");
    }

    if ptr.is_null() {
        return Err(format!(
            "{name}_utf8 must not be null when {name}_len is greater than zero"
        ));
    }

    let bytes = unsafe { slice::from_raw_parts(ptr.cast::<u8>(), len) };
    str::from_utf8(bytes).map_err(|_| format!("{name}_utf8 must be valid UTF-8"))
}

fn validate_ppi(output_format: TypstBridgeOutputFormat, ppi: f64) -> Result<(), String> {
    if !ppi.is_finite() {
        return Err("ppi must be finite".to_owned());
    }

    match output_format {
        TypstBridgeOutputFormat::Png if ppi <= 0.0 => {
            Err("ppi must be positive for PNG output".to_owned())
        }
        TypstBridgeOutputFormat::Pdf | TypstBridgeOutputFormat::Svg if ppi < 0.0 => {
            Err("ppi must not be negative".to_owned())
        }
        _ => Ok(()),
    }
}

#[cfg(test)]
mod tests {
    use std::ffi::CString;

    use super::*;
    use crate::abi::TypstBridgeString;
    use crate::memory::free_result;

    fn valid_request(source: &CString) -> TypstBridgeCompileRequest {
        TypstBridgeCompileRequest {
            abi_version: TYPST_BRIDGE_ABI_VERSION,
            source_utf8: source.as_ptr(),
            source_len: source.as_bytes().len(),
            working_dir_utf8: std::ptr::null(),
            working_dir_len: 0,
            root_file_name_utf8: std::ptr::null(),
            root_file_name_len: 0,
            font_paths: std::ptr::null(),
            font_paths_count: 0,
            output_format: TypstBridgeOutputFormat::Pdf as u32,
            ppi: 0.0,
            flags: 0,
        }
    }

    #[test]
    fn null_request_returns_invalid_argument() {
        let result = compile(std::ptr::null());
        assert!(!result.is_null());
        unsafe {
            assert_eq!((*result).status, TypstBridgeStatus::InvalidArgument);
            free_result(result);
        }
    }

    #[test]
    fn invalid_abi_returns_invalid_argument() {
        let source = CString::new("Hello").unwrap();
        let mut request = valid_request(&source);
        request.abi_version = 99;

        let result = compile(&request);
        unsafe {
            assert_eq!((*result).status, TypstBridgeStatus::InvalidArgument);
            free_result(result);
        }
    }

    #[test]
    fn valid_compile_returns_unsupported() {
        let source = CString::new("Hello").unwrap();
        let request = valid_request(&source);

        let result = compile(&request);
        unsafe {
            assert_eq!((*result).status, TypstBridgeStatus::Unsupported);
            assert_eq!((*result).outputs_count, 0);
            assert_eq!((*result).diagnostics_count, 1);
            free_result(result);
        }
    }

    #[test]
    fn font_paths_must_be_valid_utf8() {
        let source = CString::new("Hello").unwrap();
        let invalid_font_path = [0xff_u8];
        let font_paths = [TypstBridgeString {
            value_utf8: invalid_font_path.as_ptr().cast::<c_char>(),
            value_len: invalid_font_path.len(),
        }];
        let mut request = valid_request(&source);
        request.font_paths = font_paths.as_ptr();
        request.font_paths_count = font_paths.len();

        let result = compile(&request);
        unsafe {
            assert_eq!((*result).status, TypstBridgeStatus::InvalidArgument);
            free_result(result);
        }
    }

    #[test]
    fn font_paths_pointer_is_required_when_count_is_non_zero() {
        let source = CString::new("Hello").unwrap();
        let mut request = valid_request(&source);
        request.font_paths_count = 1;

        let result = compile(&request);
        unsafe {
            assert_eq!((*result).status, TypstBridgeStatus::InvalidArgument);
            free_result(result);
        }
    }

    #[test]
    fn repeated_free_cycles_and_null_free_are_safe() {
        unsafe { free_result(std::ptr::null_mut()) };

        for _ in 0..32 {
            let result = compile(std::ptr::null());
            unsafe { free_result(result) };
        }
    }
}
