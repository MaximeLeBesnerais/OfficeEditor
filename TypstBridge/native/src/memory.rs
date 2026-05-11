use std::ffi::{c_char, CString};
use std::ptr;

use crate::abi::{TypstBridgeCompileResult, TypstBridgeDiagnostic, TypstBridgeOutputItem};

pub fn into_raw_string(value: &str) -> (*mut c_char, usize) {
    let sanitized = value.replace('\0', "�");
    let len = sanitized.len();
    let ptr = CString::new(sanitized)
        .expect("interior NULs were removed")
        .into_raw();

    (ptr, len)
}

pub fn into_raw_vec<T>(items: Vec<T>) -> (*mut T, usize) {
    if items.is_empty() {
        return (ptr::null_mut(), 0);
    }

    let count = items.len();
    let boxed = items.into_boxed_slice();
    (Box::into_raw(boxed) as *mut T, count)
}

pub unsafe fn free_string(value: *mut c_char) {
    if !value.is_null() {
        drop(CString::from_raw(value));
    }
}

pub unsafe fn free_result(result: *mut TypstBridgeCompileResult) {
    if result.is_null() {
        return;
    }

    let result = Box::from_raw(result);

    if !result.outputs.is_null() {
        let outputs = Box::from_raw(ptr::slice_from_raw_parts_mut(
            result.outputs,
            result.outputs_count,
        ));
        for output in outputs {
            free_string(output.file_name_utf8);
            if !output.data.is_null() {
                drop(Box::from_raw(ptr::slice_from_raw_parts_mut(
                    output.data,
                    output.data_len,
                )));
            }
        }
    }

    if !result.diagnostics.is_null() {
        let diagnostics = Box::from_raw(ptr::slice_from_raw_parts_mut(
            result.diagnostics,
            result.diagnostics_count,
        ));
        for diagnostic in diagnostics {
            free_diagnostic(diagnostic);
        }
    }

    free_string(result.message_utf8);
}

unsafe fn free_diagnostic(diagnostic: TypstBridgeDiagnostic) {
    free_string(diagnostic.message_utf8);
    free_string(diagnostic.file_utf8);
}

#[allow(dead_code)]
pub fn into_output_data(data: Vec<u8>) -> (*mut u8, usize) {
    if data.is_empty() {
        return (ptr::null_mut(), 0);
    }

    let len = data.len();
    let boxed = data.into_boxed_slice();
    (Box::into_raw(boxed) as *mut u8, len)
}

#[allow(dead_code)]
pub fn output_item(page_index: u32, file_name: &str, data: Vec<u8>) -> TypstBridgeOutputItem {
    let (file_name_utf8, file_name_len) = into_raw_string(file_name);
    let (data, data_len) = into_output_data(data);

    TypstBridgeOutputItem {
        page_index,
        file_name_utf8,
        file_name_len,
        data,
        data_len,
    }
}

#[cfg(test)]
mod tests {
    use crate::abi::{TypstBridgeCompileResult, TypstBridgeStatus};
    use crate::memory::{free_result, into_raw_vec, output_item};

    #[test]
    fn frees_output_data_from_vec_with_extra_capacity() {
        let mut data = Vec::with_capacity(1024);
        data.extend_from_slice(&[1, 2, 3, 4]);
        assert!(data.capacity() > data.len());

        let outputs = vec![output_item(0, "page.png", data)];
        let (outputs, outputs_count) = into_raw_vec(outputs);
        let result = Box::into_raw(Box::new(TypstBridgeCompileResult {
            status: TypstBridgeStatus::Ok,
            outputs,
            outputs_count,
            diagnostics: std::ptr::null_mut(),
            diagnostics_count: 0,
            message_utf8: std::ptr::null_mut(),
            message_len: 0,
        }));

        unsafe { free_result(result) };
    }
}
