use std::ffi::c_char;
use std::path::PathBuf;
use std::{slice, str};

use typst_layout::PagedDocument;

use crate::abi::{
    TypstBridgeCompileRequest, TypstBridgeCompileResult, TypstBridgeOutputFormat,
    TypstBridgeStatus, TYPST_BRIDGE_ABI_VERSION,
};
use crate::diagnostics;
use crate::evict;
use crate::fonts;
use crate::memory::{into_raw_string, into_raw_vec, output_item};
use crate::render_pdf;
use crate::render_png;
use crate::render_svg;
use crate::world::BridgeWorld;

pub fn compile(request: *const TypstBridgeCompileRequest) -> *mut TypstBridgeCompileResult {
    let result = match validate_request(request) {
        Ok(request) => compile_valid(request),
        Err(message) => result(TypstBridgeStatus::InvalidArgument, &message, true),
    };
    evict::note_compile_finished();

    Box::into_raw(Box::new(result))
}

pub fn panic_result(message: &str) -> *mut TypstBridgeCompileResult {
    Box::into_raw(Box::new(result(TypstBridgeStatus::Panic, message, true)))
}

/// Builds an error compile result for failures that happen outside the normal
/// request-validation path (e.g. invalid session compile arguments).
pub(crate) fn error_result(status: TypstBridgeStatus, message: &str) -> TypstBridgeCompileResult {
    result(status, message, true)
}

fn compile_valid(request: ValidRequest) -> TypstBridgeCompileResult {
    let fonts = match fonts::load_font_paths(&request.font_paths, &request.working_dir) {
        Ok(fonts) => fonts,
        Err(message) => return result(TypstBridgeStatus::InvalidArgument, &message, true),
    };

    let world = BridgeWorld::new(
        request.source,
        request.working_dir,
        &request.root_file_name,
        fonts,
    );

    compile_world(
        &world,
        &request.root_file_name,
        request.output_format,
        request.ppi,
    )
}

/// Compiles an already-constructed world and renders the requested outputs.
/// Shared by the single-shot entry point and persistent sessions.
pub(crate) fn compile_world(
    world: &BridgeWorld,
    root_file_name: &str,
    output_format: TypstBridgeOutputFormat,
    ppi: f64,
) -> TypstBridgeCompileResult {
    let compiled = typst::compile::<PagedDocument>(world);
    match compiled.output {
        Ok(document) => {
            let diagnostics = diagnostics::from_typst_many(world, compiled.warnings);
            match output_format {
                TypstBridgeOutputFormat::Pdf => {
                    render_pdf_result(world, root_file_name, &document, diagnostics)
                }
                TypstBridgeOutputFormat::Svg => render_svg_result(&document, diagnostics),
                TypstBridgeOutputFormat::Png => render_png_result(&document, ppi, diagnostics),
            }
        }
        Err(errors) => {
            let diagnostics =
                diagnostics::from_typst_many(world, errors.into_iter().chain(compiled.warnings));
            let message = diagnostics_message(&diagnostics, "Typst compilation failed");
            result_with_diagnostics(TypstBridgeStatus::Compile, &message, diagnostics)
        }
    }
}

fn render_svg_result(
    document: &PagedDocument,
    diagnostics: Vec<crate::abi::TypstBridgeDiagnostic>,
) -> TypstBridgeCompileResult {
    let outputs = render_svg::render(document)
        .into_iter()
        .map(|(page_index, data)| output_item(page_index, &svg_file_name(page_index), data))
        .collect();

    result_with_outputs(TypstBridgeStatus::Ok, "", outputs, diagnostics)
}

fn svg_file_name(page_index: u32) -> String {
    format!("page-{:03}.svg", page_index + 1)
}

fn render_png_result(
    document: &PagedDocument,
    ppi: f64,
    diagnostics: Vec<crate::abi::TypstBridgeDiagnostic>,
) -> TypstBridgeCompileResult {
    match render_png::render(document, ppi) {
        Ok(outputs) => {
            let outputs = outputs
                .into_iter()
                .map(|(page_index, data)| output_item(page_index, &png_file_name(page_index), data))
                .collect();

            result_with_outputs(TypstBridgeStatus::Ok, "", outputs, diagnostics)
        }
        Err(message) => {
            let mut diagnostics = diagnostics;
            diagnostics.push(diagnostics::error(&message));
            result_with_diagnostics(TypstBridgeStatus::Render, &message, diagnostics)
        }
    }
}

fn png_file_name(page_index: u32) -> String {
    format!("page-{:03}.png", page_index + 1)
}

fn render_pdf_result(
    world: &BridgeWorld,
    root_file_name: &str,
    document: &PagedDocument,
    diagnostics: Vec<crate::abi::TypstBridgeDiagnostic>,
) -> TypstBridgeCompileResult {
    match render_pdf::render(document) {
        Ok(data) => result_with_outputs(
            TypstBridgeStatus::Ok,
            "",
            vec![output_item(0, &pdf_file_name(root_file_name), data)],
            diagnostics,
        ),
        Err(errors) => {
            let mut diagnostics = diagnostics;
            diagnostics.extend(diagnostics::from_typst_many(world, errors));
            let message = diagnostics_message(&diagnostics, "Typst PDF export failed");
            result_with_diagnostics(TypstBridgeStatus::Render, &message, diagnostics)
        }
    }
}

fn pdf_file_name(root_file_name: &str) -> String {
    let trimmed = root_file_name.trim();
    if trimmed.is_empty() {
        return "output.pdf".to_owned();
    }

    let path = std::path::Path::new(trimmed);
    let stem = path
        .file_stem()
        .and_then(|stem| stem.to_str())
        .filter(|stem| !stem.is_empty())
        .unwrap_or("output");

    format!("{stem}.pdf")
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

fn result_with_diagnostics(
    status: TypstBridgeStatus,
    message: &str,
    diagnostics: Vec<crate::abi::TypstBridgeDiagnostic>,
) -> TypstBridgeCompileResult {
    let (message_utf8, message_len) = into_raw_string(message);
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

fn result_with_outputs(
    status: TypstBridgeStatus,
    message: &str,
    outputs: Vec<crate::abi::TypstBridgeOutputItem>,
    diagnostics: Vec<crate::abi::TypstBridgeDiagnostic>,
) -> TypstBridgeCompileResult {
    let (message_utf8, message_len) = into_raw_string(message);
    let (outputs, outputs_count) = into_raw_vec(outputs);
    let (diagnostics, diagnostics_count) = into_raw_vec(diagnostics);

    TypstBridgeCompileResult {
        status,
        outputs,
        outputs_count,
        diagnostics,
        diagnostics_count,
        message_utf8,
        message_len,
    }
}

fn diagnostics_message(
    diagnostics: &[crate::abi::TypstBridgeDiagnostic],
    fallback: &str,
) -> String {
    if diagnostics.is_empty() {
        return fallback.to_owned();
    }

    let first = &diagnostics[0];
    if first.message_utf8.is_null() || first.message_len == 0 {
        return fallback.to_owned();
    }

    let message =
        unsafe { slice::from_raw_parts(first.message_utf8.cast::<u8>(), first.message_len) };
    match str::from_utf8(message) {
        Ok(message) => format!("{fallback}: {message}"),
        Err(_) => fallback.to_owned(),
    }
}

struct ValidRequest {
    source: String,
    working_dir: PathBuf,
    root_file_name: String,
    font_paths: Vec<String>,
    output_format: TypstBridgeOutputFormat,
    ppi: f64,
}

fn validate_request(request: *const TypstBridgeCompileRequest) -> Result<ValidRequest, String> {
    if request.is_null() {
        return Err("Compile request pointer must not be null".to_owned());
    }

    let request = unsafe { &*request };

    // The v2 single-shot entry point remains supported: the request layout is
    // unchanged between ABI v2 and v3, so both versions are accepted here.
    if request.abi_version != TYPST_BRIDGE_ABI_VERSION && request.abi_version != 2 {
        return Err(format!(
            "Unsupported ABI version {}; expected {}",
            request.abi_version, TYPST_BRIDGE_ABI_VERSION
        ));
    }

    let source = read_utf8(request.source_utf8, request.source_len, "source")?.to_owned();
    let working_dir = read_utf8(
        request.working_dir_utf8,
        request.working_dir_len,
        "working_dir",
    )?;
    let root_file_name = read_utf8(
        request.root_file_name_utf8,
        request.root_file_name_len,
        "root_file_name",
    )?
    .to_owned();

    let output_format = TypstBridgeOutputFormat::from_raw(request.output_format)
        .ok_or_else(|| format!("Unsupported output format {}", request.output_format))?;

    validate_ppi(output_format, request.ppi)?;

    let font_paths = read_font_paths(request.font_paths, request.font_paths_count)?;
    let working_dir = resolve_working_dir(working_dir)?;

    Ok(ValidRequest {
        source,
        working_dir,
        root_file_name,
        font_paths,
        output_format,
        ppi: request.ppi,
    })
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

/// Reads and validates the raw font-path array from ABI string pairs.
/// Shared by request validation and session creation.
pub(crate) fn read_font_paths(
    font_paths: *const crate::abi::TypstBridgeString,
    font_paths_count: usize,
) -> Result<Vec<String>, String> {
    let mut paths = Vec::new();
    if font_paths_count == 0 {
        return Ok(paths);
    }

    if font_paths.is_null() {
        return Err(
            "font_paths must not be null when font_paths_count is greater than zero".to_owned(),
        );
    }

    let raw_font_paths = unsafe { slice::from_raw_parts(font_paths, font_paths_count) };
    for (index, font_path) in raw_font_paths.iter().enumerate() {
        paths.push(
            read_utf8(
                font_path.value_utf8,
                font_path.value_len,
                &format!("font_paths[{index}].value"),
            )?
            .to_owned(),
        );
    }

    Ok(paths)
}

/// Resolves an optional working directory the same way request validation
/// does: an empty value falls back to the process current directory.
pub(crate) fn resolve_working_dir(working_dir: &str) -> Result<PathBuf, String> {
    if working_dir.is_empty() {
        std::env::current_dir().map_err(|error| format!("Failed to get current directory: {error}"))
    } else {
        Ok(PathBuf::from(working_dir))
    }
}

/// Validates a raw ABI string pair without copying. Shared with the session
/// entry points.
pub(crate) fn read_utf8_str<'a>(
    ptr: *const c_char,
    len: usize,
    name: &str,
) -> Result<&'a str, String> {
    read_utf8(ptr, len, name)
}

pub(crate) fn validate_ppi(
    output_format: TypstBridgeOutputFormat,
    ppi: f64,
) -> Result<(), String> {
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
    use std::fs;
    use std::time::{SystemTime, UNIX_EPOCH};

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

    fn temp_dir(name: &str) -> std::path::PathBuf {
        let unique = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let dir = std::env::temp_dir().join(format!("typst-bridge-{name}-{unique}"));
        fs::create_dir(&dir).unwrap();
        dir
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

    unsafe fn assert_pdf_result(result: *mut TypstBridgeCompileResult) {
        assert_eq!((*result).status, TypstBridgeStatus::Ok);
        assert_eq!((*result).outputs_count, 1);

        let output = &*(*result).outputs;
        assert_eq!(output.page_index, 0);
        assert_eq!(output.file_name_len, "output.pdf".len());
        let file_name =
            slice::from_raw_parts(output.file_name_utf8.cast::<u8>(), output.file_name_len);
        assert_eq!(str::from_utf8(file_name).unwrap(), "output.pdf");
        assert!(output.data_len > 4);

        let data = slice::from_raw_parts(output.data, output.data_len);
        assert_eq!(&data[..4], b"%PDF");
    }

    #[test]
    fn minimal_pdf_returns_one_pdf_output() {
        let source = CString::new("Hello").unwrap();
        let request = valid_request(&source);

        let result = compile(&request);
        unsafe {
            assert_pdf_result(result);
            free_result(result);
        }
    }

    #[test]
    fn pdf_output_file_name_uses_root_file_stem() {
        let source = CString::new("Hello").unwrap();
        let root_file_name = CString::new("deck.typ").unwrap();
        let mut request = valid_request(&source);
        request.root_file_name_utf8 = root_file_name.as_ptr();
        request.root_file_name_len = root_file_name.as_bytes().len();

        let result = compile(&request);
        unsafe {
            assert_eq!((*result).status, TypstBridgeStatus::Ok);
            assert_eq!((*result).outputs_count, 1);

            let output = &*(*result).outputs;
            let file_name =
                slice::from_raw_parts(output.file_name_utf8.cast::<u8>(), output.file_name_len);
            assert_eq!(str::from_utf8(file_name).unwrap(), "deck.pdf");
            free_result(result);
        }
    }

    #[test]
    fn multi_page_pdf_returns_one_pdf_output() {
        let source = CString::new("First page\n#pagebreak()\nSecond page").unwrap();
        let request = valid_request(&source);

        let result = compile(&request);
        unsafe {
            assert_pdf_result(result);
            free_result(result);
        }
    }

    #[test]
    fn single_page_svg_returns_one_svg_output() {
        let source = CString::new("Hello").unwrap();
        let mut request = valid_request(&source);
        request.output_format = TypstBridgeOutputFormat::Svg as u32;

        let result = compile(&request);
        unsafe {
            assert_eq!((*result).status, TypstBridgeStatus::Ok);
            assert_eq!((*result).outputs_count, 1);

            let output = &*(*result).outputs;
            assert_eq!(output.page_index, 0);

            let file_name =
                slice::from_raw_parts(output.file_name_utf8.cast::<u8>(), output.file_name_len);
            assert_eq!(str::from_utf8(file_name).unwrap(), "page-001.svg");

            let data = slice::from_raw_parts(output.data, output.data_len);
            let svg = str::from_utf8(data).unwrap();
            assert!(svg.contains("<svg") || svg.contains("<?xml"));

            free_result(result);
        }
    }

    #[test]
    fn multi_page_svg_returns_outputs_in_page_order() {
        let source = CString::new("First page\n#pagebreak()\nSecond page").unwrap();
        let mut request = valid_request(&source);
        request.output_format = TypstBridgeOutputFormat::Svg as u32;

        let result = compile(&request);
        unsafe {
            assert_eq!((*result).status, TypstBridgeStatus::Ok);
            assert_eq!((*result).outputs_count, 2);

            let outputs = slice::from_raw_parts((*result).outputs, (*result).outputs_count);
            for (index, output) in outputs.iter().enumerate() {
                assert_eq!(output.page_index, index as u32);
                let file_name =
                    slice::from_raw_parts(output.file_name_utf8.cast::<u8>(), output.file_name_len);
                assert_eq!(
                    str::from_utf8(file_name).unwrap(),
                    format!("page-{:03}.svg", index + 1)
                );
            }

            free_result(result);
        }
    }

    #[test]
    fn invalid_typst_returns_compile() {
        let source = CString::new("#let =").unwrap();
        let request = valid_request(&source);

        let result = compile(&request);
        unsafe {
            assert_eq!((*result).status, TypstBridgeStatus::Compile);
            assert!((*result).diagnostics_count >= 1);
            free_result(result);
        }
    }

    #[test]
    fn asset_resolution_is_relative_to_working_dir() {
        let dir = temp_dir("asset");
        fs::write(dir.join("data.txt"), "from asset").unwrap();

        let source = CString::new("#read(\"data.txt\")").unwrap();
        let working_dir = CString::new(dir.to_string_lossy().as_bytes()).unwrap();
        let mut request = valid_request(&source);
        request.working_dir_utf8 = working_dir.as_ptr();
        request.working_dir_len = working_dir.as_bytes().len();

        let result = compile(&request);
        unsafe {
            assert_pdf_result(result);
            free_result(result);
        }

        fs::remove_dir_all(dir).unwrap();
    }

    #[test]
    fn single_page_png_returns_one_png_output() {
        let source = CString::new("Hello").unwrap();
        let mut request = valid_request(&source);
        request.output_format = TypstBridgeOutputFormat::Png as u32;
        request.ppi = 96.0;

        let result = compile(&request);
        unsafe {
            assert_eq!((*result).status, TypstBridgeStatus::Ok);
            assert_eq!((*result).outputs_count, 1);

            let output = &*(*result).outputs;
            assert_eq!(output.page_index, 0);

            let file_name =
                slice::from_raw_parts(output.file_name_utf8.cast::<u8>(), output.file_name_len);
            assert_eq!(str::from_utf8(file_name).unwrap(), "page-001.png");

            let data = slice::from_raw_parts(output.data, output.data_len);
            assert_eq!(&data[..8], b"\x89PNG\r\n\x1a\n");
            free_result(result);
        }
    }

    #[test]
    fn multi_page_png_returns_outputs_in_page_order() {
        let source = CString::new("First page\n#pagebreak()\nSecond page").unwrap();
        let mut request = valid_request(&source);
        request.output_format = TypstBridgeOutputFormat::Png as u32;
        request.ppi = 96.0;

        let result = compile(&request);
        unsafe {
            assert_eq!((*result).status, TypstBridgeStatus::Ok);
            assert_eq!((*result).outputs_count, 2);

            let outputs = slice::from_raw_parts((*result).outputs, (*result).outputs_count);
            for (index, output) in outputs.iter().enumerate() {
                assert_eq!(output.page_index, index as u32);
                let file_name =
                    slice::from_raw_parts(output.file_name_utf8.cast::<u8>(), output.file_name_len);
                assert_eq!(
                    str::from_utf8(file_name).unwrap(),
                    format!("page-{:03}.png", index + 1)
                );
            }

            free_result(result);
        }
    }

    #[test]
    fn invalid_zero_png_ppi_returns_invalid_argument() {
        let source = CString::new("Hello").unwrap();
        let mut request = valid_request(&source);
        request.output_format = TypstBridgeOutputFormat::Png as u32;
        request.ppi = 0.0;

        let result = compile(&request);
        unsafe {
            assert_eq!((*result).status, TypstBridgeStatus::InvalidArgument);
            free_result(result);
        }
    }

    #[test]
    fn png_ppi_changes_output_dimensions() {
        let source = CString::new("#set page(width: 1in, height: 1in)\nHello").unwrap();
        let mut request = valid_request(&source);
        request.output_format = TypstBridgeOutputFormat::Png as u32;

        request.ppi = 72.0;
        let low_ppi = compile(&request);
        request.ppi = 144.0;
        let high_ppi = compile(&request);

        unsafe {
            assert_eq!((*low_ppi).status, TypstBridgeStatus::Ok);
            assert_eq!((*high_ppi).status, TypstBridgeStatus::Ok);

            let low_output = &*(*low_ppi).outputs;
            let high_output = &*(*high_ppi).outputs;
            let low_data = slice::from_raw_parts(low_output.data, low_output.data_len);
            let high_data = slice::from_raw_parts(high_output.data, high_output.data_len);

            assert_eq!(png_dimensions(low_data), (72, 72));
            assert_eq!(png_dimensions(high_data), (144, 144));
            assert_ne!(low_data, high_data);

            free_result(low_ppi);
            free_result(high_ppi);
        }
    }

    fn png_dimensions(data: &[u8]) -> (u32, u32) {
        assert_eq!(&data[..8], b"\x89PNG\r\n\x1a\n");
        let width = u32::from_be_bytes(data[16..20].try_into().unwrap());
        let height = u32::from_be_bytes(data[20..24].try_into().unwrap());
        (width, height)
    }

    #[test]
    fn invalid_font_path_returns_invalid_argument() {
        let source = CString::new("Hello").unwrap();
        let missing_font = CString::new("/definitely/missing/font.ttf").unwrap();
        let font_paths = [TypstBridgeString {
            value_utf8: missing_font.as_ptr(),
            value_len: missing_font.as_bytes().len(),
        }];
        let mut request = valid_request(&source);
        request.font_paths = font_paths.as_ptr();
        request.font_paths_count = font_paths.len();

        let result = compile(&request);
        unsafe {
            assert_eq!((*result).status, TypstBridgeStatus::InvalidArgument);
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

    #[test]
    fn repeated_pdf_compile_free_cycles_are_safe() {
        let source = CString::new("Hello").unwrap();

        for _ in 0..8 {
            let request = valid_request(&source);
            let result = compile(&request);
            unsafe {
                assert_pdf_result(result);
                free_result(result);
            }
        }
    }
}
