//! Persistent warm-world compile sessions (ABI v3).
//!
//! A session keeps the `BridgeWorld` — library, font set, working-directory
//! resolution — alive across compiles, so comemo-memoized layout work survives
//! an edited-slide recompile. The host updates the main source in place with
//! `typst_bridge_session_update_source` and recompiles with
//! `typst_bridge_session_compile`.
//!
//! Thread-safety: each session serializes its own state behind a `Mutex`, so a
//! single session handle may be used from multiple threads (calls are
//! serialized), and distinct sessions compile fully independently. Errors are
//! reported through the same `thread_local` last-error string as the rest of
//! the ABI, which is safe under concurrent sessions.

use std::ffi::c_char;
use std::sync::Mutex;

use crate::abi::{
    TypstBridgeCompileResult, TypstBridgeOutputFormat, TypstBridgeStatus, TypstBridgeString,
};
use crate::compiler;
use crate::evict;
use crate::fonts;
use crate::world::BridgeWorld;

pub struct BridgeSession {
    state: Mutex<SessionState>,
}

struct SessionState {
    world: BridgeWorld,
    root_file_name: String,
    has_source: bool,
}

impl BridgeSession {
    fn lock_state(&self) -> std::sync::MutexGuard<'_, SessionState> {
        self.state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
    }
}

pub fn create(
    working_dir_utf8: *const c_char,
    working_dir_len: usize,
    font_paths: *const TypstBridgeString,
    font_paths_count: usize,
    out_session: *mut *mut BridgeSession,
) -> TypstBridgeStatus {
    if out_session.is_null() {
        crate::set_last_error("out_session pointer must not be null");
        return TypstBridgeStatus::InvalidArgument;
    }
    unsafe { *out_session = std::ptr::null_mut() };

    match create_inner(working_dir_utf8, working_dir_len, font_paths, font_paths_count) {
        Ok(session) => {
            unsafe { *out_session = Box::into_raw(Box::new(session)) };
            TypstBridgeStatus::Ok
        }
        Err(message) => {
            crate::set_last_error(&message);
            TypstBridgeStatus::InvalidArgument
        }
    }
}

fn create_inner(
    working_dir_utf8: *const c_char,
    working_dir_len: usize,
    font_paths: *const TypstBridgeString,
    font_paths_count: usize,
) -> Result<BridgeSession, String> {
    let working_dir = compiler::read_utf8_str(working_dir_utf8, working_dir_len, "working_dir")?;
    let working_dir = compiler::resolve_working_dir(working_dir)?;
    let font_paths = compiler::read_font_paths(font_paths, font_paths_count)?;
    let fonts = fonts::load_font_paths(&font_paths, &working_dir)?;

    Ok(BridgeSession {
        state: Mutex::new(SessionState {
            world: BridgeWorld::new(String::new(), working_dir, "main.typ", fonts),
            root_file_name: "main.typ".to_owned(),
            has_source: false,
        }),
    })
}

pub fn update_source(
    session: *mut BridgeSession,
    source_utf8: *const c_char,
    source_len: usize,
    root_file_name_utf8: *const c_char,
    root_file_name_len: usize,
) -> TypstBridgeStatus {
    if session.is_null() {
        crate::set_last_error("session pointer must not be null");
        return TypstBridgeStatus::InvalidArgument;
    }

    let source = match compiler::read_utf8_str(source_utf8, source_len, "source") {
        Ok(source) => source,
        Err(message) => {
            crate::set_last_error(&message);
            return TypstBridgeStatus::InvalidArgument;
        }
    };
    let root_file_name =
        match compiler::read_utf8_str(root_file_name_utf8, root_file_name_len, "root_file_name") {
            Ok(root_file_name) => root_file_name,
            Err(message) => {
                crate::set_last_error(&message);
                return TypstBridgeStatus::InvalidArgument;
            }
        };

    let session = unsafe { &*session };
    let mut state = session.lock_state();
    state.world.set_source(source.to_owned(), root_file_name);
    state.root_file_name = root_file_name.to_owned();
    state.has_source = true;
    TypstBridgeStatus::Ok
}

pub fn compile(
    session: *mut BridgeSession,
    output_format: u32,
    ppi: f64,
) -> *mut TypstBridgeCompileResult {
    let result = compile_inner(session, output_format, ppi);
    evict::note_compile_finished();
    Box::into_raw(Box::new(result))
}

fn compile_inner(
    session: *mut BridgeSession,
    output_format: u32,
    ppi: f64,
) -> TypstBridgeCompileResult {
    if session.is_null() {
        return compiler::error_result(
            TypstBridgeStatus::InvalidArgument,
            "session pointer must not be null",
        );
    }

    let output_format = match TypstBridgeOutputFormat::from_raw(output_format) {
        Some(format) => format,
        None => {
            return compiler::error_result(
                TypstBridgeStatus::InvalidArgument,
                &format!("Unsupported output format {output_format}"),
            );
        }
    };

    if let Err(message) = compiler::validate_ppi(output_format, ppi) {
        return compiler::error_result(TypstBridgeStatus::InvalidArgument, &message);
    }

    let session = unsafe { &*session };
    let state = session.lock_state();
    if !state.has_source {
        return compiler::error_result(
            TypstBridgeStatus::InvalidArgument,
            "session has no source; call typst_bridge_session_update_source first",
        );
    }

    compiler::compile_world(&state.world, &state.root_file_name, output_format, ppi)
}

pub fn free(session: *mut BridgeSession) {
    if session.is_null() {
        return;
    }
    unsafe { drop(Box::from_raw(session)) };
}

#[cfg(test)]
mod tests {
    use std::ffi::CString;
    use std::slice;
    use std::str;

    use super::*;
    use crate::memory::free_result;

    fn create_session() -> *mut BridgeSession {
        let mut session: *mut BridgeSession = std::ptr::null_mut();
        let status = create(
            std::ptr::null(),
            0,
            std::ptr::null(),
            0,
            &mut session,
        );
        assert_eq!(status, TypstBridgeStatus::Ok);
        assert!(!session.is_null());
        session
    }

    fn update(session: *mut BridgeSession, source: &CString) {
        let status = update_source(
            session,
            source.as_ptr(),
            source.as_bytes().len(),
            std::ptr::null(),
            0,
        );
        assert_eq!(status, TypstBridgeStatus::Ok);
    }

    unsafe fn pdf_bytes(result: *mut TypstBridgeCompileResult) -> Vec<u8> {
        assert_eq!((*result).status, TypstBridgeStatus::Ok);
        assert_eq!((*result).outputs_count, 1);
        let output = &*(*result).outputs;
        assert!(output.data_len > 4);
        slice::from_raw_parts(output.data, output.data_len).to_vec()
    }

    #[test]
    fn session_create_update_compile_free_round_trip() {
        let session = create_session();
        let source = CString::new("Hello from a session").unwrap();
        update(session, &source);

        let result = compile(session, TypstBridgeOutputFormat::Pdf as u32, 0.0);
        unsafe {
            let bytes = pdf_bytes(result);
            assert_eq!(&bytes[..4], b"%PDF");
            free_result(result);
        }

        free(session);
        free(std::ptr::null_mut());
    }

    #[test]
    fn session_compile_before_update_source_is_invalid_argument() {
        let session = create_session();
        let result = compile(session, TypstBridgeOutputFormat::Pdf as u32, 0.0);
        unsafe {
            assert_eq!((*result).status, TypstBridgeStatus::InvalidArgument);
            free_result(result);
        }
        free(session);
    }

    #[test]
    fn session_recompile_of_unchanged_source_is_deterministic() {
        let session = create_session();
        let source = CString::new("Deterministic session output").unwrap();
        update(session, &source);

        let first = compile(session, TypstBridgeOutputFormat::Pdf as u32, 0.0);
        let second = compile(session, TypstBridgeOutputFormat::Pdf as u32, 0.0);
        unsafe {
            assert_eq!(pdf_bytes(first), pdf_bytes(second));
            free_result(first);
            free_result(second);
        }
        free(session);
    }

    #[test]
    fn session_edited_source_recompiles_and_matches_cold_compile() {
        let session = create_session();
        let original = CString::new("Original page").unwrap();
        update(session, &original);
        let first = compile(session, TypstBridgeOutputFormat::Pdf as u32, 0.0);

        let edited = CString::new("Edited page\n#pagebreak()\nSecond page").unwrap();
        update(session, &edited);
        let second = compile(session, TypstBridgeOutputFormat::Pdf as u32, 0.0);

        // The session compile of the edited source must match a cold
        // single-shot compile of the same source byte-for-byte.
        let cold_request_source = edited.clone();
        let request = crate::abi::TypstBridgeCompileRequest {
            abi_version: crate::abi::TYPST_BRIDGE_ABI_VERSION,
            source_utf8: cold_request_source.as_ptr(),
            source_len: cold_request_source.as_bytes().len(),
            working_dir_utf8: std::ptr::null(),
            working_dir_len: 0,
            root_file_name_utf8: std::ptr::null(),
            root_file_name_len: 0,
            font_paths: std::ptr::null(),
            font_paths_count: 0,
            output_format: TypstBridgeOutputFormat::Pdf as u32,
            ppi: 0.0,
            flags: 0,
        };
        let cold = compiler::compile(&request);

        unsafe {
            let session_bytes = pdf_bytes(second);
            assert_ne!(pdf_bytes(first), session_bytes);
            assert_eq!(session_bytes, pdf_bytes(cold));
            free_result(first);
            free_result(second);
            free_result(cold);
        }
        free(session);
    }

    #[test]
    fn compiles_across_eviction_boundaries_stay_deterministic() {
        let source = CString::new("Eviction boundary determinism").unwrap();
        let session = create_session();
        update(session, &source);

        let before = compile(session, TypstBridgeOutputFormat::Pdf as u32, 0.0);
        evict::evict(0); // force a full eviction pass between compiles
        let after = compile(session, TypstBridgeOutputFormat::Pdf as u32, 0.0);

        unsafe {
            assert_eq!(pdf_bytes(before), pdf_bytes(after));
            free_result(before);
            free_result(after);
        }
        free(session);
    }

    #[test]
    fn null_session_compile_returns_invalid_argument() {
        let result = compile(std::ptr::null_mut(), TypstBridgeOutputFormat::Pdf as u32, 0.0);
        unsafe {
            assert_eq!((*result).status, TypstBridgeStatus::InvalidArgument);
            free_result(result);
        }
    }
}
