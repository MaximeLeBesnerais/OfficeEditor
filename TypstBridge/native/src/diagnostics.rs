use std::ptr;

use typst::diag::{Severity, SourceDiagnostic};
use typst::{World, WorldExt};

use crate::abi::{TypstBridgeDiagnostic, TYPST_BRIDGE_DIAG_ERROR, TYPST_BRIDGE_DIAG_WARNING};
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

pub fn from_typst(world: &dyn World, diagnostic: &SourceDiagnostic) -> TypstBridgeDiagnostic {
    let severity = match diagnostic.severity {
        Severity::Error => TYPST_BRIDGE_DIAG_ERROR,
        Severity::Warning => TYPST_BRIDGE_DIAG_WARNING,
    };
    let message = diagnostic.message.to_string();
    let (message_utf8, message_len) = into_raw_string(&message);

    let (file_utf8, file_len, line, column) = diagnostic
        .span
        .id()
        .and_then(|id| {
            let file = id.vpath().get_without_slash().to_string();
            let source = world.source(id).ok()?;
            let start = world.range(diagnostic.span)?.start;
            let (line, column) = source.lines().byte_to_line_column(start)?;
            Some((file, line as u32 + 1, column as u32 + 1))
        })
        .map(|(file, line, column)| {
            let (file_utf8, file_len) = into_raw_string(&file);
            (file_utf8, file_len, line, column)
        })
        .unwrap_or((ptr::null_mut(), 0, 0, 0));

    TypstBridgeDiagnostic {
        severity,
        message_utf8,
        message_len,
        file_utf8,
        file_len,
        line,
        column,
    }
}

pub fn from_typst_many(
    world: &dyn World,
    diagnostics: impl IntoIterator<Item = SourceDiagnostic>,
) -> Vec<TypstBridgeDiagnostic> {
    diagnostics
        .into_iter()
        .map(|diagnostic| from_typst(world, &diagnostic))
        .collect()
}
