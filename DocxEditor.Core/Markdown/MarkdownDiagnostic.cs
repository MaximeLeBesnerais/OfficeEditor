using DocxEditor.Core.Markdown.Model;

namespace DocxEditor.Core.Markdown;

/// <summary>
/// A diagnostic produced while parsing markdown. Diagnostics are emitted for
/// constructs that were retained verbatim rather than fully structured, and for
/// anything that could not be represented.
/// </summary>
public sealed record MarkdownDiagnostic(
    MarkdownDiagnosticSeverity Severity,
    string Message,
    SourceSpan? Span = null,
    string? NodeKind = null);
