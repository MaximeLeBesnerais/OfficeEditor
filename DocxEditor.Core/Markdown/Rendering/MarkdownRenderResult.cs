using DocxEditor.Core.Markdown;
using DocxEditor.Core.Markdown.Model;

namespace DocxEditor.Core.Markdown.Rendering;

/// <summary>
/// The outcome of a rich markdown conversion: the parse result plus every diagnostic produced
/// while parsing and rendering. Callers can branch on <see cref="HasErrors"/> /
/// <see cref="HasWarnings"/>; a conversion that fell back to visible text (unresolved image,
/// HTML retained verbatim) is reflected here in strict mode and never silently dropped.
/// </summary>
public sealed record MarkdownRenderResult(
    MarkdownParseResult ParseResult,
    IReadOnlyList<MarkdownDiagnostic> Diagnostics)
{
    public bool HasErrors => Diagnostics.Any(d => d.Severity == MarkdownDiagnosticSeverity.Error);

    public bool HasWarnings => Diagnostics.Any(d => d.Severity == MarkdownDiagnosticSeverity.Warning);
}
