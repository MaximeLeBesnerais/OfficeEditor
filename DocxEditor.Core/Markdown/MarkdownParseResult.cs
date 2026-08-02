using DocxEditor.Core.Markdown.Model;

namespace DocxEditor.Core.Markdown;

/// <summary>
/// The outcome of a rich markdown parse: the document IR plus any diagnostics
/// produced while parsing.
/// </summary>
public sealed record MarkdownParseResult(MarkdownDocument Document, IReadOnlyList<MarkdownDiagnostic> Diagnostics)
{
    public bool HasErrors => Diagnostics.Any(d => d.Severity == MarkdownDiagnosticSeverity.Error);

    public bool HasWarnings => Diagnostics.Any(d => d.Severity == MarkdownDiagnosticSeverity.Warning);
}
