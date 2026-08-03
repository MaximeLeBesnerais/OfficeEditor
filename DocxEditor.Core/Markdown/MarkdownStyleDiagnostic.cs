using DocxEditor.Core.Markdown.Model;

namespace DocxEditor.Core.Markdown;

/// <summary>
/// A diagnostic produced while resolving a markdown style reference against the
/// document's styles. <see cref="Path"/> locates the emitting element (for example
/// <c>blocks[3].inlines[1]</c>) and <see cref="Element"/> carries the semantic key
/// (for example <c>codeInline</c>), mirroring the shape of the rich markdown
/// <see cref="MarkdownDiagnostic"/> records so both can be surfaced together.
/// </summary>
public sealed record MarkdownStyleDiagnostic(
    MarkdownDiagnosticSeverity Severity,
    string Message,
    string? Element = null,
    string? Path = null)
{
    /// <summary>Converts this diagnostic to the parser-level diagnostic shape.</summary>
    public MarkdownDiagnostic ToMarkdownDiagnostic() =>
        new(Severity, Message, null, Element);
}
