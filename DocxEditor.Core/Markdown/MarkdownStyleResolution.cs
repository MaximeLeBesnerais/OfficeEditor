using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxEditor.Core.Markdown;

/// <summary>
/// The outcome of resolving a markdown style reference against a document's styles.
/// </summary>
public sealed record MarkdownStyleResolution(
    bool Resolved,
    string? StyleId,
    MarkdownStyleKind Kind,
    Style? FallbackStyle,
    IReadOnlyList<MarkdownStyleDiagnostic> Diagnostics)
{
    /// <summary>The "no style applied" outcome for a null/empty reference.</summary>
    public static MarkdownStyleResolution None(MarkdownStyleKind kind) =>
        new(false, null, kind, null, []);

    /// <summary>True when no style should be emitted at all.</summary>
    public bool HasStyle => StyleId is not null;
}
