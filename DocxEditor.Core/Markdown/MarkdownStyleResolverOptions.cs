namespace DocxEditor.Core.Markdown;

/// <summary>
/// Options controlling markdown style resolution against a document's styles part.
/// </summary>
public sealed class MarkdownStyleResolverOptions
{
    /// <summary>
    /// When true, unresolved references produce <see cref="MarkdownDiagnosticSeverity.Error"/>
    /// diagnostics and no style is emitted. When false (permissive), unresolved references
    /// produce a warning and a fallback style of the expected kind is generated. Existing
    /// template styles are never modified in either mode.
    /// </summary>
    public bool Strict { get; init; }

    public static MarkdownStyleResolverOptions Default { get; } = new();

    public static MarkdownStyleResolverOptions StrictMode { get; } = new() { Strict = true };
}
