namespace DocxEditor.Core.Markdown.Model;

/// <summary>
/// Kinds of emphasis formatting retained from markdown delimiters.
/// </summary>
public enum EmphasisKind
{
    /// <summary>*emphasized* or _emphasized_</summary>
    Italic,
    /// <summary>**strong** or __strong__</summary>
    Bold,
    /// <summary>~~strikethrough~~</summary>
    Strikethrough,
    /// <summary>~subscript~</summary>
    Subscript,
    /// <summary>^superscript^</summary>
    Superscript,
    /// <summary>++inserted++</summary>
    Inserted,
    /// <summary>==marked==</summary>
    Marked
}

/// <summary>
/// Alignment of a table column.
/// </summary>
public enum MarkdownTableAlignment
{
    Left,
    Center,
    Right
}

/// <summary>
/// Severity of a diagnostic produced while parsing markdown.
/// </summary>
public enum MarkdownDiagnosticSeverity
{
    Info,
    Warning,
    Error
}
