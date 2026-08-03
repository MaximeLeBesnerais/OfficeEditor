namespace DocxEditor.Core.Markdown.Model;

/// <summary>Plain text content. Escaped literal characters are flagged via <see cref="IsEscaped"/>.</summary>
public sealed record MarkdownText : MarkdownInline
{
    public string Text { get; init; } = string.Empty;

    public bool IsEscaped { get; init; }
}

/// <summary>
/// Emphasis (italic, bold, strikethrough, subscript, superscript, inserted or
/// marked). Nested formatting is preserved: <see cref="Children"/> may itself
/// contain emphasis, links, code and other inline nodes.
/// </summary>
public sealed record MarkdownEmphasis : MarkdownInline
{
    public EmphasisKind Kind { get; init; }

    public IReadOnlyList<MarkdownInline> Children { get; init; } = [];
}

/// <summary>An inline code span. Content is retained verbatim including inner backticks.</summary>
public sealed record MarkdownCode : MarkdownInline
{
    public string Content { get; init; } = string.Empty;

    public char Delimiter { get; init; }

    public int DelimiterCount { get; init; }
}

/// <summary>A hyperlink with optional title and (nested) display text.</summary>
public sealed record MarkdownLink : MarkdownInline
{
    public string? Url { get; init; }

    public string? Title { get; init; }

    public bool IsAutoLink { get; init; }

    public bool IsShortcut { get; init; }

    public IReadOnlyList<MarkdownInline> Children { get; init; } = [];
}

/// <summary>An image; children hold the alt text.</summary>
public sealed record MarkdownImage : MarkdownInline
{
    public string? Url { get; init; }

    public string? Title { get; init; }

    public IReadOnlyList<MarkdownInline> Children { get; init; } = [];
}

/// <summary>A hard (two trailing spaces or backslash) or soft (plain newline) line break.</summary>
public sealed record MarkdownLineBreak : MarkdownInline
{
    public bool IsHard { get; init; }

    public bool IsBackslash { get; init; }
}

/// <summary>Raw inline HTML retained as an explicit node rather than being dropped.</summary>
public sealed record MarkdownHtml : MarkdownInline
{
    public string Tag { get; init; } = string.Empty;
}

/// <summary>An HTML entity that was safely decoded. <see cref="Original"/> is the source text
/// (e.g. "&amp;amp;") and <see cref="Decoded"/> is the decoded value (e.g. "&amp;").</summary>
public sealed record MarkdownEntity : MarkdownInline
{
    public string Original { get; init; } = string.Empty;

    public string Decoded { get; init; } = string.Empty;
}

/// <summary>An emoji shortcode (e.g. :smile:) with its decoded Unicode <see cref="Text"/>.</summary>
public sealed record MarkdownEmoji : MarkdownInline
{
    public string Match { get; init; } = string.Empty;

    public string Text { get; init; } = string.Empty;
}

/// <summary>A task-list checkbox marker ([ ] or [x]). Carried as an inline inside a list item.</summary>
public sealed record MarkdownTaskCheckbox : MarkdownInline
{
    public bool Checked { get; init; }
}

/// <summary>A footnote reference (or, when <see cref="IsBackLink"/>, the generated backlink inside
/// a footnote definition).</summary>
public sealed record MarkdownFootnoteReference : MarkdownInline
{
    public string Label { get; init; } = string.Empty;

    public int Index { get; init; }

    public bool IsBackLink { get; init; }
}

/// <summary>An inline construct that was not mapped to a structured node. It is retained
/// verbatim (with a diagnostic in strict mode) instead of being silently dropped.</summary>
public sealed record MarkdownUnknownInline : MarkdownInline
{
    /// <summary>The Markdig inline type that could not be mapped.</summary>
    public string Kind { get; init; } = string.Empty;

    public string Raw { get; init; } = string.Empty;
}
