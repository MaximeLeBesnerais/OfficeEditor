namespace DocxEditor.Core.Markdown.Model;

/// <summary>A heading (ATX or Setext).</summary>
public sealed record MarkdownHeading : MarkdownBlock
{
    public int Level { get; init; }

    public IReadOnlyList<MarkdownInline> Inlines { get; init; } = [];
}

/// <summary>A paragraph of inline content.</summary>
public sealed record MarkdownParagraph : MarkdownBlock
{
    public IReadOnlyList<MarkdownInline> Inlines { get; init; } = [];
}

/// <summary>
/// A (possibly nested) bullet or ordered list. <see cref="Items"/> are the direct
/// list items; nesting is preserved because items carry their own block content.
/// </summary>
public sealed record MarkdownList : MarkdownBlock
{
    public bool Ordered { get; init; }

    /// <summary>The bullet character for unordered lists ('-', '*', '+') or null for ordered lists.</summary>
    public char? BulletType { get; init; }

    /// <summary>The start number of an ordered list as authored, e.g. "5".</summary>
    public string? OrderedStart { get; init; }

    public bool IsLoose { get; init; }

    public IReadOnlyList<MarkdownListItem> Items { get; init; } = [];
}

/// <summary>A single list item. Task state is retained through a <see cref="MarkdownTaskCheckbox"/>
/// inline inside the item content.</summary>
public sealed record MarkdownListItem : MarkdownBlock
{
    /// <summary>The numeric order of an ordered-list item as authored (1-based), or 0 for bullets.</summary>
    public int Order { get; init; }

    public IReadOnlyList<MarkdownBlock> Blocks { get; init; } = [];
}

/// <summary>A block quote; nested blocks are preserved.</summary>
public sealed record MarkdownQuote : MarkdownBlock
{
    /// <summary>Set when the quote was parsed as an alert (for example [!NOTE]) and the alert kind was retained.</summary>
    public string? AlertKind { get; init; }

    public IReadOnlyList<MarkdownBlock> Blocks { get; init; } = [];
}

/// <summary>A fenced or indented code block. <see cref="Text"/> preserves all newlines and indentation.</summary>
public sealed record MarkdownCodeBlock : MarkdownBlock
{
    public string Text { get; init; } = string.Empty;

    public string? Language { get; init; }

    public string? Arguments { get; init; }
}

/// <summary>A pipe or grid table.</summary>
public sealed record MarkdownTable : MarkdownBlock
{
    public IReadOnlyList<MarkdownTableColumn> Columns { get; init; } = [];

    public IReadOnlyList<MarkdownTableRow> Rows { get; init; } = [];
}

public sealed record MarkdownTableColumn
{
    public MarkdownTableAlignment? Alignment { get; init; }

    public float? Width { get; init; }
}

public sealed record MarkdownTableRow
{
    public bool IsHeader { get; init; }

    public IReadOnlyList<MarkdownTableCell> Cells { get; init; } = [];
}

/// <summary>A single table cell; its content is a list of blocks (typically a paragraph).</summary>
public sealed record MarkdownTableCell
{
    public int ColumnIndex { get; init; } = -1;

    public int ColumnSpan { get; init; } = 1;

    public int RowSpan { get; init; } = 1;

    public IReadOnlyList<MarkdownBlock> Blocks { get; init; } = [];
}

/// <summary>A thematic break (---, ***, ___).</summary>
public sealed record MarkdownThematicBreak : MarkdownBlock
{
    public char Character { get; init; }

    public int Count { get; init; }
}

/// <summary>Raw HTML retained as an explicit node rather than being dropped.</summary>
public sealed record MarkdownHtmlBlock : MarkdownBlock
{
    public string Text { get; init; } = string.Empty;

    /// <summary>The Markdig HTML block category, e.g. "InterruptingBlock".</summary>
    public string? Type { get; init; }
}

/// <summary>YAML front matter retained verbatim.</summary>
public sealed record MarkdownYamlFrontMatter : MarkdownBlock
{
    public string Yaml { get; init; } = string.Empty;
}

/// <summary>A definition list: one or more items, each with terms and definitions.</summary>
public sealed record MarkdownDefinitionList : MarkdownBlock
{
    public IReadOnlyList<MarkdownDefinitionItem> Items { get; init; } = [];
}

public sealed record MarkdownDefinitionItem : MarkdownBlock
{
    public IReadOnlyList<MarkdownDefinitionTerm> Terms { get; init; } = [];

    public IReadOnlyList<MarkdownBlock> Definitions { get; init; } = [];
}

public sealed record MarkdownDefinitionTerm : MarkdownBlock
{
    public IReadOnlyList<MarkdownInline> Inlines { get; init; } = [];
}

/// <summary>The footnotes section retained at the end of the document.</summary>
public sealed record MarkdownFootnotesBlock : MarkdownBlock
{
    public IReadOnlyList<MarkdownFootnote> Footnotes { get; init; } = [];
}

public sealed record MarkdownFootnote : MarkdownBlock
{
    public string Label { get; init; } = string.Empty;

    public int Order { get; init; }

    public IReadOnlyList<MarkdownBlock> Blocks { get; init; } = [];
}

/// <summary>Retained link reference definitions (for example <c>[id]: url</c>) that Markdig
/// stores as metadata rather than document content.</summary>
public sealed record MarkdownLinkReferenceDefinitions : MarkdownBlock
{
    public IReadOnlyList<MarkdownLinkReferenceDefinition> Definitions { get; init; } = [];
}

public sealed record MarkdownLinkReferenceDefinition : MarkdownBlock
{
    public string Label { get; init; } = string.Empty;

    public string Url { get; init; } = string.Empty;

    public string? Title { get; init; }
}

/// <summary>A block construct that was not mapped to a structured node. It is retained
/// verbatim (with a diagnostic in strict mode) instead of being silently dropped.</summary>
public sealed record MarkdownUnknownBlock : MarkdownBlock
{
    /// <summary>The Markdig block type that could not be mapped.</summary>
    public string Kind { get; init; } = string.Empty;

    public string Raw { get; init; } = string.Empty;

    /// <summary>Optional structured detail retained from the original construct.</summary>
    public string? Info { get; init; }
}
