namespace DocxEditor.Core.Markdown.Model;

/// <summary>
/// Base type for every node of the rich markdown IR. Nodes are immutable and
/// recursive: <see cref="MarkdownBlock"/> nodes may contain other blocks and
/// inline nodes, and <see cref="MarkdownInline"/> nodes may contain nested
/// inline nodes (for example emphasis inside emphasis).
/// </summary>
public abstract record MarkdownNode
{
    /// <summary>
    /// The location of this node in the source, when Markdig exposes one.
    /// </summary>
    public SourceSpan? Span { get; init; }
}

/// <summary>
/// Base type for block-level markdown constructs.
/// </summary>
public abstract record MarkdownBlock : MarkdownNode;

/// <summary>
/// Base type for inline markdown constructs.
/// </summary>
public abstract record MarkdownInline : MarkdownNode;

/// <summary>
/// The root of a parsed markdown document.
/// </summary>
public sealed record MarkdownDocument : MarkdownNode
{
    public IReadOnlyList<MarkdownBlock> Blocks { get; init; } = [];
}
