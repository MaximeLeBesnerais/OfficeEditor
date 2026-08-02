namespace DocxEditor.Core.Generation.Model;

/// <summary>
/// Base of every flow block. Flow is the DOCX-native ordering model: paragraphs, tables
/// and images laid out in document order inside a section. <see cref="Style"/> is the
/// optional paragraph/table/list/picture style reference.
/// </summary>
public abstract record FlowBlock
{
    /// <summary>Optional style reference (resolved by the emitter).</summary>
    public string? Style { get; init; }
}

/// <summary>A body paragraph. Content holds text/runs, token, alignment and spacing.</summary>
public sealed record ParagraphBlock : FlowBlock
{
    /// <summary>Paragraph content. Required.</summary>
    public required TextModel Content { get; init; }
}

/// <summary>A heading paragraph at a given outline level.</summary>
public sealed record HeadingBlock : FlowBlock
{
    /// <summary>Heading level (1..6).</summary>
    public int Level { get; init; } = 1;

    /// <summary>Heading content. Required.</summary>
    public required TextModel Content { get; init; }
}

/// <summary>A bullet or ordered list. v1 is single-level; ordered lists start at
/// <see cref="StartIndex"/> (1 by default).</summary>
public sealed record ListBlock : FlowBlock
{
    /// <summary>List flavor. Defaults to <see cref="ListKind.Bullet"/>.</summary>
    public ListKind Kind { get; init; } = ListKind.Bullet;

    /// <summary>First index for ordered lists (≥ 1). Ignored for bullet lists.</summary>
    public int? StartIndex { get; init; }

    /// <summary>Items in document order. Required, at least one.</summary>
    public required IReadOnlyList<TextModel> Items { get; init; }
}

/// <summary>
/// A flow callout: an emphasized quote/note box that stays in the flow (emitters render it
/// as a shaded/bordered paragraph or single-cell table). For a free-floating callout box
/// use the positioned tier's <c>callout</c>.
/// </summary>
public sealed record CalloutBlock : FlowBlock
{
    /// <summary>Callout tone (drives the default visual treatment).</summary>
    public CalloutTone Tone { get; init; } = CalloutTone.Note;

    /// <summary>Callout text content. Required.</summary>
    public required TextModel Content { get; init; }
}

/// <summary>A hard page break marker (no properties).</summary>
public sealed record PageBreakBlock : FlowBlock;

/// <summary>
/// Section-safe raw grouping container: a transparent wrapper over an ordered list of
/// nested flow blocks. It exists for authoring convenience (move/copy a unit of content);
/// emitters may flatten it into the parent flow or wrap it in a structural container.
/// It can never contain sections — only flow blocks.
/// </summary>
public sealed record FlowContainerBlock : FlowBlock
{
    /// <summary>Nested flow blocks. Required, at least one.</summary>
    public required IReadOnlyList<FlowBlock> Blocks { get; init; }
}
