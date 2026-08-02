namespace DocxEditor.Core.Generation.Model;

/// <summary>
/// Text content model shared by paragraphs, headings, list items, callouts and table
/// cells. Exactly one of <see cref="Text"/> / <see cref="Runs"/> is set. A <see
/// cref="Token"/> references a design typography token; alignment and spacing apply to the
/// paragraph the content lives in.
/// </summary>
public sealed record TextModel
{
    /// <summary>Plain text content ("text" in JSON); mutually exclusive with <see cref="Runs"/>.</summary>
    public string? Text { get; init; }

    /// <summary>Styled runs ("runs" in JSON); mutually exclusive with <see cref="Text"/>.</summary>
    public IReadOnlyList<Run>? Runs { get; init; }

    /// <summary>Named design typography token reference (see <c>design.typography</c>).</summary>
    public string? Token { get; init; }

    /// <summary>
    /// Semantic text role (see <see cref="TextRole"/>). Null = infer from the block kind
    /// (heading level / body). The active theme supplies default formatting for the role,
    /// which <see cref="Token"/> and direct run formatting override.
    /// </summary>
    public TextRole? Role { get; init; }

    /// <summary>Horizontal alignment of the paragraph. Null = emitter/inherited default.</summary>
    public TextAlignment? Alignment { get; init; }

    /// <summary>Paragraph spacing (before/after/line). Null = emitter/inherited default.</summary>
    public ParagraphSpacing? Spacing { get; init; }
}

/// <summary>One styled run. The atomic unit of character formatting.</summary>
public sealed record Run
{
    /// <summary>Run text. Required.</summary>
    public required string Text { get; init; }

    /// <summary>Character style reference (emitter resolves it in the template or styles part).</summary>
    public string? Style { get; init; }

    /// <summary>Font slot token ("display" | "body") or a raw family name.</summary>
    public string? FontFamily { get; init; }

    /// <summary>Font size in points (&gt; 0).</summary>
    public double? FontSizePt { get; init; }

    /// <summary>Palette token name or #RRGGBB literal.</summary>
    public string? Color { get; init; }

    /// <summary>Bold weight.</summary>
    public bool Bold { get; init; }

    /// <summary>Italic style.</summary>
    public bool Italic { get; init; }

    /// <summary>Underline.</summary>
    public bool Underline { get; init; }

    /// <summary>All-caps rendering (<c>w:caps</c>).</summary>
    public bool AllCaps { get; init; }
}

/// <summary>Paragraph spacing. <see cref="BeforePt"/> / <see cref="AfterPt"/> may come from
/// a design spacing token (resolved at parse time); <see cref="LineMultiple"/> is a raw
/// line-spacing multiple (1 = single, 1.5 = one-and-a-half, 2 = double).</summary>
public sealed record ParagraphSpacing
{
    /// <summary>Space before the paragraph in points (≥ 0).</summary>
    public double? BeforePt { get; init; }

    /// <summary>Space after the paragraph in points (≥ 0).</summary>
    public double? AfterPt { get; init; }

    /// <summary>Line spacing multiple (&gt; 0).</summary>
    public double? LineMultiple { get; init; }
}
