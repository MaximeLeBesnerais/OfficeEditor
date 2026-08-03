using PptxEditor.Core.Generation.Model;

namespace PptxEditor.Core.Generation.Archetypes;

/// <summary>
/// Strongly typed payloads for the archetype slide functions (cover,
/// section, kpi_row, two_col, table_slide). Archetypes are thin compositions of the
/// v1 component set: these records carry only content, never geometry — the archetype
/// functions supply the layout. All strings are plain text; component-level styling
/// (fonts, colors, sizes) comes from the design tokens.
/// </summary>
public abstract record ArchetypeContent;

/// <summary>
/// <c>cover</c> payload: the deck's opening slide — a vertically centered title block
/// (kicker over display-font title over muted subtitle) with an accent rule beneath.
/// </summary>
public sealed record CoverContent : ArchetypeContent
{
    /// <summary>Deck title in the display font at <c>metrics.titleSizePt</c>. Required.</summary>
    public required string Title { get; init; }

    /// <summary>Muted supporting line under the title.</summary>
    public string? Subtitle { get; init; }

    /// <summary>Small accent overline above the title.</summary>
    public string? Kicker { get; init; }
}

/// <summary>
/// <c>section</c> payload: a section divider, bottom-anchored behind an accent rule.
/// With <see cref="Index"/> the slide becomes a big-number divider (the index renders
/// as a kpi figure beside the title); without it, a plain title block.
/// </summary>
public sealed record SectionContent : ArchetypeContent
{
    /// <summary>Section title in the display font at <c>metrics.titleSizePt</c>. Required.</summary>
    public required string Title { get; init; }

    /// <summary>Optional section number/label (e.g. "02") rendered as the big figure.</summary>
    public string? Index { get; init; }

    /// <summary>Muted supporting line under the title.</summary>
    public string? Subtitle { get; init; }

    /// <summary>Small overline. With <see cref="Index"/> it labels the figure; otherwise it kicks the title.</summary>
    public string? Kicker { get; init; }
}

/// <summary>
/// <c>kpi_row</c> payload: an optional heading over a row of KPI cards sharing the
/// slide width equally (the shrink contract applies per KPI value).
/// </summary>
public sealed record KpiRowContent : ArchetypeContent
{
    /// <summary>Optional slide title.</summary>
    public string? Title { get; init; }

    /// <summary>Optional muted line under the title.</summary>
    public string? Subtitle { get; init; }

    /// <summary>The KPI cells, 1–6 items, in document order. Required.</summary>
    public required IReadOnlyList<KpiItem> Kpis { get; init; }
}

/// <summary>One KPI cell inside a <c>kpi_row</c> slide (mirrors the <c>kpi</c> component payload).</summary>
public sealed record KpiItem
{
    /// <summary>The headline figure (e.g. "+34%"). Required.</summary>
    public required string Value { get; init; }

    /// <summary>What the figure measures (e.g. "Revenue"). Required.</summary>
    public required string Label { get; init; }

    /// <summary>Optional trend line, rendered in the accent color.</summary>
    public string? Delta { get; init; }
}

/// <summary>
/// <c>two_col</c> payload: an optional heading over two side-by-side component cells.
/// Each slot is one v1 component (card, kpi, title_block, bullet_list, divider, badge,
/// image_card, table_block) with its own content payload; relative widths come from
/// <see cref="Weights"/> (grow weights, default 1:1).
/// </summary>
public sealed record TwoColContent : ArchetypeContent
{
    /// <summary>Optional slide title.</summary>
    public string? Title { get; init; }

    /// <summary>Optional muted line under the title.</summary>
    public string? Subtitle { get; init; }

    /// <summary>Left cell component (size/at are assigned by the archetype). Required.</summary>
    public required ComponentElement Left { get; init; }

    /// <summary>Right cell component (size/at are assigned by the archetype). Required.</summary>
    public required ComponentElement Right { get; init; }

    /// <summary>Optional relative column widths (exactly two grow weights &gt; 0, left then right). Default: [1, 1].</summary>
    public IReadOnlyList<double>? Weights { get; init; }
}

/// <summary>
/// <c>table_slide</c> payload: an optional heading over a table that fills the
/// remaining slide height (the look is the <c>table_block</c> component's).
/// </summary>
public sealed record TableSlideContent : ArchetypeContent
{
    /// <summary>Optional slide title.</summary>
    public string? Title { get; init; }

    /// <summary>Optional muted line under the title.</summary>
    public string? Subtitle { get; init; }

    /// <summary>Column header labels. Required, at least one.</summary>
    public required IReadOnlyList<string> Columns { get; init; }

    /// <summary>Body rows; each row must have exactly <see cref="Columns"/> cells. May be empty.</summary>
    public required IReadOnlyList<IReadOnlyList<string>> Rows { get; init; }

    /// <summary>Whether the header row (primary fill) is emitted. Defaults to true.</summary>
    public bool Header { get; init; } = true;

    /// <summary>Optional relative column widths (grow weights, &gt; 0, one per column). Default: all equal.</summary>
    public IReadOnlyList<double>? ColumnWeights { get; init; }

    /// <summary>Row height in points. Defaults to the body-size line estimate plus cell padding.</summary>
    public double? RowHeight { get; init; }
}
