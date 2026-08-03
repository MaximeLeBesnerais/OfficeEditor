using PptxEditor.Core.Generation.Model;
using PptxEditor.Core.Models;

namespace PptxEditor.Core.Generation.Components;

/// <summary>
/// Strongly typed payloads for the v1 component set (component layer). The schema layer
/// (<see cref="Model.ComponentElement"/>) carries component content as a raw JSON bag;
/// these records are the contract the component layer validates it against (P1 handoff
/// note on <see cref="Model.ComponentElement.Content"/>). All color strings are palette
/// token names or #RRGGBB literals; all dimensions are points.
/// </summary>
public abstract record ComponentContent;

/// <summary>
/// <c>card</c> payload: a titled content surface. <see cref="Title"/> is required;
/// <see cref="Subtitle"/> and <see cref="Body"/> are optional. Overflow contract: each
/// text field shrinks (default text policy); the card surface itself errors on structural
/// overflow (default container policy).
/// </summary>
public sealed record CardContent : ComponentContent
{
    /// <summary>Card heading (one line; shrinks if too long). Required.</summary>
    public required string Title { get; init; }

    /// <summary>Muted secondary line under the title.</summary>
    public string? Subtitle { get; init; }

    /// <summary>Body copy; takes the card's remaining vertical space (grow 1).</summary>
    public string? Body { get; init; }
}

/// <summary>
/// <c>kpi</c> payload: a big number with a label and an optional delta. Overflow
/// contract: all texts shrink; the surface errors on structural overflow.
/// </summary>
public sealed record KpiContent : ComponentContent
{
    /// <summary>The headline figure (e.g. "+34%"). Required.</summary>
    public required string Value { get; init; }

    /// <summary>What the figure measures (e.g. "Revenue"). Required.</summary>
    public required string Label { get; init; }

    /// <summary>Optional trend line (e.g. "+12% vs LY"), rendered in the accent color.</summary>
    public string? Delta { get; init; }
}

/// <summary>
/// <c>title_block</c> payload: slide heading group (kicker over title over subtitle).
/// Overflow contract: texts shrink; structural overflow errors.
/// </summary>
public sealed record TitleBlockContent : ComponentContent
{
    /// <summary>Slide title in the display font at <c>metrics.titleSizePt</c>. Required.</summary>
    public required string Title { get; init; }

    /// <summary>Muted supporting line under the title.</summary>
    public string? Subtitle { get; init; }

    /// <summary>Small accent overline above the title.</summary>
    public string? Kicker { get; init; }
}

/// <summary>
/// <c>bullet_list</c> payload: a titled or untitled list of single-line items with
/// accent markers. Overflow contract: item texts shrink (long items get smaller, never
/// silently clip); too many items for the box is a structural error.
/// </summary>
public sealed record BulletListContent : ComponentContent
{
    /// <summary>Item texts, one line each. Required, at least one.</summary>
    public required IReadOnlyList<string> Items { get; init; }

    /// <summary>Optional bold heading above the items.</summary>
    public string? Title { get; init; }

    /// <summary>Marker color override (palette token or #RRGGBB). Defaults to "accent".</summary>
    public string? MarkerColor { get; init; }
}

/// <summary>
/// <c>divider</c> payload: a straight rule. Defaults: horizontal, "muted", 1 pt. The
/// cross-axis dimension defaults to the stroke width when the author does not fix it.
/// No text — geometry overflow is governed by the parent container's policy.
/// </summary>
public sealed record DividerContent : ComponentContent
{
    /// <summary>Stroke color (palette token or #RRGGBB). Defaults to "muted".</summary>
    public string Color { get; init; } = "muted";

    /// <summary>Stroke width in points (&gt; 0). Defaults to 1.</summary>
    public double WidthPt { get; init; } = 1;

    /// <summary>Line direction. Defaults to <see cref="LineOrientation.Horizontal"/>.</summary>
    public LineOrientation Orientation { get; init; } = LineOrientation.Horizontal;
}

/// <summary>
/// <c>badge</c> payload: a small pill label. Height is computed from the font metrics
/// line estimate unless the author fixes it; the author must give the badge a width (or
/// grow) on the parent's layout axis. Overflow contract: the label shrinks.
/// </summary>
public sealed record BadgeContent : ComponentContent
{
    /// <summary>Label text (one line). Required.</summary>
    public required string Text { get; init; }

    /// <summary>Pill fill color (palette token or #RRGGBB). Defaults to "accent".</summary>
    public string Color { get; init; } = "accent";

    /// <summary>Label color (palette token or #RRGGBB). Defaults to "paper".</summary>
    public string TextColor { get; init; } = "paper";
}

/// <summary>
/// <c>image_card</c> payload: an image over an optional caption block on a card surface.
/// Overflow contract: the image is clipped by its fit mode (default crop); caption texts
/// shrink; structural overflow errors.
/// </summary>
public sealed record ImageCardContent : ComponentContent
{
    /// <summary>Image source (path, URL or base64 payload — interpreted by emitters). Required.</summary>
    public required string Source { get; init; }

    /// <summary>Caption heading under the image.</summary>
    public string? Title { get; init; }

    /// <summary>Muted caption line under the title.</summary>
    public string? Subtitle { get; init; }

    /// <summary>Image fit mode. Defaults to <see cref="ImageFitMode.Crop"/>.</summary>
    public ImageFitMode Fit { get; init; } = ImageFitMode.Crop;

    /// <summary>Accessibility alt text.</summary>
    public string? Alt { get; init; }

    /// <summary>Share of the card's vertical space the image grows into (&gt; 0). Defaults to 1.</summary>
    public double ImageGrow { get; init; } = 1;
}

/// <summary>
/// <c>table_block</c> payload: a header row plus body rows over the primitive grid
/// (the component layer delegates the table look to the existing row/column emission path — the
/// component expands to plain row containers and cell texts, no OOXML/Typst calls of its
/// own). Overflow contract: cell texts shrink; rows that do not fit the box are a
/// structural error.
/// </summary>
public sealed record TableBlockContent : ComponentContent
{
    /// <summary>Column header labels. Required, at least one.</summary>
    public required IReadOnlyList<string> Columns { get; init; }

    /// <summary>Body rows; each row must have exactly <see cref="Columns"/> cells. May be empty.</summary>
    public required IReadOnlyList<IReadOnlyList<string>> Rows { get; init; }

    /// <summary>Whether the header row (primary fill) is emitted. Defaults to true.</summary>
    public bool Header { get; init; } = true;

    /// <summary>Optional relative column widths (grow weights, &gt; 0, one per column). Default: all equal.</summary>
    public IReadOnlyList<double>? ColumnWeights { get; init; }

    /// <summary>Row height in points. Defaults to the body-size line estimate plus cell padding.</summary>
    public double RowHeight { get; init; }
}
