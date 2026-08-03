namespace PptxEditor.Core.Generation.Model;

/// <summary>Layout algorithm of a container. Absence of layout = free canvas.</summary>
public enum LayoutMode
{
    /// <summary>Children along the horizontal axis.</summary>
    Row,

    /// <summary>Children along the vertical axis.</summary>
    Column,

    /// <summary>Explicit columns; rows derived from child count, document order placement.</summary>
    Grid
}

/// <summary>Main-axis distribution of children inside a layout container.</summary>
public enum Justify
{
    /// <summary>Pack children at the start of the axis.</summary>
    Start,

    /// <summary>Center children on the axis.</summary>
    Center,

    /// <summary>Pack children at the end of the axis.</summary>
    End,

    /// <summary>Even gaps between children, none at the edges.</summary>
    SpaceBetween,

    /// <summary>Even gaps between children and at the edges.</summary>
    SpaceEvenly
}

/// <summary>Cross-axis alignment of children inside a layout container.</summary>
public enum AlignItems
{
    /// <summary>Align to the start of the cross axis.</summary>
    Start,

    /// <summary>Center on the cross axis.</summary>
    Center,

    /// <summary>Align to the end of the cross axis.</summary>
    End,

    /// <summary>Stretch to fill the cross axis (default).</summary>
    Stretch
}

/// <summary>
/// Layout declaration of a container: mode, gaps and axis alignment.
/// <see cref="RowGap"/>/<see cref="ColumnGap"/> are grid-only refinements of <see cref="Gap"/>.
/// </summary>
public sealed record LayoutSpec
{
    /// <summary>Layout algorithm. Required.</summary>
    public required LayoutMode Mode { get; init; }

    /// <summary>Gap between children in points (≥ 0).</summary>
    public double Gap { get; init; }

    /// <summary>Grid column count (≥ 1). Required when <see cref="Mode"/> is <see cref="LayoutMode.Grid"/>.</summary>
    public int? Columns { get; init; }

    /// <summary>Grid-only row gap override in points.</summary>
    public double? RowGap { get; init; }

    /// <summary>Grid-only column gap override in points.</summary>
    public double? ColumnGap { get; init; }

    /// <summary>Main-axis distribution. Defaults to <see cref="Justify.Start"/>.</summary>
    public Justify Justify { get; init; } = Justify.Start;

    /// <summary>Cross-axis alignment. Defaults to <see cref="AlignItems.Stretch"/>.</summary>
    public AlignItems Align { get; init; } = AlignItems.Stretch;
}

/// <summary>
/// Edge insets in points (top, right, bottom, left). Used for container "padding" and
/// text "insets". JSON: a single number (all), [v, h] or [top, right, bottom, left].
/// </summary>
public readonly record struct EdgeInsets(double Top, double Right, double Bottom, double Left)
{
    /// <summary>Uniform insets on all four edges.</summary>
    public static EdgeInsets All(double value) => new(value, value, value, value);

    /// <summary>Vertical/horizontal insets.</summary>
    public static EdgeInsets Symmetric(double vertical, double horizontal) => new(vertical, horizontal, vertical, horizontal);
}
