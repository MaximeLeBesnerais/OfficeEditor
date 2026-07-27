namespace PptxEditor.Core.Converters.Charts;

/// <summary>Top-level chart kind found under <c>c:plotArea</c>.</summary>
public enum ChartKind
{
    Bar,
    Line,
    Pie,
    Other
}

/// <summary><c>c:barDir</c> — vertical columns or horizontal bars.</summary>
public enum BarDirection
{
    Column,
    Bar
}

/// <summary><c>c:grouping</c> for bar charts.</summary>
public enum BarGrouping
{
    Clustered,
    Stacked,
    PercentStacked
}

/// <summary><c>c:legendPos</c>; PowerPoint defaults to right when the element is absent.</summary>
public enum ChartLegendPosition
{
    Right,
    Left,
    Top,
    Bottom,
    TopRight
}

/// <summary>
/// Immutable model of a parsed <c>c:chartSpace</c> part. All values come from the cached
/// <c>c:strCache</c>/<c>c:numCache</c> points — cell references (<c>c:f</c>) are never
/// evaluated. Colours are resolved to <c>#RRGGBB</c> at parse time.
/// </summary>
public sealed class ChartModel
{
    public ChartKind Kind { get; init; } = ChartKind.Other;
    public BarDirection Direction { get; init; } = BarDirection.Column;
    public BarGrouping Grouping { get; init; } = BarGrouping.Clustered;
    public IReadOnlyList<ChartSeries> Series { get; init; } = [];
    /// <summary>Category labels, taken from the first series' category cache.</summary>
    public IReadOnlyList<string> Categories { get; init; } = [];
    /// <summary>Gap between category clusters as a percentage of bar width (<c>c:gapWidth</c>, default 150).</summary>
    public double GapWidthPercent { get; init; } = 150.0;
    /// <summary>Overlap between series bars within a cluster (<c>c:overlap</c>, −100…100).</summary>
    public double OverlapPercent { get; init; }
    /// <summary>Chart-level data label settings; per-series settings override these.</summary>
    public ChartDataLabels? DataLabels { get; init; }
    public ChartLegend? Legend { get; init; }
    public ChartAxis? CategoryAxis { get; init; }
    public ChartValueAxis? ValueAxis { get; init; }
    /// <summary>Pie/doughnut only: <c>c:firstSliceAng</c> — angle of the first slice in
    /// degrees, clockwise from 12 o'clock. Defaults to 0 when absent.</summary>
    public double FirstSliceAngleDegrees { get; init; }
}

public sealed class ChartSeries
{
    public string Name { get; init; } = string.Empty;
    /// <summary>Resolved solid fill, or null to use the default palette by series index.</summary>
    public string? FillColor { get; init; }
    /// <summary>Cached values indexed by category; null for missing points.</summary>
    public IReadOnlyList<double?> Values { get; init; } = [];
    public ChartDataLabels? DataLabels { get; init; }
    /// <summary>Pie/doughnut per-data-point fills (<c>c:dPt/c:spPr/a:solidFill</c>),
    /// indexed by point; null entries fall back to <see cref="FillColor"/>/palette.</summary>
    public IReadOnlyList<string?> PointFillColors { get; init; } = [];
    /// <summary>Pie/doughnut slice outline colour (first <c>c:dPt</c> line), if any.</summary>
    public string? PointLineColor { get; init; }
    /// <summary>Pie/doughnut slice outline width in points (<c>a:ln w</c>, EMU → pt).</summary>
    public double PointLineWidthPt { get; init; }
}

/// <summary>Data label settings from a <c>c:dLbls</c> element.</summary>
public sealed class ChartDataLabels
{
    public bool ShowValue { get; init; }
    /// <summary>Font size in points (<c>a:defRPr sz</c> is 1/100 pt).</summary>
    public double? FontSize { get; init; }
    public string? Color { get; init; }
    public string? FontFamily { get; init; }
}

public sealed class ChartLegend
{
    public ChartLegendPosition Position { get; init; } = ChartLegendPosition.Right;
    public double? FontSize { get; init; }
    public string? Color { get; init; }
    public string? FontFamily { get; init; }
}

/// <summary>Axis cosmetics shared by category and value axes.</summary>
public class ChartAxis
{
    /// <summary><c>c:delete val="1"</c> — axis (line, ticks and labels) hidden.</summary>
    public bool Deleted { get; init; }
    /// <summary>Explicit axis line colour (<c>c:spPr/a:ln</c> solid fill), if any.</summary>
    public string? LineColor { get; init; }
    public double? LabelFontSize { get; init; }
    public string? LabelColor { get; init; }
    public string? LabelFontFamily { get; init; }
}

public sealed class ChartValueAxis : ChartAxis
{
    /// <summary>Explicit <c>c:scaling</c> bounds; null = auto.</summary>
    public double? Min { get; init; }
    public double? Max { get; init; }
    public double? MajorUnit { get; init; }
    /// <summary>Major gridline colour, or null when the axis has no major gridlines.</summary>
    public string? GridlineColor { get; init; }
    /// <summary><c>c:numFmt formatCode</c> (e.g. "General", "0%").</summary>
    public string? NumberFormat { get; init; }
}
