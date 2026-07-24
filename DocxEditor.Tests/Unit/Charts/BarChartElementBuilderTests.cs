using PptxEditor.Core.Converters.Charts;
using PptxEditor.Core.Models;
using Xunit;

namespace DocxEditor.Tests.Unit.Charts;

/// <summary>
/// Emission math for clustered bar/column decomposition: value→bar scaling against the
/// auto axis, negative values, gap/overlap slot math, bottom-to-top category order for
/// horizontal bars, label placement, and legend swatches.
/// </summary>
public sealed class BarChartElementBuilderTests
{
    private static ChartModel Model(
        BarDirection direction,
        IReadOnlyList<string> categories,
        params (string Name, string Color, double?[] Values)[] series)
        => new()
        {
            Kind = ChartKind.Bar,
            Direction = direction,
            Grouping = BarGrouping.Clustered,
            Categories = categories,
            Series = series.Select(s => new ChartSeries
            {
                Name = s.Name,
                FillColor = s.Color,
                Values = s.Values
            }).ToList(),
            DataLabels = new ChartDataLabels { ShowValue = true, FontSize = 9, Color = "#111111" },
            ValueAxis = new ChartValueAxis { GridlineColor = "#DDDDDD", LabelFontSize = 9 },
            CategoryAxis = new ChartAxis { LabelFontSize = 9 }
        };

    private static List<TypstElement> Bars(IReadOnlyList<TypstElement> elements, string fill)
        => elements.Where(e => e.Type == "Shape" && e.Shape?.FillColor == fill).ToList();

    [Fact]
    public void Build_ColumnChart_BarHeightsScaleToAxisMax()
    {
        // Labels shown → axis 0..120 (unit 20) for a data max of 100.
        var model = Model(BarDirection.Column, ["A", "B"], ("S1", "#FF0000", [100.0, 50.0]));
        var elements = BarChartElementBuilder.Build(model, 0, 0, 400, 300);

        var bars = Bars(elements, "#FF0000");
        Assert.Equal(2, bars.Count);

        // Plot height = 300 − category band (9·1.5+4 = 17.5) = 282.5.
        const double plotH = 300 - (9 * 1.5 + 4);
        Assert.Equal(plotH * 100 / 120, bars[0].Height, precision: 3);
        Assert.Equal(plotH * 50 / 120, bars[1].Height, precision: 3);

        // Both bars sit on the baseline (plot bottom); taller bar starts higher.
        var baseline = bars[0].Y + bars[0].Height;
        Assert.Equal(baseline, bars[1].Y + bars[1].Height, precision: 3);
        Assert.True(bars[0].Y < bars[1].Y);
    }

    [Fact]
    public void Build_NegativeValues_BarGrowsDownFromZeroBaseline()
    {
        var model = Model(BarDirection.Column, ["A", "B"], ("S1", "#0000FF", [30.0, -10.0]));
        var elements = BarChartElementBuilder.Build(model, 0, 0, 400, 300);

        var bars = Bars(elements, "#0000FF");
        Assert.Equal(2, bars.Count);

        // Axis: data −10..30 → unit 10, min −10, max 40. Plot height as above.
        const double plotH = 300 - (9 * 1.5 + 4);
        var positive = bars.OrderBy(b => b.Y).First();
        var negative = bars.OrderBy(b => b.Y).Last();

        Assert.Equal(plotH * 30 / 50, positive.Height, precision: 3);
        Assert.Equal(plotH * 10 / 50, negative.Height, precision: 3);
        // The positive bar ends exactly where the negative bar begins: the zero baseline.
        Assert.Equal(positive.Y + positive.Height, negative.Y, precision: 3);
    }

    [Fact]
    public void Build_ClusteredSlots_RespectGapWidthAndOverlap()
    {
        // gapWidth 60, overlap −20 (sales deck values), 2 series.
        var model = Model(BarDirection.Column, ["A"],
            ("S1", "#111111", [10.0]), ("S2", "#222222", [20.0]));
        model = new ChartModel
        {
            Kind = model.Kind, Direction = model.Direction, Grouping = model.Grouping,
            Categories = model.Categories, Series = model.Series,
            DataLabels = model.DataLabels, ValueAxis = model.ValueAxis, CategoryAxis = model.CategoryAxis,
            GapWidthPercent = 60, OverlapPercent = -20
        };
        var elements = BarChartElementBuilder.Build(model, 0, 0, 280, 300);

        // Plot width = 280 − value band (ticks 0/20 → "20" ≈ 2·9·0.62+8 = 19.16).
        double valBand = 2 * 9 * 0.62 + 8;
        double slot = 280 - valBand;
        // bar = slot / (2 − (−0.2) + 0.6) = slot / 2.8
        var expectedBar = slot / 2.8;

        var s1 = Assert.Single(Bars(elements, "#111111"));
        var s2 = Assert.Single(Bars(elements, "#222222"));
        Assert.Equal(expectedBar, s1.Width, precision: 3);
        Assert.Equal(expectedBar, s2.Width, precision: 3);
        // Series bars are 1.2 bar-widths apart ((1 − overlap) · bar) inside the cluster.
        Assert.Equal(expectedBar * 1.2, s2.X - s1.X, precision: 3);
        // The cluster is centered in the slot: equal margins on both sides.
        var clusterLeft = s1.X - valBand;
        var clusterRight = (valBand + slot) - (s2.X + s2.Width);
        Assert.Equal(clusterLeft, clusterRight, precision: 3);
    }

    [Fact]
    public void Build_HorizontalChart_PlotsCategoriesBottomToTop()
    {
        var model = Model(BarDirection.Bar, ["First", "Second", "Third"],
            ("S1", "#00FF00", [10.0, 20.0, 30.0]));
        var elements = BarChartElementBuilder.Build(model, 0, 0, 400, 300);

        // Category labels (right-aligned text in the left band), bottom-to-top.
        var catLabels = elements
            .Where(e => e.Type == "Text" && model.Categories.Contains(e.Text?.Content))
            .OrderBy(e => e.Y)
            .Select(e => e.Text!.Content)
            .ToList();
        Assert.Equal(new[] { "Third", "Second", "First" }, catLabels);

        // Bar lengths follow values along X: the "Third" bar (30) is the longest.
        var bars = Bars(elements, "#00FF00").OrderBy(b => b.Y).ToList();
        Assert.Equal(3, bars.Count);
        Assert.True(bars[0].Width > bars[1].Width);
        Assert.True(bars[1].Width > bars[2].Width);
    }

    [Fact]
    public void Build_ValueLabels_PlacedOutsideBarEnds()
    {
        // Data labels use the dLbls colour (#111111); tick labels use the axis colour, so
        // filtering by colour isolates the value label from the "50" axis tick.
        static bool IsDataLabel(TypstElement e)
            => e.Type == "Text" && e.Text?.Content == "50" && e.Text.Formatting.Color == "#111111";

        var column = Model(BarDirection.Column, ["A"], ("S1", "#FF0000", [50.0]));
        var colElements = BarChartElementBuilder.Build(column, 0, 0, 400, 300);
        var colBar = Assert.Single(Bars(colElements, "#FF0000"));
        var colLabel = Assert.Single(colElements, IsDataLabel);
        // Column: label above the bar top, horizontally centered on it.
        Assert.True(colLabel.Y + colLabel.Height <= colBar.Y + 0.5);
        Assert.Equal(colBar.X + colBar.Width / 2, colLabel.X + colLabel.Width / 2, precision: 1);

        var horizontal = Model(BarDirection.Bar, ["A"], ("S1", "#FF0000", [50.0]));
        var horElements = BarChartElementBuilder.Build(horizontal, 0, 0, 400, 300);
        var horBar = Assert.Single(Bars(horElements, "#FF0000"));
        var horLabel = Assert.Single(horElements, IsDataLabel);
        // Horizontal: label to the right of the bar end, vertically centered on it.
        Assert.True(horLabel.X >= horBar.X + horBar.Width - 0.5);
        Assert.Equal(horBar.Y + horBar.Height / 2, horLabel.Y + horLabel.Height / 2, precision: 1);
    }

    [Fact]
    public void Build_GridlinesAndAxisLines_EmittedAtTicks()
    {
        var model = Model(BarDirection.Column, ["A"], ("S1", "#FF0000", [100.0]));
        var elements = BarChartElementBuilder.Build(model, 0, 0, 400, 300);

        // Axis 0..120 unit 20 → 6 interior gridlines (min tick skipped), as on the
        // reference render of sales slide 9.
        var gridlines = elements.Where(e => e.Type == "Shape" && e.Shape?.FillColor == "#DDDDDD").ToList();
        Assert.Equal(6, gridlines.Count);
        Assert.All(gridlines, g => Assert.True(g.Height < 1.0));
        // Gridlines span the full plot width (frame − value-label band for "120").
        var expectedPlotWidth = 400 - (3 * 9 * 0.62 + 8);
        Assert.All(gridlines, g => Assert.Equal(expectedPlotWidth, g.Width, precision: 3));
    }

    [Fact]
    public void Build_LegendRight_SwatchesMatchSeriesColors_ReversedForHorizontal()
    {
        var model = Model(BarDirection.Bar, ["A", "B"],
            ("Gray", "#B6B5B5", [100.0, 100.0]), ("Red", "#C00000", [68.0, 58.0]));
        model = new ChartModel
        {
            Kind = model.Kind, Direction = model.Direction, Grouping = model.Grouping,
            Categories = model.Categories, Series = model.Series,
            DataLabels = model.DataLabels, ValueAxis = model.ValueAxis, CategoryAxis = model.CategoryAxis,
            Legend = new ChartLegend { Position = ChartLegendPosition.Right, FontSize = 9 }
        };

        var elements = BarChartElementBuilder.Build(model, 0, 0, 600, 330);

        // Legend entries: text labels + same-colour swatches, reversed for barDir=bar
        // (Red/"Red" above Gray/"Gray"), matching the reference render.
        var redLabel = Assert.Single(elements, e => e.Text?.Content == "Red");
        var grayLabel = Assert.Single(elements, e => e.Text?.Content == "Gray");
        Assert.True(redLabel.Y < grayLabel.Y);

        var redSwatch = Assert.Single(elements,
            e => e.Type == "Shape" && e.Shape?.FillColor == "#C00000" && e.Width < 20 && e.Height < 20);
        var graySwatch = Assert.Single(elements,
            e => e.Type == "Shape" && e.Shape?.FillColor == "#B6B5B5" && e.Width < 20 && e.Height < 20);
        Assert.Equal(redLabel.Y, redSwatch.Y, 5.0);
        Assert.Equal(grayLabel.Y, graySwatch.Y, 5.0);
        Assert.True(redSwatch.X < redLabel.X);
    }

    [Fact]
    public void Build_NoDataOrNoRoom_ReturnsEmpty()
    {
        var empty = Model(BarDirection.Column, ["A"], ("S1", "#FF0000", new double?[] { null }));
        Assert.Empty(BarChartElementBuilder.Build(empty, 0, 0, 400, 300));

        var model = Model(BarDirection.Column, ["A"], ("S1", "#FF0000", [1.0]));
        Assert.Empty(BarChartElementBuilder.Build(model, 0, 0, 15, 15));
    }

    [Fact]
    public void Build_AllElements_StayInsideFrame()
    {
        var model = Model(BarDirection.Bar, ["Lead conversion", "Win rate", "Deal cycle speed"],
            ("Best-in-class", "#B6B5B5", [100.0, 100.0, 100.0]),
            ("Halcyon today", "#C00000", [68.0, 58.0, 74.0]));
        model = new ChartModel
        {
            Kind = model.Kind, Direction = model.Direction, Grouping = model.Grouping,
            Categories = model.Categories, Series = model.Series,
            DataLabels = model.DataLabels, ValueAxis = model.ValueAxis, CategoryAxis = model.CategoryAxis,
            GapWidthPercent = 60, OverlapPercent = -20,
            Legend = new ChartLegend { Position = ChartLegendPosition.Right, FontSize = 9 }
        };

        var elements = BarChartElementBuilder.Build(model, 60, 110, 600, 330);

        Assert.NotEmpty(elements);
        Assert.All(elements, e =>
        {
            Assert.True(e.X >= 59.0, $"X {e.X} outside frame ({e.Type} {e.Text?.Content})");
            Assert.True(e.Y >= 109.0, $"Y {e.Y} outside frame ({e.Type} {e.Text?.Content})");
            Assert.True(e.X + e.Width <= 662.0, $"right {e.X + e.Width} outside frame ({e.Type} {e.Text?.Content})");
            Assert.True(e.Y + e.Height <= 442.0, $"bottom {e.Y + e.Height} outside frame ({e.Type} {e.Text?.Content})");
        });
    }
}
