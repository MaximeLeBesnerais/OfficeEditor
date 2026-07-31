using PptxEditor.Core.Converters.Charts;
using Xunit;

namespace DocxEditor.Tests.Unit.Charts;

public sealed class PieChartElementBuilderTests
{
    [Fact]
    public void Build_DoughnutBottomLegend_UsesPositiveSliceCategoryIndices()
    {
        var model = new ChartModel
        {
            Kind = ChartKind.Pie,
            HoleSizePercent = 45,
            Categories = ["zero", "item 2", "item 3"],
            Legend = new ChartLegend { Position = ChartLegendPosition.Bottom },
            Series = [new ChartSeries
            {
                Values = [0, 60, 40],
                PointFillColors = ["#000000", "#123456", "#654321"]
            }]
        };

        var elements = PieChartElementBuilder.Build(model, 0, 0, 200, 100);

        Assert.Contains(elements, e => e.Type == "Text" && e.Text!.Content == "item 2");
        Assert.Contains(elements, e => e.Type == "Text" && e.Text!.Content == "item 3");
        Assert.DoesNotContain(elements, e => e.Type == "Text" && e.Text!.Content == "zero");
    }

    [Fact]
    public void Build_PieWithoutDoughnutHole_DoesNotApproximateLegend()
    {
        var model = new ChartModel
        {
            Kind = ChartKind.Pie,
            Categories = ["item 1"],
            Legend = new ChartLegend { Position = ChartLegendPosition.Bottom },
            Series = [new ChartSeries { Values = [100] }]
        };

        var elements = PieChartElementBuilder.Build(model, 0, 0, 200, 100);

        Assert.DoesNotContain(elements, e => e.Type == "Text" && e.Text!.Content == "item 1");
    }

    [Fact]
    public void Build_DoughnutLegend_UsesGreedyRowsForLongAndNonPositiveValues()
    {
        var model = new ChartModel
        {
            Kind = ChartKind.Pie,
            HoleSizePercent = 45,
            Categories =
            [
                "A very long category label",
                "negative",
                "Another very long category label",
                "zero",
                "Third long category label"
            ],
            Legend = new ChartLegend { Position = ChartLegendPosition.Bottom },
            Series = [new ChartSeries { Values = [60, -20, 40, 0, 20] }]
        };

        var elements = PieChartElementBuilder.Build(model, 0, 0, 200, 160);
        var legendTexts = elements
            .Where(e => e.Type == "Text")
            .Select(e => e.Text!.Content)
            .ToList();
        var legendRows = elements
            .Where(e => e.Type == "Text")
            .Select(e => e.Y)
            .Distinct()
            .Count();

        Assert.Equal(3, legendTexts.Count);
        Assert.DoesNotContain("negative", legendTexts);
        Assert.DoesNotContain("zero", legendTexts);
        Assert.Equal(3, legendRows);
    }
}
