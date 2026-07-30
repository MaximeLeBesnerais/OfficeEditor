using PptxEditor.Core.Converters.Charts;
using PptxEditor.Core.Models;
using Xunit;

namespace DocxEditor.Tests.Unit;

/// <summary>Focused regressions for the local FISHBONE/FRANCE chart and connector markup.</summary>
public sealed class PptxToTypstConverterNightFishFranceTests
{
    [Fact]
    public void Parse_DoughnutChart_PreservesHoleSizeAndTransformedPointColor()
    {
        const string xml = """
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea><c:doughnutChart><c:holeSize val="45"/><c:ser>
                <c:tx><c:v>Series</c:v></c:tx>
                <c:cat><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt></c:strCache></c:cat>
                <c:val><c:numCache><c:pt idx="0"><c:v>100</c:v></c:pt></c:numCache></c:val>
                <c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="336699"><a:lumMod val="50000"/></a:srgbClr></a:solidFill></c:spPr></c:dPt>
              </c:ser></c:doughnutChart></c:plotArea></c:chart>
            </c:chartSpace>
            """;

        var model = ChartPartParser.Parse(xml);

        Assert.NotNull(model);
        Assert.Equal(45, model!.HoleSizePercent);
        Assert.Equal("#1A334C", model.Series[0].PointFillColors[0]);
    }

    [Fact]
    public void Build_DoughnutChart_EmitsOuterAndInnerContours()
    {
        var model = new ChartModel
        {
            Kind = ChartKind.Pie,
            HoleSizePercent = 45,
            Series =
            [
                new ChartSeries
                {
                    Values = [60, 40],
                    PointFillColors = ["#123456", "#654321"]
                }
            ]
        };

        var elements = PieChartElementBuilder.Build(model, 0, 0, 200, 100);

        Assert.Equal(2, elements.Count);
        Assert.All(elements, element => Assert.Equal(2, element.Shape!.Subpaths.Count));
    }

    [Fact]
    public void Build_StackedBars_UsesOneCategoryBarWithSeriesSegments()
    {
        var model = new ChartModel
        {
            Kind = ChartKind.Bar,
            Grouping = BarGrouping.PercentStacked,
            Categories = ["A"],
            Series =
            [
                new ChartSeries { Values = [25], FillColor = "#111111" },
                new ChartSeries { Values = [75], FillColor = "#222222" }
            ]
        };

        var elements = BarChartElementBuilder.Build(model, 0, 0, 240, 160);
        var bars = elements.Where(e => e.Shape?.FillColor is "#111111" or "#222222").ToList();

        Assert.Equal(2, bars.Count);
        Assert.Equal(bars[0].X, bars[1].X);
        Assert.Equal(bars[0].Width, bars[1].Width);
        Assert.Equal(3, Math.Round(bars[1].Height / bars[0].Height));
    }
}
