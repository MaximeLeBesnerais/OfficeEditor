using PptxEditor.Core.Converters.Charts;
using Xunit;

namespace DocxEditor.Tests.Unit.Charts;

/// <summary>
/// Model parsing from synthetic chart-part XML: series names from <c>c:tx</c> caches,
/// categories/values from cached points, explicit series colours, scheme-colour
/// resolution, axes, legend and data-label settings. References (<c>c:f</c>) must never
/// be evaluated — the parser reads caches only.
/// </summary>
public sealed class ChartPartParserTests
{
    private const string Ns =
        "xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" " +
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";

    private static string ChartSpace(string plotAreaContent, string chartTrailer = "") =>
        $"<c:chartSpace {Ns}><c:chart><c:plotArea><c:layout/>{plotAreaContent}</c:plotArea>{chartTrailer}</c:chart></c:chartSpace>";

    private static string Series(int idx, string name, string fill, string cats, string vals, string extra = "") =>
        $"<c:ser><c:idx val=\"{idx}\"/><c:order val=\"{idx}\"/>" +
        $"<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val=\"1\"/>" +
        $"<c:pt idx=\"0\"><c:v>{name}</c:v></c:pt></c:strCache></c:strRef></c:tx>" +
        (fill.Length > 0 ? $"<c:spPr><a:solidFill>{fill}</a:solidFill></c:spPr>" : "") +
        extra +
        $"<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f><c:strCache>{cats}</c:strCache></c:strRef></c:cat>" +
        $"<c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f><c:numCache><c:formatCode>General</c:formatCode>{vals}" +
        "</c:numCache></c:numRef></c:val></c:ser>";

    [Fact]
    public void Parse_ClusteredBarChart_ReadsSeriesNamesCategoriesAndValuesFromCaches()
    {
        var xml = ChartSpace(
            "<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>" +
            Series(0, "Alpha", "<a:srgbClr val=\"112233\"/>",
                "<c:pt idx=\"0\"><c:v>Q1</c:v></c:pt><c:pt idx=\"1\"><c:v>Q2</c:v></c:pt><c:pt idx=\"2\"><c:v>Q3</c:v></c:pt>",
                "<c:pt idx=\"0\"><c:v>10</c:v></c:pt><c:pt idx=\"1\"><c:v>20</c:v></c:pt><c:pt idx=\"2\"><c:v>30</c:v></c:pt>") +
            Series(1, "Beta", "<a:srgbClr val=\"AABBCC\"/>",
                "<c:pt idx=\"0\"><c:v>Q1</c:v></c:pt><c:pt idx=\"1\"><c:v>Q2</c:v></c:pt><c:pt idx=\"2\"><c:v>Q3</c:v></c:pt>",
                "<c:pt idx=\"0\"><c:v>5</c:v></c:pt><c:pt idx=\"1\"><c:v>-7.5</c:v></c:pt><c:pt idx=\"2\"><c:v>15</c:v></c:pt>") +
            "</c:barChart>");

        var model = ChartPartParser.Parse(xml);

        Assert.NotNull(model);
        Assert.Equal(ChartKind.Bar, model.Kind);
        Assert.Equal(BarDirection.Column, model.Direction);
        Assert.Equal(BarGrouping.Clustered, model.Grouping);
        Assert.Equal(150.0, model.GapWidthPercent); // default when c:gapWidth absent

        Assert.Equal(2, model.Series.Count);
        Assert.Equal("Alpha", model.Series[0].Name);
        Assert.Equal("#112233", model.Series[0].FillColor);
        Assert.Equal(new double?[] { 10, 20, 30 }, model.Series[0].Values);
        Assert.Equal("Beta", model.Series[1].Name);
        Assert.Equal("#AABBCC", model.Series[1].FillColor);
        Assert.Equal(new double?[] { 5, -7.5, 15 }, model.Series[1].Values);

        Assert.Equal(new[] { "Q1", "Q2", "Q3" }, model.Categories);
    }

    [Fact]
    public void Parse_HorizontalBarChart_ReadsDirectionGroupingGapAndOverlap()
    {
        var xml = ChartSpace(
            "<c:barChart><c:barDir val=\"bar\"/><c:grouping val=\"clustered\"/>" +
            Series(0, "S", "<a:srgbClr val=\"C00000\"/>",
                "<c:pt idx=\"0\"><c:v>A</c:v></c:pt>", "<c:pt idx=\"0\"><c:v>1</c:v></c:pt>") +
            "<c:gapWidth val=\"60\"/><c:overlap val=\"-20\"/></c:barChart>");

        var model = ChartPartParser.Parse(xml);

        Assert.NotNull(model);
        Assert.Equal(BarDirection.Bar, model.Direction);
        Assert.Equal(60.0, model.GapWidthPercent);
        Assert.Equal(-20.0, model.OverlapPercent);
    }

    [Fact]
    public void Parse_SchemeColorFill_ResolvesThroughTheme()
    {
        var xml = ChartSpace(
            "<c:barChart><c:barDir val=\"col\"/>" +
            Series(0, "S", "<a:schemeClr val=\"accent2\"/>",
                "<c:pt idx=\"0\"><c:v>A</c:v></c:pt>", "<c:pt idx=\"0\"><c:v>1</c:v></c:pt>") +
            "</c:barChart>");

        var model = ChartPartParser.Parse(xml, name => name == "accent2" ? "#ED7D31" : null);

        Assert.NotNull(model);
        Assert.Equal("#ED7D31", model.Series[0].FillColor);
    }

    [Fact]
    public void Parse_MissingValuePoint_LeavesNullSlot()
    {
        var xml = ChartSpace(
            "<c:barChart><c:barDir val=\"col\"/>" +
            Series(0, "S", "<a:srgbClr val=\"112233\"/>",
                "<c:pt idx=\"0\"><c:v>A</c:v></c:pt><c:pt idx=\"1\"><c:v>B</c:v></c:pt><c:pt idx=\"2\"><c:v>C</c:v></c:pt>",
                // idx 1 deliberately missing (blank cell cached as a gap).
                "<c:pt idx=\"0\"><c:v>1</c:v></c:pt><c:pt idx=\"2\"><c:v>3</c:v></c:pt>") +
            "</c:barChart>");

        var model = ChartPartParser.Parse(xml);

        Assert.NotNull(model);
        Assert.Equal(new double?[] { 1, null, 3 }, model.Series[0].Values);
    }

    [Fact]
    public void Parse_Axes_ReadBoundsGridlinesLabelsAndNumberFormat()
    {
        var axisXml =
            "<c:catAx><c:axId val=\"1\"/><c:scaling/><c:delete val=\"0\"/>" +
            "<c:spPr><a:ln><a:solidFill><a:srgbClr val=\"9FB3D9\"/></a:solidFill></a:ln></c:spPr>" +
            "<c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz=\"900\">" +
            "<a:solidFill><a:srgbClr val=\"747474\"/></a:solidFill><a:latin typeface=\"Verdana\"/>" +
            "</a:defRPr></a:pPr></a:p></c:txPr></c:catAx>" +
            "<c:valAx><c:axId val=\"2\"/><c:scaling><c:max val=\"200\"/><c:min val=\"-50\"/></c:scaling>" +
            "<c:delete val=\"0\"/>" +
            "<c:majorGridlines><c:spPr><a:ln><a:solidFill><a:srgbClr val=\"D7D7D7\"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>" +
            "<c:numFmt formatCode=\"0%\" sourceLinked=\"0\"/><c:majorUnit val=\"50\"/>" +
            "<c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz=\"1000\">" +
            "<a:solidFill><a:srgbClr val=\"222222\"/></a:solidFill><a:latin typeface=\"Calibri\"/>" +
            "</a:defRPr></a:pPr></a:p></c:txPr></c:valAx>";

        var xml = ChartSpace(
            "<c:barChart><c:barDir val=\"col\"/>" +
            Series(0, "S", "<a:srgbClr val=\"112233\"/>",
                "<c:pt idx=\"0\"><c:v>A</c:v></c:pt>", "<c:pt idx=\"0\"><c:v>1</c:v></c:pt>") +
            "</c:barChart>" + axisXml);

        var model = ChartPartParser.Parse(xml);

        Assert.NotNull(model);
        Assert.NotNull(model.CategoryAxis);
        Assert.False(model.CategoryAxis.Deleted);
        Assert.Equal("#9FB3D9", model.CategoryAxis.LineColor);
        Assert.Equal(9.0, model.CategoryAxis.LabelFontSize);
        Assert.Equal("#747474", model.CategoryAxis.LabelColor);
        Assert.Equal("Verdana", model.CategoryAxis.LabelFontFamily);

        Assert.NotNull(model.ValueAxis);
        Assert.Equal(200.0, model.ValueAxis.Max);
        Assert.Equal(-50.0, model.ValueAxis.Min);
        Assert.Equal(50.0, model.ValueAxis.MajorUnit);
        Assert.Equal("#D7D7D7", model.ValueAxis.GridlineColor);
        Assert.Equal("0%", model.ValueAxis.NumberFormat);
        Assert.Equal(10.0, model.ValueAxis.LabelFontSize);
        Assert.Equal("#222222", model.ValueAxis.LabelColor);
        Assert.Equal("Calibri", model.ValueAxis.LabelFontFamily);
    }

    [Fact]
    public void Parse_Legend_ReadsPositionAndDefaultsToRightWhenAbsent()
    {
        var withPos = ChartSpace(
            "<c:barChart><c:barDir val=\"col\"/>" +
            Series(0, "S", "<a:srgbClr val=\"112233\"/>",
                "<c:pt idx=\"0\"><c:v>A</c:v></c:pt>", "<c:pt idx=\"0\"><c:v>1</c:v></c:pt>") +
            "</c:barChart>",
            "<c:legend><c:legendPos val=\"b\"/></c:legend>");
        var withoutPos = ChartSpace(
            "<c:barChart><c:barDir val=\"col\"/>" +
            Series(0, "S", "<a:srgbClr val=\"112233\"/>",
                "<c:pt idx=\"0\"><c:v>A</c:v></c:pt>", "<c:pt idx=\"0\"><c:v>1</c:v></c:pt>") +
            "</c:barChart>",
            "<c:legend><c:overlay val=\"0\"/></c:legend>");
        var noLegend = ChartSpace(
            "<c:barChart><c:barDir val=\"col\"/>" +
            Series(0, "S", "<a:srgbClr val=\"112233\"/>",
                "<c:pt idx=\"0\"><c:v>A</c:v></c:pt>", "<c:pt idx=\"0\"><c:v>1</c:v></c:pt>") +
            "</c:barChart>");

        Assert.Equal(ChartLegendPosition.Bottom, ChartPartParser.Parse(withPos)?.Legend?.Position);
        Assert.Equal(ChartLegendPosition.Right, ChartPartParser.Parse(withoutPos)?.Legend?.Position);
        Assert.Null(ChartPartParser.Parse(noLegend)?.Legend);
    }

    [Fact]
    public void Parse_DataLabels_ReadShowValueAndTextProperties()
    {
        var dLbls =
            "<c:dLbls><c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz=\"900\">" +
            "<a:solidFill><a:srgbClr val=\"1B1B1B\"/></a:solidFill><a:latin typeface=\"Verdana\"/>" +
            "</a:defRPr></a:pPr></a:p></c:txPr>" +
            "<c:showLegendKey val=\"0\"/><c:showVal val=\"1\"/><c:showCatName val=\"0\"/></c:dLbls>";

        var xml = ChartSpace(
            "<c:barChart><c:barDir val=\"col\"/>" +
            Series(0, "S", "<a:srgbClr val=\"112233\"/>",
                "<c:pt idx=\"0\"><c:v>A</c:v></c:pt>", "<c:pt idx=\"0\"><c:v>1</c:v></c:pt>", extra: dLbls) +
            dLbls +
            "</c:barChart>");

        var model = ChartPartParser.Parse(xml);

        Assert.NotNull(model);
        Assert.NotNull(model.DataLabels);
        Assert.True(model.DataLabels.ShowValue);
        Assert.Equal(9.0, model.DataLabels.FontSize);
        Assert.Equal("#1B1B1B", model.DataLabels.Color);
        Assert.Equal("Verdana", model.DataLabels.FontFamily);
        Assert.NotNull(model.Series[0].DataLabels);
        Assert.True(model.Series[0].DataLabels!.ShowValue);
    }

    [Fact]
    public void Parse_NonBarCharts_ClassifyKindForFallbackRouting()
    {
        var line = ChartSpace("<c:lineChart><c:grouping val=\"standard\"/></c:lineChart>");
        var pie = ChartSpace("<c:pieChart><c:varyColors val=\"1\"/></c:pieChart>");
        var area = ChartSpace("<c:areaChart><c:grouping val=\"standard\"/></c:areaChart>");

        Assert.Equal(ChartKind.Line, ChartPartParser.Parse(line)?.Kind);
        Assert.Equal(ChartKind.Pie, ChartPartParser.Parse(pie)?.Kind);
        Assert.Equal(ChartKind.Other, ChartPartParser.Parse(area)?.Kind);
    }

    [Fact]
    public void Parse_GarbageOrForeignXml_ReturnsNull()
    {
        Assert.Null(ChartPartParser.Parse(""));
        Assert.Null(ChartPartParser.Parse("   "));
        Assert.Null(ChartPartParser.Parse("not xml at all"));
        Assert.Null(ChartPartParser.Parse("<root/>"));
        Assert.Null(ChartPartParser.Parse("<c:chartSpace " + Ns.Replace('\"', '\'') + "/>"));
    }
}
