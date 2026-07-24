using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PptxEditor.Core.Converters;
using Xunit;

namespace DocxEditor.Tests.Unit.Charts;

/// <summary>
/// End-to-end routing of chart graphicFrames through the chart pipeline: clustered bar
/// charts decompose into primitives, unsupported chart types (line/pie) keep the visible
/// placeholder with an accurate warning, and the reference deck's slide 9 chart converts
/// without a placeholder.
/// </summary>
public sealed class ChartConversionIntegrationTests : IDisposable
{
    private const long EmusPerPoint = 12700;

    private const string ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private const string ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private readonly string _tempDir;

    public ChartConversionIntegrationTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), nameof(ChartConversionIntegrationTests), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose() => SlideOpsTestHelpers.BestEffortDelete(_tempDir);

    private const string ClusteredBarChartXml =
        $"<c:chartSpace xmlns:c=\"{ChartNs}\" xmlns:a=\"{ANs}\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
        "<c:chart><c:plotArea><c:layout/>" +
        "<c:barChart><c:barDir val=\"bar\"/><c:grouping val=\"clustered\"/><c:varyColors val=\"1\"/>" +
        "<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>" +
        "<c:tx><c:strRef><c:strCache><c:pt idx=\"0\"><c:v>Benchmark</c:v></c:pt></c:strCache></c:strRef></c:tx>" +
        "<c:spPr><a:solidFill><a:srgbClr val=\"B6B5B5\"/></a:solidFill></c:spPr>" +
        "<c:cat><c:strRef><c:strCache><c:pt idx=\"0\"><c:v>Win rate</c:v></c:pt><c:pt idx=\"1\"><c:v>Deal speed</c:v></c:pt></c:strCache></c:strRef></c:cat>" +
        "<c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode><c:pt idx=\"0\"><c:v>100</c:v></c:pt><c:pt idx=\"1\"><c:v>100</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>" +
        "<c:ser><c:idx val=\"1\"/><c:order val=\"1\"/>" +
        "<c:tx><c:strRef><c:strCache><c:pt idx=\"0\"><c:v>Today</c:v></c:pt></c:strCache></c:strRef></c:tx>" +
        "<c:spPr><a:solidFill><a:srgbClr val=\"C00000\"/></a:solidFill></c:spPr>" +
        "<c:cat><c:strRef><c:strCache><c:pt idx=\"0\"><c:v>Win rate</c:v></c:pt><c:pt idx=\"1\"><c:v>Deal speed</c:v></c:pt></c:strCache></c:strRef></c:cat>" +
        "<c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode><c:pt idx=\"0\"><c:v>58</c:v></c:pt><c:pt idx=\"1\"><c:v>74</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>" +
        "<c:gapWidth val=\"60\"/><c:overlap val=\"-20\"/></c:barChart>" +
        "<c:catAx><c:axId val=\"1\"/><c:scaling/><c:delete val=\"0\"/></c:catAx>" +
        "<c:valAx><c:axId val=\"2\"/><c:scaling/><c:delete val=\"0\"/>" +
        "<c:majorGridlines><c:spPr><a:ln><a:solidFill><a:srgbClr val=\"D7D7D7\"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>" +
        "</c:valAx></c:plotArea>" +
        "<c:legend><c:legendPos val=\"r\"/></c:legend>" +
        "</c:chart></c:chartSpace>";

    private const string LineChartXml =
        $"<c:chartSpace xmlns:c=\"{ChartNs}\" xmlns:a=\"{ANs}\">" +
        "<c:chart><c:plotArea><c:layout/>" +
        "<c:lineChart><c:grouping val=\"standard\"/>" +
        "<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>" +
        "<c:cat><c:strRef><c:strCache><c:pt idx=\"0\"><c:v>A</c:v></c:pt></c:strCache></c:strRef></c:cat>" +
        "<c:val><c:numRef><c:numCache><c:pt idx=\"0\"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>" +
        "</c:lineChart></c:plotArea></c:chart></c:chartSpace>";

    [Fact]
    public void Convert_ClusteredBarChart_DecomposesIntoPrimitivesWithoutPlaceholder()
    {
        var path = CreateDeckWithChart(ClusteredBarChartXml, "Sales chart");
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var slide = Assert.Single(converter.Convert().Slides);

        // Bars from both series with their explicit colours.
        Assert.Contains(slide.Elements, e => e.Type == "Shape" && e.Shape?.FillColor == "#B6B5B5");
        Assert.Contains(slide.Elements, e => e.Type == "Shape" && e.Shape?.FillColor == "#C00000");
        // Category + legend + value labels.
        Assert.Contains(slide.Elements, e => e.Text?.Content == "Win rate");
        Assert.Contains(slide.Elements, e => e.Text?.Content == "Benchmark");
        Assert.Contains(slide.Elements, e => e.Text?.Content == "100");
        // No placeholder, no warning.
        Assert.DoesNotContain(slide.Elements,
            e => e.Text?.Content.Contains("not supported", StringComparison.Ordinal) == true);
        Assert.Empty(slide.Warnings);

        // The emitted Typst source places real rects, not the placeholder label.
        var source = converter.GenerateTypstSource(converter.Convert());
        Assert.Contains("rgb(\"#C00000\")", source);
        Assert.DoesNotContain("Chart (not supported)", source);
    }

    [Fact]
    public void Convert_LineChart_KeepsPlaceholderWithAccurateWarning()
    {
        var path = CreateDeckWithChart(LineChartXml, "Trend chart");
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var slide = Assert.Single(converter.Convert().Slides);

        Assert.Contains(slide.Elements,
            e => e.Text?.Content.Contains("Chart (not supported)", StringComparison.Ordinal) == true);
        var warning = Assert.Single(slide.Warnings);
        Assert.Contains("line chart", warning);
        Assert.Contains("Trend chart", warning);
        Assert.Contains("placeholder", warning);
    }

    [Fact]
    public void Convert_ReferenceDeck_Slide9ChartConvertsWithoutPlaceholder()
    {
        var referencePath = Path.Combine(
            Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "examples", "REF", "PPTX")),
            "sales_acceleration_deck.pptx");
        Assert.True(File.Exists(referencePath), $"Reference deck not found: {referencePath}");

        using var document = PresentationDocument.Open(referencePath, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var slide9 = presentation.Slides[8];

        // The chart frame is converted: both series colours appear as bar fills and the
        // category axis labels are present.
        Assert.Contains(slide9.Elements, e => e.Type == "Shape" && e.Shape?.FillColor == "#B6B5B5");
        Assert.Contains(slide9.Elements, e => e.Type == "Shape" && e.Shape?.FillColor == "#C00000");
        Assert.Contains(slide9.Elements, e => e.Text?.Content == "Lead conversion");
        Assert.Contains(slide9.Elements, e => e.Text?.Content == "Best-in-class");
        Assert.DoesNotContain(slide9.Elements,
            e => e.Text?.Content.Contains("Chart (not supported)", StringComparison.Ordinal) == true);
        Assert.DoesNotContain(slide9.Warnings, w => w.Contains("Chart", StringComparison.Ordinal));
    }

    private string CreateDeckWithChart(string chartXml, string frameName)
    {
        var path = Path.Combine(_tempDir, $"{Guid.NewGuid():N}.pptx");

        using (var document = PresentationDocument.Create(path, PresentationDocumentType.Presentation))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new Presentation
            {
                SlideMasterIdList = new SlideMasterIdList(),
                SlideIdList = new SlideIdList(),
                SlideSize = new SlideSize { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 }
            };

            var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
            slideMasterPart.SlideMaster = new SlideMaster(
                new CommonSlideData(new ShapeTree(
                    new NonVisualGroupShapeProperties(
                        new NonVisualDrawingProperties { Id = 1, Name = "" },
                        new NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new Drawing.TransformGroup()))),
                new ColorMap
                {
                    Background1 = Drawing.ColorSchemeIndexValues.Light1,
                    Text1 = Drawing.ColorSchemeIndexValues.Dark1,
                    Background2 = Drawing.ColorSchemeIndexValues.Light2,
                    Text2 = Drawing.ColorSchemeIndexValues.Dark2,
                    Accent1 = Drawing.ColorSchemeIndexValues.Accent1,
                    Accent2 = Drawing.ColorSchemeIndexValues.Accent2,
                    Accent3 = Drawing.ColorSchemeIndexValues.Accent3,
                    Accent4 = Drawing.ColorSchemeIndexValues.Accent4,
                    Accent5 = Drawing.ColorSchemeIndexValues.Accent5,
                    Accent6 = Drawing.ColorSchemeIndexValues.Accent6,
                    Hyperlink = Drawing.ColorSchemeIndexValues.Hyperlink,
                    FollowedHyperlink = Drawing.ColorSchemeIndexValues.FollowedHyperlink
                },
                new SlideLayoutIdList());

            var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
            slideLayoutPart.SlideLayout = new P.SlideLayout(new CommonSlideData(new ShapeTree(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = 1, Name = "" },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(new Drawing.TransformGroup()))));
            slideLayoutPart.AddPart(slideMasterPart);
            slideMasterPart.SlideMaster.SlideLayoutIdList!.Append(new SlideLayoutId
            {
                Id = 2147483649,
                RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart)
            });

            presentationPart.Presentation.SlideMasterIdList.Append(new SlideMasterId
            {
                Id = 2147483648,
                RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
            });

            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.Slide = new Slide(new CommonSlideData(new ShapeTree(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = 1, Name = "" },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(new Drawing.TransformGroup()))));
            slidePart.AddPart(slideLayoutPart);

            var chartPart = slidePart.AddNewPart<ChartPart>();
            using (var stream = chartPart.GetStream(FileMode.Create))
            using (var writer = new StreamWriter(stream))
            {
                writer.Write(chartXml);
            }

            var relationshipId = slidePart.GetIdOfPart(chartPart);
            slidePart.Slide.CommonSlideData!.ShapeTree!.Append(new P.GraphicFrame(
                new NonVisualGraphicFrameProperties(
                    new NonVisualDrawingProperties { Id = 10, Name = frameName },
                    new NonVisualGraphicFrameDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new P.Transform(
                    new Drawing.Offset { X = 60 * EmusPerPoint, Y = 110 * EmusPerPoint },
                    new Drawing.Extents { Cx = 600 * EmusPerPoint, Cy = 330 * EmusPerPoint }),
                new Drawing.Graphic(new Drawing.GraphicData(
                    new Drawing.Charts.ChartReference { Id = relationshipId })
                {
                    Uri = ChartNs
                })));

            presentationPart.Presentation.SlideIdList.Append(new SlideId
            {
                Id = 256,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });
        }

        return path;
    }
}
