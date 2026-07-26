using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PptxEditor.Core.Converters;
using PptxEditor.Core.Converters.Charts;
using Xunit;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Regression tests for defects found on the "Space" deck :
///
/// 1. Layout pictures must resolve their image relationships against the LAYOUT
///    part. Relationship IDs are part-scoped; when the slide owns an image part
///    with the same relationship ID as a layout picture's blip embed, the
///    converter picked the slide's image (e.g. the iMac frame rendered as a
///    slide-content screenshot).
///
/// 2. A slide-level <c>p:pic</c> placeholder without an explicit
///    <c>a:xfrm</c> must inherit its position from the matching layout
///    placeholder (ph type="pic"). Previously it fell back to the default
///    (0,0,100x50pt), rendering the picture tiny at the top-left corner.
/// </summary>
public sealed class PptxToTypstConverterSpaceTests : IDisposable
{
    private readonly string _tempDir;

    public PptxToTypstConverterSpaceTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), nameof(PptxToTypstConverterSpaceTests), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        SlideOpsTestHelpers.BestEffortDelete(_tempDir);
    }

    [Fact]
    public void Convert_LayoutPicture_ResolvesImageAgainstLayoutPartNotSlidePart()
    {
        var path = CreateDeckWithCollidingImageRelIds(out var layoutPixelSize);
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();

        var layoutImage = Assert.Single(presentation.Slides[0].Elements,
            e => e.Type == "Image" && e.Name == "Layout Picture");
        Assert.NotNull(layoutImage.Image);
        Assert.Equal(layoutPixelSize, layoutImage.Image!.PixelWidth);
    }

    [Fact]
    public void Convert_SlidePicturePlaceholderWithoutXfrm_InheritsLayoutPlaceholderPosition()
    {
        var path = CreateDeckWithPicturePlaceholder();
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();

        var image = Assert.Single(presentation.Slides[0].Elements, e => e.Type == "Image");
        // Layout pic placeholder xfrm: off (1000000, 800000), ext (3000000, 1500000) EMU.
        Assert.Equal(1000000 / 12700.0, image.X, 2);
        Assert.Equal(800000 / 12700.0, image.Y, 2);
        Assert.Equal(3000000 / 12700.0, image.Width, 2);
        Assert.Equal(1500000 / 12700.0, image.Height, 2);
    }

    /// <summary>
    /// Deck where the slide layout has a picture (embed rel ID collides with an
    /// image rel ID on the slide). Both relationship collections start at rId1,
    /// so the layout picture's embed ID also exists on the slide part — pointing
    /// at a DIFFERENT image. The layout picture must render the layout's image.
    /// </summary>
    private string CreateDeckWithCollidingImageRelIds(out int layoutPixelSize)
    {
        var deckPath = Path.Combine(_tempDir, $"{Guid.NewGuid():N}.pptx");
        var layoutPng = WritePng(Path.Combine(_tempDir, "layout.png"), Png2x1);
        var slidePng = SlideOpsTestHelpers.WriteMinimalPng(_tempDir); // 1x1
        layoutPixelSize = 2;

        using (var document = PresentationDocument.Create(deckPath, PresentationDocumentType.Presentation))
        {
            var (presentationPart, slideLayoutPart) = CreateShell(document);

            // Layout image part + layout picture referencing it.
            var layoutImagePart = slideLayoutPart.AddNewPart<ImagePart>("image/png", "rIdColliding");
            using (var s = File.OpenRead(layoutPng))
            {
                layoutImagePart.FeedData(s);
            }

            var layoutEmbedId = slideLayoutPart.GetIdOfPart(layoutImagePart);
            slideLayoutPart.SlideLayout = new P.SlideLayout(new CommonSlideData(CreateShapeTree(
                PlainPicture("Layout Picture", layoutEmbedId, 1000000, 500000, 2540000, 1270000))));

            // Slide with its own image (different bytes/dimensions). The slide's
            // rel ID for this image collides with the layout picture's embed ID
            // because both parts number relationships independently.
            var slidePart = presentationPart.AddNewPart<SlidePart>();
            var slideImagePart = slidePart.AddNewPart<ImagePart>("image/png", "rIdColliding");
            using (var s = File.OpenRead(slidePng))
            {
                slideImagePart.FeedData(s);
            }

            var slideEmbedId = slidePart.GetIdOfPart(slideImagePart);
            Assert.Equal(layoutEmbedId, slideEmbedId); // pre-condition: IDs really collide

            slidePart.Slide = new Slide(new CommonSlideData(CreateShapeTree(
                PlainPicture("Slide Picture", slideEmbedId, 5000000, 3000000, 1000000, 1000000))));
            slidePart.AddPart(slideLayoutPart);

            presentationPart.Presentation!.SlideIdList!.Append(new SlideId
            {
                Id = 256,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });
        }

        return deckPath;
    }

    // ------------------------------------------------------------------
    // Pie chart support ( 15 rendered a blank placeholder box)
    // ------------------------------------------------------------------

    [Fact]
    public void Parse_PieChart_ReadsPerPointFillsAndFirstSliceAngle()
    {
        var model = ChartPartParser.Parse(PieChartXml);

        Assert.NotNull(model);
        Assert.Equal(ChartKind.Pie, model!.Kind);
        Assert.Equal(45.0, model.FirstSliceAngleDegrees);

        var series = Assert.Single(model.Series);
        Assert.Equal([8.2, 3.2, 1.4, 1.2], series.Values.Select(v => v!.Value).ToArray());
        Assert.Equal(["#009CBB", "#57A166", "#C83A50", "#FEC60E"], series.PointFillColors);
        Assert.Equal(1.5, series.PointLineWidthPt, 3);
        Assert.NotNull(series.PointLineColor);
    }

    [Fact]
    public void Build_PieChart_EmitsOneWedgePolygonPerSlice()
    {
        var model = new ChartModel
        {
            Kind = ChartKind.Pie,
            Categories = ["A", "B", "C", "D"],
            Series =
            [
                new ChartSeries
                {
                    Name = "Sales",
                    Values = [8.2, 3.2, 1.4, 1.2],
                    PointFillColors = ["#009CBB", "#57A166", "#C83A50", "#FEC60E"],
                    PointLineColor = "#FFFFFF",
                    PointLineWidthPt = 1.5
                }
            ]
        };

        var elements = PieChartElementBuilder.Build(model, 100, 50, 240, 200);

        var wedges = elements.Where(e => e.Type == "Shape" && e.Shape?.ShapeType == "polygon").ToList();
        Assert.Equal(4, wedges.Count);
        Assert.Equal(["#009CBB", "#57A166", "#C83A50", "#FEC60E"],
            wedges.Select(w => w.Shape!.FillColor).ToArray());
        Assert.All(wedges, w => Assert.Equal("#FFFFFF", w.Shape!.StrokeColor));

        // Every wedge point must be normalized to the 0..1 element box.
        Assert.All(wedges.SelectMany(w => w.Shape!.Points),
            p => Assert.InRange(p.X, 0, 1));
        Assert.All(wedges.SelectMany(w => w.Shape!.Points),
            p => Assert.InRange(p.Y, 0, 1));

        // First slice starts at 12 o'clock (firstSliceAng=0): the largest slice
        // must include the top-center point of the pie circle.
        var first = wedges[0];
        var topCenter = first.Shape!.Points
            .Where(p => Math.Abs(p.X - 0.5) < 0.02)
            .OrderBy(p => p.Y)
            .FirstOrDefault();
        Assert.True(topCenter.Y < 0.15, "first slice should touch the top of the pie");
    }

    [Fact]
    public void Build_PieChart_SingleValue_EmitsFullEllipse()
    {
        var model = new ChartModel
        {
            Kind = ChartKind.Pie,
            Categories = ["A"],
            Series = [new ChartSeries { Name = "S", Values = [5.0], PointFillColors = ["#009CBB"] }]
        };

        var elements = PieChartElementBuilder.Build(model, 0, 0, 200, 200);

        var ellipse = Assert.Single(elements, e => e.Type == "Shape" && e.Shape?.ShapeType == "ellipse");
        Assert.Equal("#009CBB", ellipse.Shape!.FillColor);
    }

    /// <summary>Synthetic pie chart part mirroring  15's chart1.xml.</summary>
    private const string PieChartXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <c:chart>
            <c:autoTitleDeleted val="1"/>
            <c:plotArea>
              <c:layout/>
              <c:pieChart>
                <c:varyColors val="1"/>
                <c:ser>
                  <c:idx val="0"/>
                  <c:order val="0"/>
                  <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Sales</c:v></c:pt></c:strCache></c:strRef></c:tx>
                  <c:dPt>
                    <c:idx val="0"/>
                    <c:spPr>
                      <a:solidFill><a:srgbClr val="009CBB"/></a:solidFill>
                      <a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>
                    </c:spPr>
                  </c:dPt>
                  <c:dPt>
                    <c:idx val="1"/>
                    <c:spPr>
                      <a:solidFill><a:srgbClr val="57A166"/></a:solidFill>
                      <a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>
                    </c:spPr>
                  </c:dPt>
                  <c:dPt>
                    <c:idx val="2"/>
                    <c:spPr>
                      <a:solidFill><a:srgbClr val="C83A50"/></a:solidFill>
                      <a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>
                    </c:spPr>
                  </c:dPt>
                  <c:dPt>
                    <c:idx val="3"/>
                    <c:spPr>
                      <a:solidFill><a:srgbClr val="FEC60E"/></a:solidFill>
                      <a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>
                    </c:spPr>
                  </c:dPt>
                  <c:cat><c:strRef><c:strCache>
                    <c:pt idx="0"><c:v>1st Qtr</c:v></c:pt>
                    <c:pt idx="1"><c:v>2nd Qtr</c:v></c:pt>
                    <c:pt idx="2"><c:v>3rd Qtr</c:v></c:pt>
                    <c:pt idx="3"><c:v>4th Qtr</c:v></c:pt>
                  </c:strCache></c:strRef></c:cat>
                  <c:val><c:numRef><c:numCache>
                    <c:pt idx="0"><c:v>8.2</c:v></c:pt>
                    <c:pt idx="1"><c:v>3.2</c:v></c:pt>
                    <c:pt idx="2"><c:v>1.4</c:v></c:pt>
                    <c:pt idx="3"><c:v>1.2</c:v></c:pt>
                  </c:numCache></c:numRef></c:val>
                </c:ser>
                <c:firstSliceAng val="45"/>
              </c:pieChart>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
        """;

    /// <summary>
    /// Deck whose slide has a pic placeholder (ph type="pic" idx="10") with an
    /// EMPTY spPr (no xfrm). The layout carries the matching pic placeholder
    /// shape with an explicit xfrm that the slide picture must inherit.
    /// </summary>
    private string CreateDeckWithPicturePlaceholder()
    {
        var deckPath = Path.Combine(_tempDir, $"{Guid.NewGuid():N}.pptx");
        var pngPath = SlideOpsTestHelpers.WriteMinimalPng(_tempDir);

        using (var document = PresentationDocument.Create(deckPath, PresentationDocumentType.Presentation))
        {
            var (presentationPart, slideLayoutPart) = CreateShell(document);

            slideLayoutPart.SlideLayout = new P.SlideLayout(new CommonSlideData(CreateShapeTree(
                new P.Shape(
                    new NonVisualShapeProperties(
                        new NonVisualDrawingProperties { Id = 2, Name = "Picture Placeholder" },
                        new NonVisualShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties(
                            new PlaceholderShape { Type = PlaceholderValues.Picture, Index = 10 })),
                    new ShapeProperties(
                        new Drawing.Transform2D(
                            new Drawing.Offset { X = 1000000, Y = 800000 },
                            new Drawing.Extents { Cx = 3000000, Cy = 1500000 })),
                    new P.TextBody(
                        new Drawing.BodyProperties(),
                        new Drawing.ListStyle(),
                        new Drawing.Paragraph(new Drawing.EndParagraphRunProperties()))))));

            var slidePart = presentationPart.AddNewPart<SlidePart>();
            var imagePart = slidePart.AddImagePart(ImagePartType.Png);
            using (var s = File.OpenRead(pngPath))
            {
                imagePart.FeedData(s);
            }

            var embedId = slidePart.GetIdOfPart(imagePart);
            slidePart.Slide = new Slide(new CommonSlideData(CreateShapeTree(
                new P.Picture(
                    new P.NonVisualPictureProperties(
                        new NonVisualDrawingProperties { Id = 3, Name = "Picture Placeholder 6" },
                        new P.NonVisualPictureDrawingProperties(
                            new Drawing.PictureLocks { NoGrouping = true }),
                        new ApplicationNonVisualDrawingProperties(
                            new PlaceholderShape { Type = PlaceholderValues.Picture, Index = 10 })),
                    new P.BlipFill(
                        new Drawing.Blip { Embed = embedId },
                        new Drawing.Stretch(new Drawing.FillRectangle())),
                    new ShapeProperties()))));
            slidePart.AddPart(slideLayoutPart);

            presentationPart.Presentation!.SlideIdList!.Append(new SlideId
            {
                Id = 256,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });
        }

        return deckPath;
    }

    private static (PresentationPart PresentationPart, SlideLayoutPart SlideLayoutPart) CreateShell(
        PresentationDocument document)
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
            new CommonSlideData(CreateShapeTree()),
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

        return (presentationPart, slideLayoutPart);
    }

    private static P.Picture PlainPicture(string name, string embedId, long x, long y, long cx, long cy)
    {
        return new P.Picture(
            new P.NonVisualPictureProperties(
                new NonVisualDrawingProperties { Id = 10, Name = name },
                new P.NonVisualPictureDrawingProperties(
                    new Drawing.PictureLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties()),
            new P.BlipFill(
                new Drawing.Blip { Embed = embedId },
                new Drawing.Stretch(new Drawing.FillRectangle())),
            new ShapeProperties(
                new Drawing.Transform2D(
                    new Drawing.Offset { X = x, Y = y },
                    new Drawing.Extents { Cx = cx, Cy = cy })));
    }

    private static string WritePng(string path, string base64)
    {
        File.WriteAllBytes(path, Convert.FromBase64String(base64));
        return path;
    }

    // Minimal 2x1 PNG (distinct dimensions from the 1x1 helper PNG).
    private const string Png2x1 =
        "iVBORw0KGgoAAAANSUhEUgAAAAIAAAABAQMAAADO7O3JAAAAA1BMVEX/AAAZ4gk3AAAACklEQVQI12NgAAAAAgAB4iG8MwAAAABJRU5ErkJggg==";

    private static ShapeTree CreateShapeTree(params OpenXmlElement[] elements)
    {
        var shapeTree = new ShapeTree(
            new NonVisualGroupShapeProperties(
                new NonVisualDrawingProperties { Id = 0, Name = "" },
                new NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(
                new Drawing.TransformGroup(
                    new Drawing.Offset { X = 0, Y = 0 },
                    new Drawing.Extents { Cx = 0, Cy = 0 },
                    new Drawing.ChildOffset { X = 0, Y = 0 },
                    new Drawing.ChildExtents { Cx = 0, Cy = 0 })));

        foreach (var element in elements)
        {
            shapeTree.Append(element);
        }

        return shapeTree;
    }
}
