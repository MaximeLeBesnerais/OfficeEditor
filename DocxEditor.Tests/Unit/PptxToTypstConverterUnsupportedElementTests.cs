using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PptxEditor.Core.Converters;
using Xunit;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Tests for the unsupported-graphic-frame policy: charts and other unconvertible graphic
/// frames must render a visible placeholder (framed box + label) instead of being silently
/// dropped, and must be reported on <see cref="PptxEditor.Core.Models.TypstSlide.Warnings"/>.
/// SmartArt keeps its text approximation and is reported as approximated.
/// </summary>
public sealed class PptxToTypstConverterUnsupportedElementTests : IDisposable
{
    private const long EmusPerPoint = 12700;

    private readonly string _tempDir;

    public PptxToTypstConverterUnsupportedElementTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), nameof(PptxToTypstConverterUnsupportedElementTests), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        SlideOpsTestHelpers.BestEffortDelete(_tempDir);
    }

    [Fact]
    public void Convert_ChartGraphicFrame_RendersPlaceholderAndWarns()
    {
        var frame = GraphicFrame(10, "Sales chart", x: 100, y: 50, width: 300, height: 200,
            "http://schemas.openxmlformats.org/drawingml/2006/chart");
        var path = CreateDeck(frame);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var slide = Assert.Single(presentation.Slides);

        // Visible placeholder: framed box + centered label at the frame's position.
        var box = Assert.Single(slide.Elements, e => e.Type == "Shape");
        Assert.Equal(100.0, box.X, precision: 3);
        Assert.Equal(50.0, box.Y, precision: 3);
        Assert.Equal(300.0, box.Width, precision: 3);
        Assert.Equal(200.0, box.Height, precision: 3);
        Assert.Equal("rect", box.Shape!.ShapeType);
        Assert.False(string.IsNullOrEmpty(box.Shape.FillColor));
        Assert.True(box.Shape.StrokeWidth > 0);

        var label = Assert.Single(slide.Elements, e => e.Type == "Text");
        Assert.Equal(100.0, label.X, precision: 3);
        Assert.Equal(50.0, label.Y, precision: 3);
        Assert.Contains("Chart", label.Text!.Content);
        Assert.Equal("center", label.Text.VerticalAlign);

        // Warning collected on the slide, identifying the element by name.
        var warning = Assert.Single(slide.Warnings);
        Assert.Contains("Chart", warning);
        Assert.Contains("Sales chart", warning);
        Assert.Contains("placeholder", warning);

        // The placeholder must be present in the emitted Typst source as well.
        var source = converter.GenerateTypstSource(presentation);
        Assert.Contains("#rect(", source);
        Assert.Contains("Chart (not supported)", source);
    }

    [Fact]
    public void Convert_UnknownGraphicFrame_RendersPlaceholderWithGenericLabel()
    {
        var frame = GraphicFrame(11, "Mystery object", x: 10, y: 20, width: 100, height: 60,
            "http://example.com/unsupported");
        var path = CreateDeck(frame);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var slide = Assert.Single(converter.Convert().Slides);

        Assert.Single(slide.Elements, e => e.Type == "Shape");
        var label = Assert.Single(slide.Elements, e => e.Type == "Text");
        Assert.Contains("Unsupported content", label.Text!.Content);

        var warning = Assert.Single(slide.Warnings);
        Assert.Contains("Unsupported content", warning);
        Assert.Contains("Mystery object", warning);
    }

    [Fact]
    public void Convert_UnresolvableDiagramGraphicFrame_RendersPlaceholderAndWarns()
    {
        // Diagram (SmartArt) URI but no diagram parts in the deck: the text approximation
        // cannot yield anything, so a placeholder must be rendered instead.
        var frame = GraphicFrame(12, "Empty SmartArt", x: 0, y: 0, width: 120, height: 80,
            "http://schemas.openxmlformats.org/drawingml/2006/diagram");
        var path = CreateDeck(frame);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var slide = Assert.Single(converter.Convert().Slides);

        Assert.Single(slide.Elements, e => e.Type == "Shape");
        var label = Assert.Single(slide.Elements, e => e.Type == "Text");
        Assert.Contains("SmartArt diagram", label.Text!.Content);

        var warning = Assert.Single(slide.Warnings);
        Assert.Contains("SmartArt", warning);
        Assert.Contains("Empty SmartArt", warning);
    }

    [Fact]
    public void Convert_ReferenceDeckWithSmartArt_KeepsApproximationAndWarns()
    {
        // sales_acceleration_deck.pptx slide 15 contains a SmartArt diagram ("Diagram 16",
        // text: SUSTAIN / DIAGNOSE / DESIGN / DELIVER).  Shape extraction produces
        // pre-rendered shapes (roundRect + rightArrow); text content is preserved.
        // The result must contain shapes + text and warn about theme colour fidelity.
        var referencePath = Path.Combine(ResolveReferenceDirectory(), "sales_acceleration_deck.pptx");
        Assert.True(File.Exists(referencePath), $"Reference deck not found: {referencePath}");

        using var document = PresentationDocument.Open(referencePath, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();

        var smartArtSlides = presentation.Slides
            .Where(s => s.Warnings.Any(w => w.Contains("SmartArt", StringComparison.Ordinal)))
            .ToList();
        var slide = Assert.Single(smartArtSlides);

        var warning = Assert.Single(slide.Warnings);
        Assert.Contains("rendered from pre-rendered shapes", warning);

        // Both shape elements (rects + arrows) and positioned text must be present.
        Assert.Contains(slide.Elements, e => e.Type == "Shape");
        Assert.Contains(slide.Elements, e => e.Type == "Text");
        Assert.DoesNotContain(slide.Elements,
            e => e.Text?.Content.Contains("not supported", StringComparison.Ordinal) == true);
    }

    [Fact]
    public void Convert_ReferenceDeckWithSmartArt_Slide15YieldsShapeElements()
    {
        // Integration test: convert sales_acceleration_deck.pptx and assert that
        // slide 15 yields exactly the diagram's pre-rendered shapes — 4 roundRect
        // boxes + 3 rightArrow connectors (per the drawing part) — plus the four
        // text nodes (SUSTAIN, DIAGNOSE, DESIGN, DELIVER).
        var referencePath = Path.Combine(ResolveReferenceDirectory(), "sales_acceleration_deck.pptx");
        Assert.True(File.Exists(referencePath), $"Reference deck not found: {referencePath}");

        using var document = PresentationDocument.Open(referencePath, false);

        // Frame rect of the diagram graphic frame on slide 15 (for the
        // bounding-box normalisation assertions below).
        var (frameX, frameY, frameW, frameH) = FindDiagramFrameRect(document);

        using var converter = new PptxToTypstConverter(document);
        var presentation = converter.Convert();

        var smartArtSlides = presentation.Slides
            .Where(s => s.Warnings.Any(w => w.Contains("SmartArt", StringComparison.Ordinal)))
            .ToList();
        var slide = Assert.Single(smartArtSlides);

        // Diagram-derived elements carry the dsp:sp modelId; the slide's own
        // shapes (title, text boxes) do not.
        var diagramShapes = slide.Elements
            .Where(e => e.Type == "Shape" && e.ModelId != null)
            .ToList();
        Assert.Equal(7, diagramShapes.Count);

        var rects = diagramShapes.Where(s => s.Shape?.ShapeType == "rect").ToList();
        var polygons = diagramShapes.Where(s => s.Shape?.ShapeType == "polygon").ToList();
        Assert.Equal(4, rects.Count);
        Assert.Equal(3, polygons.Count);

        // Each shape has geometry and the theme accent fill (accent1 = #C00000).
        Assert.All(diagramShapes, s =>
        {
            Assert.True(s.Width > 0);
            Assert.True(s.Height > 0);
            Assert.NotNull(s.Shape);
        });
        Assert.All(rects, s => Assert.Equal("#C00000", s.Shape!.FillColor));
        // roundRect adj val 10000 → 10% of the smaller (normalised) side.
        Assert.All(rects, s => Assert.True(s.Shape!.CornerRadius > 0));

        // Drawing-space → frame normalisation: the shapes' bounding box must span
        // the graphic frame's rect, not sit top-left-anchored at native size.
        var minX = diagramShapes.Min(s => s.X);
        var minY = diagramShapes.Min(s => s.Y);
        var maxX = diagramShapes.Max(s => s.X + s.Width);
        var maxY = diagramShapes.Max(s => s.Y + s.Height);
        Assert.Equal(frameX, minX, 1.0);
        Assert.Equal(frameY, minY, 1.0);
        Assert.Equal(frameX + frameW, maxX, 1.0);
        Assert.Equal(frameY + frameH, maxY, 1.0);

        // Exactly the 4 node labels, with dsp:style fontRef colour (lt1 = white)
        // applied as the paragraph default — not the unresolved black default.
        var diagramTexts = slide.Elements
            .Where(e => e.Type == "Text" && e.ModelId != null)
            .ToList();
        Assert.Equal(4, diagramTexts.Count);
        var labels = diagramTexts.Select(t => t.Text!.Content).ToList();
        Assert.Contains("SUSTAIN", labels);
        Assert.Contains("DIAGNOSE", labels);
        Assert.Contains("DESIGN", labels);
        Assert.Contains("DELIVER", labels);
        Assert.All(diagramTexts, t =>
            Assert.Equal("#FFFFFF", t.Text!.Formatting.Color));
    }

    private static (double X, double Y, double Width, double Height) FindDiagramFrameRect(
        PresentationDocument document)
    {
        foreach (var slidePart in document.PresentationPart!.SlideParts)
        {
            var slide = slidePart.Slide;
            if (slide?.CommonSlideData?.ShapeTree == null) continue;

            var graphicFrame = slide.CommonSlideData.ShapeTree
                .Elements<P.GraphicFrame>()
                .FirstOrDefault(gf =>
                    (gf.Graphic?.GraphicData?.Uri?.Value ?? "")
                        .Contains("/drawingml/2006/diagram", StringComparison.Ordinal));

            if (graphicFrame?.Transform?.Offset == null || graphicFrame.Transform.Extents == null)
                continue;

            const double emusPerPoint = 12700.0;
            return (graphicFrame.Transform.Offset.X!.Value / emusPerPoint,
                    graphicFrame.Transform.Offset.Y!.Value / emusPerPoint,
                    graphicFrame.Transform.Extents.Cx!.Value / emusPerPoint,
                    graphicFrame.Transform.Extents.Cy!.Value / emusPerPoint);
        }

        throw new InvalidOperationException("No diagram graphic frame found in the reference deck.");
    }

    private static P.GraphicFrame GraphicFrame(uint id, string name, double x, double y, double width, double height, string uri)
    {
        return new P.GraphicFrame(
            new NonVisualGraphicFrameProperties(
                new NonVisualDrawingProperties { Id = id, Name = name },
                new NonVisualGraphicFrameDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new P.Transform(
                new Drawing.Offset { X = Pt(x), Y = Pt(y) },
                new Drawing.Extents { Cx = Pt(width), Cy = Pt(height) }),
            new Drawing.Graphic(new Drawing.GraphicData { Uri = uri }));
    }

    private static long Pt(double points) => (long)(points * EmusPerPoint);

    private string CreateDeck(params OpenXmlElement[] slideElements)
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
            slideLayoutPart.SlideLayout = new P.SlideLayout(new CommonSlideData(CreateShapeTree()));
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
            slidePart.Slide = new Slide(new CommonSlideData(CreateShapeTree(slideElements)));
            slidePart.AddPart(slideLayoutPart);

            presentationPart.Presentation.SlideIdList.Append(new SlideId
            {
                Id = 256,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });
        }

        return path;
    }

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

    private static string ResolveReferenceDirectory()
    {
        var baseDir = AppContext.BaseDirectory;
        return Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "..", "examples", "REF", "PPTX"));
    }
}
