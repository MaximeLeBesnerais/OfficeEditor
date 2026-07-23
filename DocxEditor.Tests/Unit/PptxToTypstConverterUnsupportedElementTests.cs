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
        // pres-pro.pptx slide 15 contains a SmartArt diagram that the converter
        // approximates as positioned text. The approximation must be kept (no
        // placeholder) and reported via Warnings.
        var referencePath = Path.Combine(ResolveReferenceDirectory(), "pres-pro.pptx");
        if (!File.Exists(referencePath))
        {
            return; // REF fixture purged pre-public-release (licensing); replacement pending
        }

        using var document = PresentationDocument.Open(referencePath, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();

        var smartArtSlides = presentation.Slides
            .Where(s => s.Warnings.Any(w => w.Contains("SmartArt", StringComparison.Ordinal)))
            .ToList();
        var slide = Assert.Single(smartArtSlides);

        var warning = Assert.Single(slide.Warnings);
        Assert.Contains("approximated", warning);

        // Approximation stays: positioned text elements, no SmartArt placeholder label.
        Assert.Contains(slide.Elements, e => e.Type == "Text");
        Assert.DoesNotContain(slide.Elements,
            e => e.Text?.Content.Contains("not supported", StringComparison.Ordinal) == true);
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
