using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using OfficeEditor.Core.Services;
using PptxEditor.Core.Converters;
using Xunit;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Regression tests for rounded-corner pictures in the PPTX → Typst converter.
/// The clipping wrapper must be <c>#block(clip: true, radius: …)</c>: Typst's
/// <c>#rect</c> has no <c>clip</c> argument, so the previous
/// <c>#rect(clip: true, …)</c> emission failed compilation with
/// "unexpected argument: clip".
/// </summary>
public sealed class PptxToTypstConverterRoundedImageTests : IDisposable
{
    private const string EnableCompileEnvVar = "OE_RUN_TYPST_COMPILE_TESTS";

    private readonly string _tempDir;

    public PptxToTypstConverterRoundedImageTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), nameof(PptxToTypstConverterRoundedImageTests), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        SlideOpsTestHelpers.BestEffortDelete(_tempDir);
    }

    [Fact]
    public void Convert_RoundedCornerPicture_EmitsClippingBlockNotRect()
    {
        var path = CreateDeckWithRoundedImage();
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var element = Assert.Single(presentation.Slides[0].Elements, e => e.Type == "Image");
        Assert.NotNull(element.Image);
        Assert.True(element.Image!.CornerRadius > 0,
            "rounded-rect geometry must produce a corner radius for the image");

        var source = converter.GenerateTypstSource(presentation);

        Assert.DoesNotContain("#rect(clip", source);
        Assert.Matches(
            @"#block\(clip: true, width: [0-9.]+pt, height: [0-9.]+pt, radius: [0-9.]+pt\)\[#image\(",
            source);
    }

    [Fact]
    public void Convert_RoundedCornerPicture_SourceCompiles()
    {
        // Opt-in: requires a Typst backend (TypstBridge → CLI chain).
        if (Environment.GetEnvironmentVariable(EnableCompileEnvVar) != "1")
        {
            return;
        }

        var path = CreateDeckWithRoundedImage();
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var source = converter.GenerateTypstSource(presentation);

        using var compiler = new TypstCompilerService();
        var result = compiler.Compile(source, new CompileOptions
        {
            Format = OutputFormat.Svg,
            WorkingDirectory = presentation.TempDirectory
        });

        Assert.True(result.Success, $"rounded-image source failed to compile: {result.ErrorMessage}");
        Assert.NotEmpty(result.Pages);
    }

    [Fact]
    public void Convert_PlainRectanglePicture_DoesNotReceiveDefaultCornerRadius()
    {
        var path = CreateDeckWithImage(Drawing.ShapeTypeValues.Rectangle);
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var element = Assert.Single(converter.Convert().Slides[0].Elements, e => e.Type == "Image");
        Assert.Equal(0.0, element.Image!.CornerRadius);
    }

    private string CreateDeckWithRoundedImage()
        => CreateDeckWithImage(Drawing.ShapeTypeValues.Round2SameRectangle);

    private string CreateDeckWithImage(Drawing.ShapeTypeValues preset)
    {
        var deckPath = Path.Combine(_tempDir, $"{Guid.NewGuid():N}.pptx");
        var pngPath = SlideOpsTestHelpers.WriteMinimalPng(_tempDir);

        using (var document = PresentationDocument.Create(deckPath, PresentationDocumentType.Presentation))
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
            var imagePart = slidePart.AddImagePart(ImagePartType.Png);
            using (var imageStream = File.OpenRead(pngPath))
            {
                imagePart.FeedData(imageStream);
            }

            var embedId = slidePart.GetIdOfPart(imagePart);
            slidePart.Slide = new Slide(new CommonSlideData(CreateShapeTree(RoundedImageShape(embedId, preset))));
            slidePart.AddPart(slideLayoutPart);

            presentationPart.Presentation.SlideIdList.Append(new SlideId
            {
                Id = 256,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });
        }

        return deckPath;
    }

    /// <summary>
    /// A blip-filled rounded-rect shape: the converter turns shape-with-picture-fill
    /// into an Image element whose CornerRadius comes from the preset geometry.
    /// </summary>
    private static P.Shape RoundedImageShape(string embedId, Drawing.ShapeTypeValues preset)
    {
        return new P.Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = 2, Name = "Rounded Picture" },
                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(
                new Drawing.Transform2D(
                    new Drawing.Offset { X = 1000000, Y = 500000 },
                    new Drawing.Extents { Cx = 2540000, Cy = 1270000 }),
                new Drawing.PresetGeometry(new Drawing.AdjustValueList())
                {
                    Preset = preset
                },
                new Drawing.BlipFill(
                    new Drawing.Blip { Embed = embedId },
                    new Drawing.Stretch(new Drawing.FillRectangle()))));
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
}
