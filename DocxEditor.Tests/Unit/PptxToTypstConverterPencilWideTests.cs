using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PptxEditor.Core.Converters;
using Xunit;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Regression tests for defects found on the  decks :
/// layout picture rId collisions, picture-placeholder position inheritance,
/// and group rotation in deeply nested group hierarchies. All decks are
/// synthetic — no licensed content is committed.
/// </summary>
public sealed class PptxToTypstConverterWideTests : IDisposable
{
    // 2x2 red and 2x2 blue PNGs (ImageMagick-generated), used to tell two
    // image parts apart when relationship ids collide across parts.
    private const string RedPngBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACAQMAAABIeJ9nAAAAIGNIUk0AAHomAACAhAAA+gAAAIDoAAB1MAAA6mAAADqYAAAXcJy6UTwAAAAGUExURf8AAP///0EdNBEAAAABYktHRAH/Ai3eAAAAB3RJTUUH6gcaDQMvfqsEZgAAACV0RVh0ZGF0ZTpjcmVhdGUAMjAyNi0wNy0yNlQxMzowMzo0NyswMDowMDnhSNUAAAAldEVYdGRhdGU6bW9kaWZ5ADIwMjYtMDctMjZUMTM6MDM6NDcrMDA6MDBIvPBpAAAAKHRFWHRkYXRlOnRpbWVzdGFtcAAyMDI2LTA3LTI2VDEzOjAzOjQ3KzAwOjAwH6nRtgAAAAxJREFUCNdjYGBgAAAABAABJzQnCgAAAABJRU5ErkJggg==";
    private const string BluePngBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACAQMAAABIeJ9nAAAAIGNIUk0AAHomAACAhAAA+gAAAIDoAAB1MAAA6mAAADqYAAAXcJy6UTwAAAAGUExURQAA/////3vcmSwAAAABYktHRAH/Ai3eAAAAB3RJTUUH6gcaDQMvfqsEZgAAACV0RVh0ZGF0ZTpjcmVhdGUAMjAyNi0wNy0yNlQxMzowMzo0NyswMDowMDnhSNUAAAAldEVYdGRhdGU6bW9kaWZ5ADIwMjYtMDctMjZUMTM6MDM6NDcrMDA6MDBIvPBpAAAAKHRFWHRkYXRlOnRpbWVzdGFtcAAyMDI2LTA3LTI2VDEzOjAzOjQ3KzAwOjAwH6nRtgAAAAxJREFUCNdjYGBgAAAABAABJzQnCgAAAABJRU5ErkJggg==";

    private readonly string _tempDir;

    public PptxToTypstConverterWideTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), nameof(PptxToTypstConverterWideTests), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        SlideOpsTestHelpers.BestEffortDelete(_tempDir);
    }

    /// <summary>
    /// A layout picture's <c>r:embed</c> is scoped to the LAYOUT part's
    /// relationships. When the same rId exists on the slide part (pointing at
    /// a different image), the layout picture must still resolve through the
    /// layout part — otherwise the layout logo renders the slide's image.
    /// </summary>
    [Fact]
    public void Convert_LayoutPictureWithCollidingRelId_ResolvesViaLayoutPart()
    {
        var deckPath = Path.Combine(_tempDir, $"{Guid.NewGuid():N}.pptx");
        var redBytes = Convert.FromBase64String(RedPngBase64);
        var blueBytes = Convert.FromBase64String(BluePngBase64);

        using (var document = PresentationDocument.Create(deckPath, PresentationDocumentType.Presentation))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new Presentation
            {
                SlideMasterIdList = new SlideMasterIdList(),
                SlideIdList = new SlideIdList(),
                SlideSize = new SlideSize { Cx = 12192000, Cy = 6858000 }
            };

            var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
            slideMasterPart.SlideMaster = new SlideMaster(
                new CommonSlideData(CreateShapeTree()),
                CreateColorMap(),
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

            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.AddPart(slideLayoutPart);
            presentationPart.Presentation.SlideIdList.Append(new SlideId
            {
                Id = 256,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });

            // Colliding rIds: slide rId7 → red, layout rId7 → blue.
            var slideImagePart = slidePart.AddImagePart(ImagePartType.Png, "rId7");
            using (var stream = new MemoryStream(redBytes)) slideImagePart.FeedData(stream);
            var layoutImagePart = slideLayoutPart.AddImagePart(ImagePartType.Png, "rId7");
            using (var stream = new MemoryStream(blueBytes)) layoutImagePart.FeedData(stream);

            slideLayoutPart.SlideLayout = new P.SlideLayout(new CommonSlideData(
                CreateShapeTree(LayoutPicture("rId7"))));
            slidePart.Slide = new Slide(new CommonSlideData(CreateShapeTree()));
        }

        using (var document = PresentationDocument.Open(deckPath, false))
        using (var converter = new PptxToTypstConverter(document))
        {
            var presentation = converter.Convert();
            var image = Assert.Single(presentation.Slides[0].Elements, e => e.Type == "Image");
            var emittedBytes = File.ReadAllBytes(image.Image!.FullPath);
            Assert.Equal(blueBytes, emittedBytes);
        }
    }

    /// <summary>
    /// A slide <c>p:pic</c> placeholder with an empty <c>p:spPr</c> inherits
    /// its position/size from the layout placeholder shape with the matching
    /// <c>ph type="pic"</c> index (ECMA-376 placeholder inheritance), instead
    /// of falling back to (0, 0, 100x50pt).
    /// </summary>
    [Fact]
    public void Convert_PicturePlaceholderWithoutTransform_InheritsLayoutPlaceholderXfrm()
    {
        var deckPath = Path.Combine(_tempDir, $"{Guid.NewGuid():N}.pptx");
        var redBytes = Convert.FromBase64String(RedPngBase64);

        using (var document = PresentationDocument.Create(deckPath, PresentationDocumentType.Presentation))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new Presentation
            {
                SlideMasterIdList = new SlideMasterIdList(),
                SlideIdList = new SlideIdList(),
                SlideSize = new SlideSize { Cx = 12192000, Cy = 6858000 }
            };

            var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
            slideMasterPart.SlideMaster = new SlideMaster(
                new CommonSlideData(CreateShapeTree()),
                CreateColorMap(),
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

            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.AddPart(slideLayoutPart);
            presentationPart.Presentation.SlideIdList.Append(new SlideId
            {
                Id = 256,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });

            var imagePart = slidePart.AddImagePart(ImagePartType.Png);
            using (var stream = new MemoryStream(redBytes)) imagePart.FeedData(stream);
            var embedId = slidePart.GetIdOfPart(imagePart);

            // Layout: pic placeholder sp at a known xfrm.
            slideLayoutPart.SlideLayout = new P.SlideLayout(new CommonSlideData(
                CreateShapeTree(LayoutPicPlaceholderShape())));
            // Slide: pic placeholder with NO transform.
            slidePart.Slide = new Slide(new CommonSlideData(
                CreateShapeTree(SlidePlaceholderPicture(embedId))));
        }

        using (var document = PresentationDocument.Open(deckPath, false))
        using (var converter = new PptxToTypstConverter(document))
        {
            var presentation = converter.Convert();
            var image = Assert.Single(presentation.Slides[0].Elements, e => e.Type == "Image");

            // Layout placeholder: off (1000000, 2000000), ext (3000000, 1500000) EMU → pt.
            AssertInRange(image.X, 1000000 / 12700.0);
            AssertInRange(image.Y, 2000000 / 12700.0);
            AssertInRange(image.Width, 3000000 / 12700.0);
            AssertInRange(image.Height, 1500000 / 12700.0);
        }
    }

    /// <summary>
    /// Group <c>a:xfrm rot</c> must rotate every child about the group's
    /// centre and add the angle to the child's own rotation. PowerPoint
    /// templates lean on this heavily (e.g. horizontal  glyphs stood
    /// upright via rot="16200000").
    /// </summary>
    [Fact]
    public void Convert_RotatedGroup_RotatesChildrenAboutGroupCentre()
    {
        var deckPath = Path.Combine(_tempDir, $"{Guid.NewGuid():N}.pptx");

        using (var document = PresentationDocument.Create(deckPath, PresentationDocumentType.Presentation))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new Presentation
            {
                SlideMasterIdList = new SlideMasterIdList(),
                SlideIdList = new SlideIdList(),
                SlideSize = new SlideSize { Cx = 12192000, Cy = 6858000 }
            };

            var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
            slideMasterPart.SlideMaster = new SlideMaster(
                new CommonSlideData(CreateShapeTree()),
                CreateColorMap(),
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
            slidePart.AddPart(slideLayoutPart);
            presentationPart.Presentation.SlideIdList.Append(new SlideId
            {
                Id = 256,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });

            // Group: rot 270°, off (1000000, 1000000), ext (2000000, 1000000),
            // child space identical (chOff 0,0 / chExt 2000000x1000000).
            var group = new P.GroupShape(
                new P.NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = 10, Name = "Rotated Group" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(
                    new Drawing.TransformGroup(
                        new Drawing.Offset { X = 1000000, Y = 1000000 },
                        new Drawing.Extents { Cx = 2000000, Cy = 1000000 },
                        new Drawing.ChildOffset { X = 0, Y = 0 },
                        new Drawing.ChildExtents { Cx = 2000000, Cy = 1000000 })
                    { Rotation = 16200000 }),
                // Child: off (0,0), ext (1000000, 500000), solid red fill.
                new P.Shape(
                    new P.NonVisualShapeProperties(
                        new NonVisualDrawingProperties { Id = 11, Name = "Child" },
                        new P.NonVisualShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new ShapeProperties(
                        new Drawing.Transform2D(
                            new Drawing.Offset { X = 0, Y = 0 },
                            new Drawing.Extents { Cx = 1000000, Cy = 500000 }),
                        new Drawing.PresetGeometry(new Drawing.AdjustValueList())
                        { Preset = Drawing.ShapeTypeValues.Rectangle },
                        new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = "FF0000" })),
                    new P.TextBody(
                        new Drawing.BodyProperties(),
                        new Drawing.ListStyle(),
                        new Drawing.Paragraph())));

            slidePart.Slide = new Slide(new CommonSlideData(CreateShapeTree(group)));
        }

        using (var document = PresentationDocument.Open(deckPath, false))
        using (var converter = new PptxToTypstConverter(document))
        {
            var presentation = converter.Convert();
            var shape = Assert.Single(presentation.Slides[0].Elements, e => e.Type == "Shape");

            // Group centre (2000000, 1500000); child centre (500000, 250000).
            // Rotating 270° clockwise about the group centre maps the child
            // centre to (750000, 3000000) EMU.
            AssertInRange(shape.Rotation, 270.0);
            AssertInRange(shape.X + shape.Width / 2, 750000 / 12700.0);
            AssertInRange(shape.Y + shape.Height / 2, 3000000 / 12700.0);
        }
    }

    private static void AssertInRange(double actual, double expected, double tolerance = 0.05)
    {
        Assert.True(Math.Abs(actual - expected) <= tolerance,
            $"expected {expected:F4} ± {tolerance}, got {actual:F4}");
    }

    private static ColorMap CreateColorMap() => new()
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
    };

    private static P.Picture LayoutPicture(string embedId)
    {
        return new P.Picture(
            new P.NonVisualPictureProperties(
                new NonVisualDrawingProperties { Id = 20, Name = "Layout Logo" },
                new P.NonVisualPictureDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new P.BlipFill(
                new Drawing.Blip { Embed = embedId },
                new Drawing.Stretch(new Drawing.FillRectangle())),
            new ShapeProperties(
                new Drawing.Transform2D(
                    new Drawing.Offset { X = 500000, Y = 6000000 },
                    new Drawing.Extents { Cx = 1500000, Cy = 400000 })));
    }

    private static P.Shape LayoutPicPlaceholderShape()
    {
        return new P.Shape(
            new P.NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = 21, Name = "Picture Placeholder" },
                new P.NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties(
                    new PlaceholderShape { Type = PlaceholderValues.Picture, Index = 10 })),
            new ShapeProperties(
                new Drawing.Transform2D(
                    new Drawing.Offset { X = 1000000, Y = 2000000 },
                    new Drawing.Extents { Cx = 3000000, Cy = 1500000 }),
                new Drawing.PresetGeometry(new Drawing.AdjustValueList())
                { Preset = Drawing.ShapeTypeValues.Rectangle }),
            new P.TextBody(
                new Drawing.BodyProperties(),
                new Drawing.ListStyle(),
                new Drawing.Paragraph()));
    }

    private static P.Picture SlidePlaceholderPicture(string embedId)
    {
        return new P.Picture(
            new P.NonVisualPictureProperties(
                new NonVisualDrawingProperties { Id = 22, Name = "Slide Picture Placeholder" },
                new P.NonVisualPictureDrawingProperties(
                    new Drawing.PictureLocks { NoChangeAspect = true }),
                new ApplicationNonVisualDrawingProperties(
                    new PlaceholderShape { Type = PlaceholderValues.Picture, Index = 10 })),
            new P.BlipFill(
                new Drawing.Blip { Embed = embedId },
                new Drawing.Stretch(new Drawing.FillRectangle())),
            new ShapeProperties());
    }

    private static ShapeTree CreateShapeTree(params OpenXmlElement[] elements)
    {
        var shapeTree = new ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new NonVisualDrawingProperties { Id = 0, Name = "" },
                new P.NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties(
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
