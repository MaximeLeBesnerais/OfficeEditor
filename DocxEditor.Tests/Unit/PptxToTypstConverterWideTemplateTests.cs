using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PptxEditor.Core.Converters;
using Xunit;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Regression tests for defects found on wide-aspect template decks:
/// layout picture rId collisions, picture-placeholder position inheritance,
/// and group rotation in deeply nested group hierarchies. All decks are
/// synthetic — no licensed content is committed.
/// </summary>
public sealed class PptxToTypstConverterWideTemplateTests : IDisposable
{
    // 2x2 red and 2x2 blue PNGs (ImageMagick-generated), used to tell two
    // image parts apart when relationship ids collide across parts.
    private const string RedPngBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACAQMAAABIeJ9nAAAAIGNIUk0AAHomAACAhAAA+gAAAIDoAAB1MAAA6mAAADqYAAAXcJy6UTwAAAAGUExURf8AAP///0EdNBEAAAABYktHRAH/Ai3eAAAAB3RJTUUH6gcaDQMvfqsEZgAAACV0RVh0ZGF0ZTpjcmVhdGUAMjAyNi0wNy0yNlQxMzowMzo0NyswMDowMDnhSNUAAAAldEVYdGRhdGU6bW9kaWZ5ADIwMjYtMDctMjZUMTM6MDM6NDcrMDA6MDBIvPBpAAAAKHRFWHRkYXRlOnRpbWVzdGFtcAAyMDI2LTA3LTI2VDEzOjAzOjQ3KzAwOjAwH6nRtgAAAAxJREFUCNdjYGBgAAAABAABJzQnCgAAAABJRU5ErkJggg==";
    private const string BluePngBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACAQMAAABIeJ9nAAAAIGNIUk0AAHomAACAhAAA+gAAAIDoAAB1MAAA6mAAADqYAAAXcJy6UTwAAAAGUExURQAA/////3vcmSwAAAABYktHRAH/Ai3eAAAAB3RJTUUH6gcaDQMvfqsEZgAAACV0RVh0ZGF0ZTpjcmVhdGUAMjAyNi0wNy0yNlQxMzowMzo0NyswMDowMDnhSNUAAAAldEVYdGRhdGU6bW9kaWZ5ADIwMjYtMDctMjZUMTM6MDM6NDcrMDA6MDBIvPBpAAAAKHRFWHRkYXRlOnRpbWVzdGFtcAAyMDI2LTA3LTI2VDEzOjAzOjQ3KzAwOjAwH6nRtgAAAAxJREFUCNdjYGBgAAAABAABJzQnCgAAAABJRU5ErkJggg==";

    private readonly string _tempDir;

    public PptxToTypstConverterWideTemplateTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), nameof(PptxToTypstConverterWideTemplateTests), Guid.NewGuid().ToString("N"));
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
    /// templates lean on this heavily (e.g. horizontal glyphs stood upright
    /// via rot="16200000").
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

            // Group centre (2000000, 1500000); child centre mapped into parent
            // space = grpOff + (500000, 250000) = (1500000, 1250000).
            // Rotating 270° clockwise about the group centre maps the child
            // centre to (1750000, 2000000) EMU.
            AssertInRange(shape.Rotation, 270.0);
            AssertInRange(shape.X + shape.Width / 2, 1750000 / 12700.0);
            AssertInRange(shape.Y + shape.Height / 2, 2000000 / 12700.0);
        }
    }

    /// <summary>
    /// algn="just" (justified) must emit Typst par(justify: true), not a
    /// left-aligned fallback and never an invalid #align(justify) — the
    /// template body paragraphs are justified, visible as stretched word
    /// spacing in the official PDFs.
    /// </summary>
    [Fact]
    public void Convert_JustifiedParagraph_EmitsParJustify()
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

            var shape = new P.Shape(
                new P.NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = 80, Name = "Body" },
                    new P.NonVisualShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new ShapeProperties(
                    new Drawing.Transform2D(
                        new Drawing.Offset { X = 1000000, Y = 1000000 },
                        new Drawing.Extents { Cx = 4000000, Cy = 1000000 }),
                    new Drawing.PresetGeometry(new Drawing.AdjustValueList())
                    { Preset = Drawing.ShapeTypeValues.Rectangle }),
                new P.TextBody(
                    new Drawing.BodyProperties(),
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph(
                        new Drawing.ParagraphProperties { Alignment = Drawing.TextAlignmentTypeValues.Justified },
                        new Drawing.Run(
                            new Drawing.RunProperties { Language = "en-US", FontSize = 1400 },
                            new Drawing.Text("justified body text that should stretch across the full box width")))));

            slidePart.Slide = new Slide(new CommonSlideData(CreateShapeTree(shape)));
        }

        using (var document = PresentationDocument.Open(deckPath, false))
        using (var converter = new PptxToTypstConverter(document))
        {
            var presentation = converter.Convert();
            var source = converter.GenerateTypstSource(presentation);

            Assert.Contains("#set par(justify: true)", source);
            Assert.DoesNotContain("#align(justify)", source);
        }
    }

    /// <summary>
    /// bodyPr wrap="none" (typically with spAutoFit): PowerPoint never wraps
    /// the line — the box grows/overflows instead. The converter must widen
    /// the emitted block beyond the shape width so fallback-font metrics
    /// cannot force a wrap (single-line cover taglines are one line in the
    /// official PDF).
    /// </summary>
    [Fact]
    public void Convert_NoWrapTextBox_EmitsBlockWiderThanShape()
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

            // 100pt-wide box (1270000 EMU) with wrap="none" and a line that
            // measures far wider than 100pt.
            var shape = new P.Shape(
                new P.NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = 70, Name = "Tagline" },
                    new P.NonVisualShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new ShapeProperties(
                    new Drawing.Transform2D(
                        new Drawing.Offset { X = 2000000, Y = 1000000 },
                        new Drawing.Extents { Cx = 1270000, Cy = 500000 }),
                    new Drawing.PresetGeometry(new Drawing.AdjustValueList())
                    { Preset = Drawing.ShapeTypeValues.Rectangle }),
                new P.TextBody(
                    new Drawing.BodyProperties { Wrap = Drawing.TextWrappingValues.None },
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph(
                        new Drawing.Run(
                            new Drawing.RunProperties { Language = "en-US", FontSize = 2800, Bold = true },
                            new Drawing.Text("Premium Presentation Collection")))));

            slidePart.Slide = new Slide(new CommonSlideData(CreateShapeTree(shape)));
        }

        using (var document = PresentationDocument.Open(deckPath, false))
        using (var converter = new PptxToTypstConverter(document))
        {
            var presentation = converter.Convert();
            var text = Assert.Single(presentation.Slides[0].Elements, e => e.Type == "Text");
            Assert.True(text.Text!.NoWrap, "wrap=\"none\" must set the NoWrap flag");

            var source = converter.GenerateTypstSource(presentation);
            var match = System.Text.RegularExpressions.Regex.Match(
                source, @"#place\(top \+ left, dx: [0-9.]+pt, dy: [0-9.]+pt\)\[#block\(width: ([0-9.]+)pt");
            Assert.True(match.Success, "expected a placed width-constrained text block");
            var emittedWidth = double.Parse(match.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
            Assert.True(emittedWidth > 100.0,
                $"no-wrap block must exceed the 100pt shape width, got {emittedWidth:F2}pt");
        }
    }

    /// <summary>
    /// A run with no color anywhere in the cascade inherits the theme's tx1
    /// (→ dk1) color — not hardcoded black. Some templates redefine dk1 as
    /// grey #95A5A6, and their body text renders grey in the official PDFs.
    /// </summary>
    [Fact]
    public void Convert_ColorlessRun_InheritsThemeTx1Color()
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

            var themePart = slideMasterPart.AddNewPart<ThemePart>();
            themePart.Theme = new Drawing.Theme(
                new Drawing.ThemeElements(
                    new Drawing.ColorScheme(
                        new Drawing.Dark1Color(new Drawing.RgbColorModelHex { Val = "95A5A6" }),
                        new Drawing.Light1Color(new Drawing.SystemColor { Val = Drawing.SystemColorValues.Window, LastColor = "FFFFFF" }),
                        new Drawing.Dark2Color(new Drawing.RgbColorModelHex { Val = "1F497D" }),
                        new Drawing.Light2Color(new Drawing.RgbColorModelHex { Val = "EEECE1" }),
                        new Drawing.Accent1Color(new Drawing.RgbColorModelHex { Val = "4F81BD" }),
                        new Drawing.Accent2Color(new Drawing.RgbColorModelHex { Val = "C0504D" }),
                        new Drawing.Accent3Color(new Drawing.RgbColorModelHex { Val = "9BBB59" }),
                        new Drawing.Accent4Color(new Drawing.RgbColorModelHex { Val = "8064A2" }),
                        new Drawing.Accent5Color(new Drawing.RgbColorModelHex { Val = "4BACC6" }),
                        new Drawing.Accent6Color(new Drawing.RgbColorModelHex { Val = "F79646" }),
                        new Drawing.Hyperlink(new Drawing.RgbColorModelHex { Val = "16A085" }),
                        new Drawing.FollowedHyperlinkColor(new Drawing.RgbColorModelHex { Val = "800080" })) { Name = "Office" },
                    new Drawing.FontScheme(new Drawing.MajorFont(), new Drawing.MinorFont()) { Name = "Office" },
                    new Drawing.FormatScheme(new Drawing.FillStyleList(), new Drawing.LineStyleList(), new Drawing.EffectStyleList(), new Drawing.BackgroundFillStyleList()) { Name = "Office" }),
                new Drawing.ObjectDefaults(),
                new Drawing.ExtraColorSchemeList()) { Name = "Office Theme" };

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

            var shape = new P.Shape(
                new P.NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = 60, Name = "Body" },
                    new P.NonVisualShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new ShapeProperties(
                    new Drawing.Transform2D(
                        new Drawing.Offset { X = 1000000, Y = 1000000 },
                        new Drawing.Extents { Cx = 3000000, Cy = 500000 }),
                    new Drawing.PresetGeometry(new Drawing.AdjustValueList())
                    { Preset = Drawing.ShapeTypeValues.Rectangle }),
                new P.TextBody(
                    new Drawing.BodyProperties(),
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph(
                        new Drawing.Run(
                            new Drawing.RunProperties { Language = "en-US", FontSize = 1400 },
                            new Drawing.Text("colorless body text")))));

            slidePart.Slide = new Slide(new CommonSlideData(CreateShapeTree(shape)));
        }

        using (var document = PresentationDocument.Open(deckPath, false))
        using (var converter = new PptxToTypstConverter(document))
        {
            var presentation = converter.Convert();
            var text = Assert.Single(presentation.Slides[0].Elements, e => e.Type == "Text");
            var run = Assert.Single(text.Text!.Paragraphs.SelectMany(p => p.Runs));

            Assert.Equal("#95A5A6", run.Formatting.Color);
        }
    }

    /// <summary>
    /// Hyperlink runs render in the theme hlink color with an underline —
    /// the behaviour of the reference renderer (and LibreOffice), which
    /// overrides even an explicit solidFill. A hyperlink that is white/50%
    /// in XML renders teal #16A085 + underline in the official PDF.
    /// </summary>
    [Fact]
    public void Convert_HyperlinkRun_UsesThemeHlinkColorAndUnderline()
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

            var themePart = slideMasterPart.AddNewPart<ThemePart>();
            themePart.Theme = new Drawing.Theme(
                new Drawing.ThemeElements(
                    new Drawing.ColorScheme(
                        new Drawing.Dark1Color(new Drawing.SystemColor { Val = Drawing.SystemColorValues.WindowText, LastColor = "000000" }),
                        new Drawing.Light1Color(new Drawing.SystemColor { Val = Drawing.SystemColorValues.Window, LastColor = "FFFFFF" }),
                        new Drawing.Dark2Color(new Drawing.RgbColorModelHex { Val = "1F497D" }),
                        new Drawing.Light2Color(new Drawing.RgbColorModelHex { Val = "EEECE1" }),
                        new Drawing.Accent1Color(new Drawing.RgbColorModelHex { Val = "4F81BD" }),
                        new Drawing.Accent2Color(new Drawing.RgbColorModelHex { Val = "C0504D" }),
                        new Drawing.Accent3Color(new Drawing.RgbColorModelHex { Val = "9BBB59" }),
                        new Drawing.Accent4Color(new Drawing.RgbColorModelHex { Val = "8064A2" }),
                        new Drawing.Accent5Color(new Drawing.RgbColorModelHex { Val = "4BACC6" }),
                        new Drawing.Accent6Color(new Drawing.RgbColorModelHex { Val = "F79646" }),
                        new Drawing.Hyperlink(new Drawing.RgbColorModelHex { Val = "16A085" }),
                        new Drawing.FollowedHyperlinkColor(new Drawing.RgbColorModelHex { Val = "800080" })) { Name = "Office" },
                    new Drawing.FontScheme(new Drawing.MajorFont(), new Drawing.MinorFont()) { Name = "Office" },
                    new Drawing.FormatScheme(new Drawing.FillStyleList(), new Drawing.LineStyleList(), new Drawing.EffectStyleList(), new Drawing.BackgroundFillStyleList()) { Name = "Office" }),
                new Drawing.ObjectDefaults(),
                new Drawing.ExtraColorSchemeList()) { Name = "Office Theme" };

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

            // Run with an explicit white fill AND a hyperlink: the hlink
            // styling wins (reference-renderer behaviour).
            var shape = new P.Shape(
                new P.NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = 50, Name = "Link" },
                    new P.NonVisualShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new ShapeProperties(
                    new Drawing.Transform2D(
                        new Drawing.Offset { X = 1000000, Y = 1000000 },
                        new Drawing.Extents { Cx = 3000000, Cy = 500000 }),
                    new Drawing.PresetGeometry(new Drawing.AdjustValueList())
                    { Preset = Drawing.ShapeTypeValues.Rectangle }),
                new P.TextBody(
                    new Drawing.BodyProperties(),
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph(
                        new Drawing.Run(
                            new Drawing.RunProperties(
                                new Drawing.SolidFill(new Drawing.PresetColor { Val = Drawing.PresetColorValues.White }),
                                new Drawing.HyperlinkOnClick { Id = "rId99" })
                            { Language = "en-US", FontSize = 1200 },
                            new Drawing.Text("www.example.com")))));

            slidePart.Slide = new Slide(new CommonSlideData(CreateShapeTree(shape)));
        }

        using (var document = PresentationDocument.Open(deckPath, false))
        using (var converter = new PptxToTypstConverter(document))
        {
            var presentation = converter.Convert();
            var text = Assert.Single(presentation.Slides[0].Elements, e => e.Type == "Text");
            var run = Assert.Single(text.Text!.Paragraphs.SelectMany(p => p.Runs));

            Assert.Equal("#16A085", run.Formatting.Color);
            Assert.True(run.Formatting.Underline);
        }
    }

    /// <summary>
    /// A picture-filled shape with blipFill rotWithShape="0" (the OOXML
    /// default) keeps its fill slide-aligned: PowerPoint does not rotate the
    /// image with the shape. For quarter-turn rotations the displayed bounds
    /// are the swapped box (a cover's portrait screenshot is a portrait rect
    /// rotated 270° whose image must stay upright/landscape).
    /// </summary>
    [Fact]
    public void Convert_BlipFillNotRotatingWithShape_KeepsImageSlideAligned()
    {
        var deckPath = Path.Combine(_tempDir, $"{Guid.NewGuid():N}.pptx");
        CreateDeckWithRotatedBlipFill(deckPath, rotateWithShape: false);

        using var document = PresentationDocument.Open(deckPath, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var image = Assert.Single(presentation.Slides[0].Elements, e => e.Type == "Image");

        // Shape: off (1000000, 2000000), ext (1000000 x 2000000) portrait,
        // rot 270° → image fills the landscape box, unrotated.
        AssertInRange(image.Rotation, 0.0);
        AssertInRange(image.Width, 2000000 / 12700.0);
        AssertInRange(image.Height, 1000000 / 12700.0);
        // Same centre as the shape: (1500000, 3000000) EMU.
        AssertInRange(image.X + image.Width / 2, 1500000 / 12700.0);
        AssertInRange(image.Y + image.Height / 2, 3000000 / 12700.0);
    }

    [Fact]
    public void Convert_BlipFillRotatingWithShape_KeepsShapeRotation()
    {
        var deckPath = Path.Combine(_tempDir, $"{Guid.NewGuid():N}.pptx");
        CreateDeckWithRotatedBlipFill(deckPath, rotateWithShape: true);

        using var document = PresentationDocument.Open(deckPath, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var image = Assert.Single(presentation.Slides[0].Elements, e => e.Type == "Image");

        AssertInRange(image.Rotation, 270.0);
        AssertInRange(image.Width, 1000000 / 12700.0);
        AssertInRange(image.Height, 2000000 / 12700.0);
    }

    private void CreateDeckWithRotatedBlipFill(string deckPath, bool rotateWithShape)
    {
        var redBytes = Convert.FromBase64String(RedPngBase64);

        using var document = PresentationDocument.Create(deckPath, PresentationDocumentType.Presentation);
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

        var imagePart = slidePart.AddImagePart(ImagePartType.Png);
        using (var stream = new MemoryStream(redBytes)) imagePart.FeedData(stream);
        var embedId = slidePart.GetIdOfPart(imagePart);

        var blipFill = new Drawing.BlipFill(
            new Drawing.Blip { Embed = embedId },
            new Drawing.Stretch(new Drawing.FillRectangle()))
        {
            RotateWithShape = rotateWithShape
        };

        var shape = new P.Shape(
            new P.NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = 40, Name = "Rotated Picture Fill" },
                new P.NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(
                new Drawing.Transform2D(
                    new Drawing.Offset { X = 1000000, Y = 2000000 },
                    new Drawing.Extents { Cx = 1000000, Cy = 2000000 })
                { Rotation = 16200000 },
                new Drawing.PresetGeometry(new Drawing.AdjustValueList())
                { Preset = Drawing.ShapeTypeValues.Rectangle },
                blipFill),
            new P.TextBody(
                new Drawing.BodyProperties(),
                new Drawing.ListStyle(),
                new Drawing.Paragraph()));

        slidePart.Slide = new Slide(new CommonSlideData(CreateShapeTree(shape)));
    }

    /// <summary>
    /// User-drawn (non-placeholder) shapes on the slide MASTER — logos,
    /// taglines, watermark art — are part of every slide using that master
    /// (unless the layout sets showMasterSp="0"). The converter imported only
    /// layout shapes, so master content silently vanished (a cover lost its
    /// "Template gallery preview" tagline and logo).
    /// </summary>
    [Fact]
    public void Convert_MasterUserDrawnShape_RendersOnSlide()
    {
        var deckPath = Path.Combine(_tempDir, $"{Guid.NewGuid():N}.pptx");
        CreateDeckWithMasterContent(deckPath, showMasterShapes: true, out var blueBytes);

        using var document = PresentationDocument.Open(deckPath, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var elements = presentation.Slides[0].Elements;

        var text = Assert.Single(elements, e => e.Type == "Text");
        Assert.Contains("Master tagline", text.Text!.Content);

        // The master picture resolves through the MASTER part's rels even
        // though the same rId exists on the slide part (F1 scoping).
        var image = Assert.Single(elements, e => e.Type == "Image");
        Assert.Equal(blueBytes, File.ReadAllBytes(image.Image!.FullPath));
    }

    [Fact]
    public void Convert_LayoutHidesMasterShapes_MasterContentSkipped()
    {
        var deckPath = Path.Combine(_tempDir, $"{Guid.NewGuid():N}.pptx");
        CreateDeckWithMasterContent(deckPath, showMasterShapes: false, out _);

        using var document = PresentationDocument.Open(deckPath, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        Assert.Empty(presentation.Slides[0].Elements);
    }

    private void CreateDeckWithMasterContent(string deckPath, bool showMasterShapes, out byte[] blueBytes)
    {
        var redBytes = Convert.FromBase64String(RedPngBase64);
        blueBytes = Convert.FromBase64String(BluePngBase64);

        using var document = PresentationDocument.Create(deckPath, PresentationDocumentType.Presentation);
        var presentationPart = document.AddPresentationPart();
        presentationPart.Presentation = new Presentation
        {
            SlideMasterIdList = new SlideMasterIdList(),
            SlideIdList = new SlideIdList(),
            SlideSize = new SlideSize { Cx = 12192000, Cy = 6858000 }
        };

        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
        var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
        slideLayoutPart.SlideLayout = new P.SlideLayout(new CommonSlideData(CreateShapeTree()))
        {
            ShowMasterShapes = showMasterShapes
        };
        slideLayoutPart.AddPart(slideMasterPart);

        // Colliding rIds again: slide rId9 → red, master rId9 → blue.
        var slidePart = presentationPart.AddNewPart<SlidePart>();
        slidePart.AddPart(slideLayoutPart);
        var slideImagePart = slidePart.AddImagePart(ImagePartType.Png, "rId9");
        using (var stream = new MemoryStream(redBytes)) slideImagePart.FeedData(stream);
        var masterImagePart = slideMasterPart.AddImagePart(ImagePartType.Png, "rId9");
        using (var stream = new MemoryStream(blueBytes)) masterImagePart.FeedData(stream);

        slideMasterPart.SlideMaster = new SlideMaster(
            new CommonSlideData(CreateShapeTree(
                MasterTaglineShape(),
                LayoutPicture("rId9"))),
            CreateColorMap(),
            new SlideLayoutIdList());

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
        slidePart.Slide = new Slide(new CommonSlideData(CreateShapeTree()));
        presentationPart.Presentation.SlideIdList.Append(new SlideId
        {
            Id = 256,
            RelationshipId = presentationPart.GetIdOfPart(slidePart)
        });
    }

    private static P.Shape MasterTaglineShape()
    {
        return new P.Shape(
            new P.NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = 30, Name = "Master Tagline" },
                new P.NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(
                new Drawing.Transform2D(
                    new Drawing.Offset { X = 3000000, Y = 3000000 },
                    new Drawing.Extents { Cx = 6000000, Cy = 1000000 }),
                new Drawing.PresetGeometry(new Drawing.AdjustValueList())
                { Preset = Drawing.ShapeTypeValues.Rectangle }),
            new P.TextBody(
                new Drawing.BodyProperties(),
                new Drawing.ListStyle(),
                new Drawing.Paragraph(
                    new Drawing.Run(
                        new Drawing.RunProperties { Language = "en-US", FontSize = 3200 },
                        new Drawing.Text("Master tagline")))));
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
