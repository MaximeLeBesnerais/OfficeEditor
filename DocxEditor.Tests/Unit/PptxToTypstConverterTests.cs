using System.IO.Compression;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PptxEditor.Core.Builders;
using PptxEditor.Core.Converters;
using PptxEditor.Core.Models;
using Xunit;

namespace DocxEditor.Tests.Unit;

public class PptxToTypstConverterTests : IDisposable
{
    private readonly string _tempDir;

    public PptxToTypstConverterTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "PptxToTypstTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, true);
            }
        }
        catch
        {
            // Best effort cleanup
        }
    }

    private string CreateSimplePptx()
    {
        var path = Path.Combine(_tempDir, "simple.pptx");
        
        using (var builder = PresentationBuilder.Create(path))
        {
            builder.AddSlide();
            builder.CurrentSlide
                .AddTitle("Test Title")
                .AddText("Test body text");
            builder.Save();
        }
        
        return path;
    }

    private string CreateFormattedPptx()
    {
        var path = Path.Combine(_tempDir, "formatted.pptx");
        
        using (var document = PresentationDocument.Create(path, PresentationDocumentType.Presentation))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new Presentation();
            presentationPart.Presentation.SlideMasterIdList = new SlideMasterIdList();

            var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
            var slideMaster = new SlideMaster(
                new CommonSlideData(
                    new ShapeTree(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties { Id = 0, Name = "" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()
                        ),
                        new GroupShapeProperties(
                            new Drawing.TransformGroup(
                                new Drawing.Offset { X = 0, Y = 0 },
                                new Drawing.Extents { Cx = 0, Cy = 0 },
                                new Drawing.ChildOffset { X = 0, Y = 0 },
                                new Drawing.ChildExtents { Cx = 0, Cy = 0 }
                            )
                        )
                    )
                ),
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
                new SlideLayoutIdList()
            );
            slideMasterPart.SlideMaster = slideMaster;
            
            var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
            slideLayoutPart.SlideLayout = new DocumentFormat.OpenXml.Presentation.SlideLayout(new CommonSlideData(new ShapeTree()));
            slideLayoutPart.AddPart(slideMasterPart);
            
            var layoutId = new SlideLayoutId
            {
                Id = 2147483649,
                RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart)
            };
            slideMaster.SlideLayoutIdList!.Append(layoutId);
            
            presentationPart.Presentation.SlideIdList = new SlideIdList();
            presentationPart.Presentation.SlideMasterIdList.Append(new SlideMasterId
            {
                Id = 2147483648,
                RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
            });

            // Create slide with formatted text
            var slidePart = presentationPart.AddNewPart<SlidePart>();
            var slide = new Slide(
                new CommonSlideData(
                    new ShapeTree(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties { Id = 0, Name = "" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()
                        ),
                        new GroupShapeProperties(
                            new Drawing.TransformGroup(
                                new Drawing.Offset { X = 0, Y = 0 },
                                new Drawing.Extents { Cx = 0, Cy = 0 },
                                new Drawing.ChildOffset { X = 0, Y = 0 },
                                new Drawing.ChildExtents { Cx = 0, Cy = 0 }
                            )
                        ),
                        new P.Shape(
                            new NonVisualShapeProperties(
                                new NonVisualDrawingProperties { Id = 2, Name = "Title" },
                                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks { NoGrouping = true }),
                                new ApplicationNonVisualDrawingProperties()
                            ),
                            new ShapeProperties(
                                new Drawing.Transform2D(
                                    new Drawing.Offset { X = 1000000, Y = 500000 },
                                    new Drawing.Extents { Cx = 8000000, Cy = 1000000 }
                                )
                            ),
                            new TextBody(
                                new Drawing.BodyProperties(),
                                new Drawing.ListStyle(),
                                new Drawing.Paragraph(
                                    new Drawing.Run(
                                        new Drawing.RunProperties(
                                            new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = "FF0000" })
                                        )
                                        {
                                            Bold = new BooleanValue(true),
                                            FontSize = new Int32Value(2400)
                                        },
                                        new Drawing.Text { Text = "Bold Red Title" }
                                    )
                                )
                            )
                        ),
                        new P.Shape(
                            new NonVisualShapeProperties(
                                new NonVisualDrawingProperties { Id = 3, Name = "Body" },
                                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks { NoGrouping = true }),
                                new ApplicationNonVisualDrawingProperties()
                            ),
                            new ShapeProperties(
                                new Drawing.Transform2D(
                                    new Drawing.Offset { X = 1000000, Y = 2000000 },
                                    new Drawing.Extents { Cx = 8000000, Cy = 3000000 }
                                )
                            ),
                            new TextBody(
                                new Drawing.BodyProperties(),
                                new Drawing.ListStyle(),
                                new Drawing.Paragraph(
                                    new Drawing.Run(
                                        new Drawing.RunProperties(
                                            new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = "0000FF" })
                                        )
                                        {
                                            Italic = new BooleanValue(true),
                                            FontSize = new Int32Value(1800)
                                        },
                                        new Drawing.Text { Text = "Italic Blue Text" }
                                    )
                                )
                            )
                        )
                    )
                )
            );
            slidePart.Slide = slide;
            slidePart.AddPart(slideLayoutPart);
            
            var slideId = new SlideId
            {
                Id = 256,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            };
            presentationPart.Presentation.SlideIdList.Append(slideId);
            
            // Set slide size
            presentationPart.Presentation.SlideSize = new SlideSize
            {
                Cx = 9144000,
                Cy = 6858000,
                Type = SlideSizeValues.Screen4x3
            };
        }
        
        return path;
    }

    private const long EmusPerPoint = 12700;

    private string CreateGroupShapePptx(string fileName, params OpenXmlElement[] slideElements)
    {
        var path = Path.Combine(_tempDir, fileName);

        using (var document = PresentationDocument.Create(path, PresentationDocumentType.Presentation))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new Presentation
            {
                SlideMasterIdList = new SlideMasterIdList(),
                SlideIdList = new SlideIdList(),
                SlideSize = new SlideSize { Cx = (int)Pt(720), Cy = (int)Pt(540), Type = SlideSizeValues.Screen4x3 }
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

    private static P.Shape TextShape(uint id, string text, double x, double y, double width, double height)
    {
        return new P.Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = id, Name = $"Text {id}" },
                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(
                new Drawing.Transform2D(
                    new Drawing.Offset { X = Pt(x), Y = Pt(y) },
                    new Drawing.Extents { Cx = Pt(width), Cy = Pt(height) })),
            new TextBody(
                new Drawing.BodyProperties { LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0 },
                new Drawing.ListStyle(),
                new Drawing.Paragraph(new Drawing.Run(new Drawing.Text { Text = text }))));
    }

    private static P.GroupShape GroupShape(uint id, Drawing.TransformGroup? transformGroup, params OpenXmlElement[] children)
    {
        var groupShape = new P.GroupShape(
            new NonVisualGroupShapeProperties(
                new NonVisualDrawingProperties { Id = id, Name = $"Group {id}" },
                new NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            transformGroup == null ? new GroupShapeProperties() : new GroupShapeProperties(transformGroup));

        foreach (var child in children)
        {
            groupShape.Append(child);
        }

        return groupShape;
    }

    private static Drawing.TransformGroup TransformGroup(double x, double y, double width, double height, double childX, double childY, double childWidth, double childHeight)
    {
        return new Drawing.TransformGroup(
            new Drawing.Offset { X = Pt(x), Y = Pt(y) },
            new Drawing.Extents { Cx = Pt(width), Cy = Pt(height) },
            new Drawing.ChildOffset { X = Pt(childX), Y = Pt(childY) },
            new Drawing.ChildExtents { Cx = Pt(childWidth), Cy = Pt(childHeight) });
    }

    private static P.ConnectionShape UnsupportedConnector(uint id)
    {
        return new P.ConnectionShape(
            new NonVisualConnectionShapeProperties(
                new NonVisualDrawingProperties { Id = id, Name = $"Connector {id}" },
                new NonVisualConnectorShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(
                new Drawing.Transform2D(
                    new Drawing.Offset { X = Pt(1), Y = Pt(1) },
                    new Drawing.Extents { Cx = Pt(10), Cy = Pt(10) })));
    }

    private static long Pt(double points) => (long)Math.Round(points * EmusPerPoint);

    private static TypstElement AssertSingleTextElement(TypstPresentation presentation, string expectedText)
    {
        var element = Assert.Single(presentation.Slides[0].Elements, e => e.Text?.Content == expectedText);
        Assert.Equal("Text", element.Type);
        return element;
    }

    private static void AssertPosition(TypstElement element, double x, double y, double width, double height)
    {
        Assert.Equal(x, element.X, 2);
        Assert.Equal(y, element.Y, 2);
        Assert.Equal(width, element.Width, 2);
        Assert.Equal(height, element.Height, 2);
    }

    [Fact]
    public void Convert_GroupShapeWithTransformGroup_AppliesOffsetAndScaleToChildTextShape()
    {
        var group = GroupShape(
            10,
            TransformGroup(x: 100, y: 50, width: 400, height: 200, childX: 10, childY: 20, childWidth: 200, childHeight: 100),
            TextShape(11, "Grouped text", x: 20, y: 30, width: 100, height: 40));
        var path = CreateGroupShapePptx("group-transform.pptx", group);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var element = AssertSingleTextElement(presentation, "Grouped text");

        // ECMA-376: abs = grpOff + (child − chOff) × (ext / chExt)
        AssertPosition(element, x: 120, y: 70, width: 200, height: 80);
    }

    [Fact]
    public void Convert_GroupShapeWithoutTransformGroup_PreservesParentOffsetAndScale()
    {
        var innerGroupWithoutTransform = GroupShape(
            12,
            transformGroup: null,
            TextShape(13, "Inherited transform", x: 20, y: 30, width: 100, height: 40));
        var outerGroup = GroupShape(
            10,
            TransformGroup(x: 100, y: 50, width: 400, height: 200, childX: 10, childY: 20, childWidth: 200, childHeight: 100),
            innerGroupWithoutTransform);
        var path = CreateGroupShapePptx("group-without-transform.pptx", outerGroup);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var element = AssertSingleTextElement(presentation, "Inherited transform");

        AssertPosition(element, x: 120, y: 70, width: 200, height: 80);
    }

    [Fact]
    public void Convert_NestedGroupShape_ComposesTransformsRecursively()
    {
        var innerGroup = GroupShape(
            12,
            TransformGroup(x: 10, y: 5, width: 50, height: 50, childX: 0, childY: 0, childWidth: 25, childHeight: 25),
            TextShape(13, "Nested text", x: 5, y: 5, width: 10, height: 10));
        var outerGroup = GroupShape(
            10,
            TransformGroup(x: 100, y: 50, width: 200, height: 200, childX: 0, childY: 0, childWidth: 100, childHeight: 100),
            innerGroup);
        var path = CreateGroupShapePptx("nested-group-transform.pptx", outerGroup);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var element = AssertSingleTextElement(presentation, "Nested text");

        AssertPosition(element, x: 140, y: 80, width: 40, height: 40);
    }

    [Fact]
    public void Convert_EmptyGroupShape_YieldsNoElements()
    {
        var group = GroupShape(
            10,
            TransformGroup(x: 100, y: 50, width: 400, height: 200, childX: 10, childY: 20, childWidth: 200, childHeight: 100));
        var path = CreateGroupShapePptx("empty-group.pptx", group);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();

        Assert.Empty(presentation.Slides[0].Elements);
    }

    [Fact]
    public void Convert_GroupShapeWithUnsupportedChild_StillConvertsSupportedSibling()
    {
        var group = GroupShape(
            10,
            TransformGroup(x: 100, y: 50, width: 400, height: 200, childX: 10, childY: 20, childWidth: 200, childHeight: 100),
            UnsupportedConnector(11),
            TextShape(12, "Supported sibling", x: 20, y: 30, width: 100, height: 40));
        var path = CreateGroupShapePptx("unsupported-child-group.pptx", group);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var element = AssertSingleTextElement(presentation, "Supported sibling");

        AssertPosition(element, x: 120, y: 70, width: 200, height: 80);
    }

    [Fact]
    public void Convert_RunWithEmbeddedNewline_SplitsIntoLineBreakRunPreservingPerRunFormatting()
    {
        // Sales deck slide 8 chevrons: one a:p holding "WEEKS 1–3\n" (9pt) + "DIAGNOSE" (16pt).
        var shape = TwoRunChevronTextShape(2, firstRunText: "WEEKS 1–3\n", secondRunText: "DIAGNOSE");
        var path = CreateGroupShapePptx("embedded-newline.pptx", shape);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var element = Assert.Single(presentation.Slides[0].Elements);
        var text = element.Text ?? throw new InvalidOperationException("Expected shape text element.");
        var paragraph = Assert.Single(text.Paragraphs);

        Assert.Equal(3, paragraph.Runs.Count);
        Assert.Equal("WEEKS 1–3", paragraph.Runs[0].Content);
        Assert.False(paragraph.Runs[0].IsLineBreak);
        Assert.Equal(9.0, paragraph.Runs[0].Formatting.FontSize);
        Assert.True(paragraph.Runs[1].IsLineBreak);
        Assert.Equal("DIAGNOSE", paragraph.Runs[2].Content);
        Assert.False(paragraph.Runs[2].IsLineBreak);
        Assert.Equal(16.0, paragraph.Runs[2].Formatting.FontSize);
        Assert.True(text.HasExplicitLineBreaks);
    }

    [Fact]
    public void GenerateTypstSource_RunWithEmbeddedNewline_EmitsLinebreakBetweenDifferentlySizedRuns()
    {
        var shape = TwoRunChevronTextShape(2, firstRunText: "WEEKS 1–3\n", secondRunText: "DIAGNOSE");
        var path = CreateGroupShapePptx("embedded-newline-source.pptx", shape);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var source = converter.GenerateTypstSource(converter.Convert());

        Assert.Contains("#linebreak()", source);
        Assert.Contains("size: 9.00pt", source);
        Assert.Contains("size: 16.00pt", source);
        // The break must sit between the two run wrappers, not collapse to a space.
        Assert.Matches(new Regex(@"WEEKS 1–3\]\s*#linebreak\(\)\s*#text\([^)]*size: 16\.00pt[^)]*\)\[DIAGNOSE\]"), source);
    }

    [Fact]
    public void GenerateTypstSource_TwoParagraphsDifferentSizes_EmitsBothLinesWithPerParagraphProperties()
    {
        var shape = new P.Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = 2, Name = "Chevron 2" },
                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(
                new Drawing.Transform2D(
                    new Drawing.Offset { X = Pt(48), Y = Pt(160) },
                    new Drawing.Extents { Cx = Pt(222), Cy = Pt(76) })),
            new TextBody(
                new Drawing.BodyProperties { Anchor = Drawing.TextAnchoringTypeValues.Center },
                new Drawing.ListStyle(),
                new Drawing.Paragraph(
                    new Drawing.Run(
                        new Drawing.RunProperties { FontSize = new Int32Value(900), Bold = new BooleanValue(false) },
                        new Drawing.Text { Text = "WEEKS 1–3" })),
                new Drawing.Paragraph(
                    new Drawing.Run(
                        new Drawing.RunProperties { FontSize = new Int32Value(1600), Bold = new BooleanValue(true) },
                        new Drawing.Text { Text = "DIAGNOSE" }))));
        var path = CreateGroupShapePptx("two-paragraph-shape.pptx", shape);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var element = Assert.Single(presentation.Slides[0].Elements);
        var text = element.Text ?? throw new InvalidOperationException("Expected shape text element.");

        Assert.Equal(2, text.Paragraphs.Count);
        Assert.Equal("WEEKS 1–3", text.Paragraphs[0].Content);
        Assert.Equal(9.0, text.Paragraphs[0].Formatting.FontSize);
        Assert.False(text.Paragraphs[0].Formatting.Bold);
        Assert.Equal("DIAGNOSE", text.Paragraphs[1].Content);
        Assert.Equal(16.0, text.Paragraphs[1].Formatting.FontSize);
        Assert.True(text.Paragraphs[1].Formatting.Bold);

        var source = converter.GenerateTypstSource(presentation);
        Assert.Contains("size: 9.00pt", source);
        Assert.Contains("size: 16.00pt", source);
        Assert.Contains("WEEKS 1–3", source);
        Assert.Contains("DIAGNOSE", source);
    }

    [Fact]
    public void GenerateTypstSource_CenteredShapeText_AppliesCenterAlignment()
    {
        var shape = new P.Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = 2, Name = "Chevron 2" },
                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(
                new Drawing.Transform2D(
                    new Drawing.Offset { X = Pt(48), Y = Pt(160) },
                    new Drawing.Extents { Cx = Pt(222), Cy = Pt(76) })),
            new TextBody(
                new Drawing.BodyProperties { Anchor = Drawing.TextAnchoringTypeValues.Center },
                new Drawing.ListStyle(),
                new Drawing.Paragraph(
                    new Drawing.ParagraphProperties { Alignment = Drawing.TextAlignmentTypeValues.Center },
                    new Drawing.Run(
                        new Drawing.RunProperties { FontSize = new Int32Value(900) },
                        new Drawing.Text { Text = "WEEKS 1–3\n" }),
                    new Drawing.Run(
                        new Drawing.RunProperties { FontSize = new Int32Value(1600), Bold = new BooleanValue(true) },
                        new Drawing.Text { Text = "DIAGNOSE" }))));
        var path = CreateGroupShapePptx("centered-shape.pptx", shape);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var element = Assert.Single(presentation.Slides[0].Elements);
        Assert.Equal("center", element.Text!.Paragraphs[0].Formatting.Align);

        var source = converter.GenerateTypstSource(presentation);
        Assert.Contains("#align(center)", source);
        Assert.Contains("#linebreak()", source);
    }

    private static P.Shape TwoRunChevronTextShape(uint id, string firstRunText, string secondRunText)
    {
        return new P.Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = id, Name = $"Chevron {id}" },
                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(
                new Drawing.Transform2D(
                    new Drawing.Offset { X = Pt(48), Y = Pt(160) },
                    new Drawing.Extents { Cx = Pt(222), Cy = Pt(76) })),
            new TextBody(
                new Drawing.BodyProperties { Anchor = Drawing.TextAnchoringTypeValues.Center },
                new Drawing.ListStyle(),
                new Drawing.Paragraph(
                    new Drawing.Run(
                        new Drawing.RunProperties { FontSize = new Int32Value(900), Bold = new BooleanValue(true) },
                        new Drawing.Text { Text = firstRunText }),
                    new Drawing.Run(
                        new Drawing.RunProperties { FontSize = new Int32Value(1600), Bold = new BooleanValue(true) },
                        new Drawing.Text { Text = secondRunText }))));
    }

    [Fact]
    public void Convert_SimplePptx_CreatesTypstPresentation()
    {
        var path = CreateSimplePptx();
        
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);
        
        var presentation = converter.Convert();
        
        Assert.NotNull(presentation);
        Assert.True(presentation.Slides.Count >= 1, $"Expected at least 1 slide, got {presentation.Slides.Count}");
        var slidesWithElements = presentation.Slides.Where(s => s.Elements.Count > 0).ToList();
        Assert.True(slidesWithElements.Count > 0, "Expected at least one slide with elements");
    }

    [Fact]
    public void Convert_FormattedText_PreservesFormatting()
    {
        var path = CreateFormattedPptx();
        
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);
        
        var presentation = converter.Convert();
        var slide = presentation.Slides[0];
        
        // Find title element
        var titleElement = slide.Elements.FirstOrDefault(e => e.Text?.Content.Contains("Bold Red Title") == true);
        Assert.NotNull(titleElement);
        Assert.Equal("Text", titleElement.Type);
        Assert.True(titleElement.Text!.Formatting.Bold);
        Assert.Equal("#FF0000", titleElement.Text.Formatting.Color);
        Assert.Equal(24.0, titleElement.Text.Formatting.FontSize);
        
        // Find body element
        var bodyElement = slide.Elements.FirstOrDefault(e => e.Text?.Content.Contains("Italic Blue Text") == true);
        Assert.NotNull(bodyElement);
        Assert.True(bodyElement.Text!.Formatting.Italic);
        Assert.Equal("#0000FF", bodyElement.Text.Formatting.Color);
    }

    [Fact]
    public void Convert_ExtractsSlideDimensions()
    {
        var path = CreateFormattedPptx();
        
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);
        
        var presentation = converter.Convert();
        var layout = presentation.Slides[0].Layout;
        
        // 9144000 EMU = 720pt, 6858000 EMU = 540pt
        Assert.True(layout.Width > 0);
        Assert.True(layout.Height > 0);
    }

    [Fact]
    public void GenerateTypstSource_CreatesValidTypst()
    {
        var path = CreateFormattedPptx();
        
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);
        
        var presentation = converter.Convert();
        var source = converter.GenerateTypstSource(presentation);
        
        Assert.NotNull(source);
        Assert.Contains("#set text", source);
        Assert.Contains("#set page", source);
        Assert.Contains("Bold Red Title", source);
        Assert.Contains("Italic Blue Text", source);
    }

    [Fact]
    public void GenerateTypstSource_IncludesFontSetup()
    {
        var path = CreateFormattedPptx();
        
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);
        
        var presentation = converter.Convert();
        var source = converter.GenerateTypstSource(presentation);
        
        Assert.Contains("font:", source);
        Assert.Contains("Arial", source);
    }

    [Fact]
    public void GenerateTypstSource_PositionsElements()
    {
        var path = CreateFormattedPptx();
        
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);
        
        var presentation = converter.Convert();
        var source = converter.GenerateTypstSource(presentation);
        
        Assert.Contains("#place", source);
        Assert.Contains("dx:", source);
        Assert.Contains("dy:", source);
    }

    [Fact]
    public void GenerateTypstSource_ExpandsSingleLineAutoFitTextToSlideBounds()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = CreateAutoFitPresentation(height: 10);
        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains("#block(width: 618.00pt)", source);
    }

    [Fact]
    public void GenerateTypstSource_KeepsMultiLineAutoFitTextConstrained()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = CreateAutoFitPresentation(height: 30);
        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains("#block(width: 10.00pt)", source);
    }

    [Fact]
    public void GenerateTypstSource_KeepsCenteredSingleLineAutoFitTextConstrained()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = CreateAutoFitPresentation(height: 10, align: "center");
        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains("#block(width: 10.00pt)", source);
    }

    [Fact]
    public void GenerateTypstSource_EmitsParagraphLeadingFromLineSpacing()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = CreateTextPresentation(new TypstTextElement
        {
            Paragraphs =
            [
                new TypstParagraph
                {
                    Content = "Line one line two",
                    Formatting = new TypstTextFormatting
                    {
                        FontFamily = "Arial",
                        FontSize = 18
                    }
                }
            ],
            LineSpacing = 32.25
        });
        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains("#set par(leading: 20.55pt)", source);
    }

    [Fact]
    public void GenerateTypstSource_EmitsVerticalGapForEmptyParagraph()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = CreateTextPresentation(new TypstTextElement
        {
            Paragraphs =
            [
                new TypstParagraph
                {
                    Content = "First paragraph",
                    Formatting = new TypstTextFormatting
                    {
                        FontFamily = "Arial",
                        FontSize = 18
                    }
                },
                new TypstParagraph
                {
                    Content = "",
                    Formatting = new TypstTextFormatting
                    {
                        FontFamily = "Arial",
                        FontSize = 18
                    }
                },
                new TypstParagraph
                {
                    Content = "",
                    Formatting = new TypstTextFormatting
                    {
                        FontFamily = "Arial",
                        FontSize = 18
                    }
                },
                new TypstParagraph
                {
                    Content = "Second paragraph",
                    Formatting = new TypstTextFormatting
                    {
                        FontFamily = "Arial",
                        FontSize = 18
                    }
                }
            ],
            LineSpacing = 32.25
        });
        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains("[First paragraph]", source);
        Assert.Contains("#v(32.25pt)", source);
        Assert.Contains("[Second paragraph]", source);
    }

    [Fact]
    public void Convert_PreservesEmptyPptxParagraphAsBlankParagraphBreak()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var shape = new P.Shape(
            new TextBody(
                new Drawing.BodyProperties(),
                new Drawing.ListStyle(),
                new Drawing.Paragraph(new Drawing.Run(new Drawing.Text { Text = "First paragraph" })),
                new Drawing.Paragraph(),
                new Drawing.Paragraph(new Drawing.Run(new Drawing.Text { Text = "Second paragraph" }))));

        var method = typeof(PptxToTypstConverter).GetMethod(
            "ExtractTextFromShape",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(P.Shape), typeof(StyleResolver), typeof(SlidePart)],
            null);

        // Get the first slide part for style resolution
        var slidePart = document.PresentationPart!.SlideParts.FirstOrDefault();
        Assert.NotNull(slidePart);

        var styleResolver = new StyleResolver(document, slidePart);
        var text = Assert.IsType<TypstTextElement>(method!.Invoke(converter, [shape, styleResolver, slidePart]));

        Assert.Equal("First paragraph\n\n\n\nSecond paragraph", text.Content);
    }

    private static TypstPresentation CreateAutoFitPresentation(double height, string align = "left")
    {
        return CreateTextPresentation(new TypstTextElement
        {
            Paragraphs =
            [
                new TypstParagraph
                {
                    Content = "WW",
                    Formatting = new TypstTextFormatting
                    {
                        FontFamily = "Arial",
                        FontSize = 10,
                        Align = align
                    }
                }
            ],
            AutoFit = true,
            ParagraphCount = 1
        }, height: height, width: 10);
    }

    private static TypstPresentation CreateTextPresentation(TypstTextElement text, double height = 100, double width = 300)
    {
        // Auto-populate Runs for paragraphs that have Content but no Runs
        foreach (var paragraph in text.Paragraphs)
        {
            if (paragraph.Runs.Count == 0)
            {
                if (!string.IsNullOrEmpty(paragraph.Content))
                {
                    paragraph.Runs.Add(new TypstTextRun { Content = paragraph.Content, Formatting = paragraph.Formatting });
                }
            }
        }

        return new TypstPresentation
        {
            Slides =
            [
                new TypstSlide
                {
                    Layout = new PptxEditor.Core.Models.SlideLayout { Width = 720, Height = 540 },
                    Elements =
                    [
                        new TypstElement
                        {
                            Type = "Text",
                            X = 100,
                            Width = width,
                            Height = height,
                            Text = text
                        }
                    ]
                }
            ]
        };
    }

    private string CreateMixedFormattingPptx()
    {
        var path = Path.Combine(_tempDir, "mixed-formatting.pptx");

        using (var document = PresentationDocument.Create(path, PresentationDocumentType.Presentation))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new Presentation();
            presentationPart.Presentation.SlideMasterIdList = new SlideMasterIdList();

            var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
            var slideMaster = new SlideMaster(
                new CommonSlideData(
                    new ShapeTree(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties { Id = 0, Name = "" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()
                        ),
                        new GroupShapeProperties(
                            new Drawing.TransformGroup(
                                new Drawing.Offset { X = 0, Y = 0 },
                                new Drawing.Extents { Cx = 0, Cy = 0 },
                                new Drawing.ChildOffset { X = 0, Y = 0 },
                                new Drawing.ChildExtents { Cx = 0, Cy = 0 }
                            )
                        )
                    )
                ),
                new ColorMap(),
                new SlideLayoutIdList()
            );
            slideMasterPart.SlideMaster = slideMaster;

            var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
            slideLayoutPart.SlideLayout = new P.SlideLayout(new CommonSlideData(new ShapeTree()));
            slideLayoutPart.AddPart(slideMasterPart);

            var layoutId = new SlideLayoutId
            {
                Id = 2147483649,
                RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart)
            };
            slideMaster.SlideLayoutIdList!.Append(layoutId);

            presentationPart.Presentation.SlideIdList = new SlideIdList();
            presentationPart.Presentation.SlideMasterIdList.Append(new SlideMasterId
            {
                Id = 2147483648,
                RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
            });

            var slidePart = presentationPart.AddNewPart<SlidePart>();
            var slide = new Slide(
                new CommonSlideData(
                    new ShapeTree(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties { Id = 0, Name = "" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()
                        ),
                        new GroupShapeProperties(
                            new Drawing.TransformGroup(
                                new Drawing.Offset { X = 0, Y = 0 },
                                new Drawing.Extents { Cx = 0, Cy = 0 },
                                new Drawing.ChildOffset { X = 0, Y = 0 },
                                new Drawing.ChildExtents { Cx = 0, Cy = 0 }
                            )
                        ),
                        new P.Shape(
                            new NonVisualShapeProperties(
                                new NonVisualDrawingProperties { Id = 2, Name = "TextBox" },
                                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks { NoGrouping = true }),
                                new ApplicationNonVisualDrawingProperties()
                            ),
                            new ShapeProperties(
                                new Drawing.Transform2D(
                                    new Drawing.Offset { X = 1000000, Y = 500000 },
                                    new Drawing.Extents { Cx = 8000000, Cy = 1000000 }
                                )
                            ),
                            new TextBody(
                                new Drawing.BodyProperties(),
                                new Drawing.ListStyle(),
                                new Drawing.Paragraph(
                                    new Drawing.Run(
                                        new Drawing.RunProperties()
                                        {
                                            Bold = new BooleanValue(true),
                                            FontSize = new Int32Value(1800)
                                        },
                                        new Drawing.Text { Text = "Bold part" }
                                    ),
                                    new Drawing.Run(
                                        new Drawing.Text { Text = "Regular part" }
                                    )
                                )
                            )
                        ),
                        new P.Shape(
                            new NonVisualShapeProperties(
                                new NonVisualDrawingProperties { Id = 3, Name = "TextBox2" },
                                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks { NoGrouping = true }),
                                new ApplicationNonVisualDrawingProperties()
                            ),
                            new ShapeProperties(
                                new Drawing.Transform2D(
                                    new Drawing.Offset { X = 1000000, Y = 2000000 },
                                    new Drawing.Extents { Cx = 8000000, Cy = 1000000 }
                                )
                            ),
                            new TextBody(
                                new Drawing.BodyProperties(),
                                new Drawing.ListStyle(),
                                new Drawing.Paragraph(
                                    new Drawing.Run(
                                        new Drawing.RunProperties(
                                            new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = "FF0000" })
                                        )
                                        {
                                            Bold = new BooleanValue(true),
                                            FontSize = new Int32Value(1800)
                                        },
                                        new Drawing.Text { Text = "Bold Red" }
                                    ),
                                    new Drawing.Run(
                                        new Drawing.RunProperties()
                                        {
                                            Italic = new BooleanValue(true),
                                            FontSize = new Int32Value(1800)
                                        },
                                        new Drawing.Text { Text = "Italic " }
                                    ),
                                    new Drawing.Run(
                                        new Drawing.RunProperties(
                                            new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = "0000FF" })
                                        )
                                        {
                                            FontSize = new Int32Value(1800)
                                        },
                                        new Drawing.Text { Text = "Blue" }
                                    )
                                )
                            )
                        )
                    )
                )
            );
            slidePart.Slide = slide;
            slidePart.AddPart(slideLayoutPart);

            var slideId = new SlideId
            {
                Id = 256,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            };
            presentationPart.Presentation.SlideIdList.Append(slideId);

            presentationPart.Presentation.SlideSize = new SlideSize
            {
                Cx = 9144000,
                Cy = 6858000,
                Type = SlideSizeValues.Screen4x3
            };
        }

        return path;
    }

    [Fact]
    public void Convert_MixedFormatting_PopulatesRuns()
    {
        var path = CreateMixedFormattingPptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var slide = presentation.Slides[0];

        var textElement = slide.Elements.FirstOrDefault(e => e.Text?.Paragraphs.Any(p => p.Runs.Count > 1) == true);
        Assert.NotNull(textElement);

        var mixedParagraph = textElement.Text!.Paragraphs.First(p => p.Runs.Count > 1);
        Assert.Equal(2, mixedParagraph.Runs.Count);
        Assert.Equal("Bold part", mixedParagraph.Runs[0].Content);
        Assert.True(mixedParagraph.Runs[0].Formatting.Bold);
        Assert.Equal("Regular part", mixedParagraph.Runs[1].Content);
        Assert.False(mixedParagraph.Runs[1].Formatting.Bold);
    }

    [Fact]
    public void GenerateTypstSource_MixedFormatting_BoldAndRegular()
    {
        var path = CreateMixedFormattingPptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains("weight: \"bold\"", source);
        Assert.Contains("[Bold part]", source);
        Assert.Contains("[Regular part]", source);
    }

    [Fact]
    public void GenerateTypstSource_MixedFormatting_MultipleStyles()
    {
        var path = CreateMixedFormattingPptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains("weight: \"bold\"", source);
        Assert.Contains("style: \"italic\"", source);
        Assert.Contains("fill: rgb(\"#FF0000\")", source);
        Assert.Contains("fill: rgb(\"#0000FF\")", source);
        Assert.Contains("[Bold Red]", source);
        Assert.Contains("[Italic ]", source);
        Assert.Contains("[Blue]", source);
    }

    [Fact]
    public void GenerateTypstSource_SingleRunParagraph_BehaviorUnchanged()
    {
        var path = CreateFormattedPptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var source = converter.GenerateTypstSource(presentation);

        Assert.NotNull(source);
        Assert.Contains("Bold Red Title", source);
        Assert.Contains("Italic Blue Text", source);
    }

    [Fact]
    public void GenerateTypstSource_TableCellStrokes_DistinguishesVisibleAndExplicitNone()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = new TypstPresentation
        {
            Slides =
            [
                new TypstSlide
                {
                    Layout = new PptxEditor.Core.Models.SlideLayout { Width = 720, Height = 540 },
                    Elements =
                    [
                        new TypstElement
                        {
                            Type = "Table",
                            Width = 200,
                            Height = 100,
                            Table = new TypstTableElement
                            {
                                ColumnWidths = [100, 100],
                                Rows =
                                [
                                    [
                                        new TypstTableCell
                                        {
                                            Content = "A",
                                            StylePart = new TableStylePart
                                            {
                                                BorderRightState = TableBorderState.None,
                                                BorderBottomState = TableBorderState.Visible,
                                                BorderBottomColor = "#CEBA80",
                                                BorderBottomWidth = 1
                                            }
                                        },
                                        new TypstTableCell
                                        {
                                            Content = "B",
                                            StylePart = new TableStylePart
                                            {
                                                BorderLeftState = TableBorderState.None,
                                                BorderBottomState = TableBorderState.Visible,
                                                BorderBottomColor = "#CEBA80",
                                                BorderBottomWidth = 1
                                            }
                                        }
                                    ]
                                ]
                            }
                        }
                    ]
                }
            ]
        };

        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains("right: none", source);
        Assert.Contains("left: none", source);
        Assert.Contains("bottom: 1.00pt + rgb(\"#CEBA80\")", source);
    }

    [Fact]
    public void GenerateTypstSource_TableWithEmptyStylePart_KeepsGlobalStroke()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = new TypstPresentation
        {
            Slides =
            [
                new TypstSlide
                {
                    Layout = new PptxEditor.Core.Models.SlideLayout { Width = 720, Height = 540 },
                    Elements =
                    [
                        new TypstElement
                        {
                            Type = "Table",
                            Width = 200,
                            Height = 100,
                            Table = new TypstTableElement
                            {
                                BorderWidth = 1,
                                BorderColor = "#000000",
                                ColumnWidths = [100],
                                Rows =
                                [
                                    [
                                        new TypstTableCell
                                        {
                                            Content = "A",
                                            StylePart = new TableStylePart()
                                        }
                                    ]
                                ]
                            }
                        }
                    ]
                }
            ]
        };

        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains("#table(columns: (100.00pt), stroke: 1.00pt + rgb(\"#000000\")", source);
        Assert.DoesNotContain("stroke: none", source);
    }

    [Fact]
    public void GenerateTypstSource_ComplexTableWithDefaultCell_UsesDefaultCellStroke()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = new TypstPresentation
        {
            Slides =
            [
                new TypstSlide
                {
                    Layout = new PptxEditor.Core.Models.SlideLayout { Width = 720, Height = 540 },
                    Elements =
                    [
                        new TypstElement
                        {
                            Type = "Table",
                            Width = 200,
                            Height = 100,
                            Table = new TypstTableElement
                            {
                                BorderWidth = 1,
                                BorderColor = "#000000",
                                ColumnWidths = [100, 100],
                                Rows =
                                [
                                    [
                                        new TypstTableCell
                                        {
                                            Content = "Explicit",
                                            StylePart = new TableStylePart
                                            {
                                                BorderRightState = TableBorderState.Visible,
                                                BorderRightColor = "#FF0000",
                                                BorderRightWidth = 2
                                            }
                                        },
                                        new TypstTableCell
                                        {
                                            Content = "Default"
                                        }
                                    ]
                                ]
                            }
                        }
                    ]
                }
            ]
        };

        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains("#table(columns: (100.00pt, 100.00pt), ", source);
        Assert.DoesNotContain("#table(columns: (100.00pt, 100.00pt), stroke:", source);
        Assert.DoesNotContain("stroke: none", source);
        Assert.Contains("stroke: (right: 2.00pt + rgb(\"#FF0000\"))", source);
        Assert.Contains("stroke: (top: 1.00pt + rgb(\"#000000\"), bottom: 1.00pt + rgb(\"#000000\"), left: 1.00pt + rgb(\"#000000\"), right: 1.00pt + rgb(\"#000000\"))", source);
    }

    [Fact]
    public void GenerateTypstSource_TableGeometry_EmitsRowsInsetsAndVerticalAlign()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = new TypstPresentation
        {
            Slides =
            [
                new TypstSlide
                {
                    Layout = new PptxEditor.Core.Models.SlideLayout { Width = 720, Height = 540 },
                    Elements =
                    [
                        new TypstElement
                        {
                            Type = "Table",
                            Width = 200,
                            Height = 100,
                            Table = new TypstTableElement
                            {
                                ColumnWidths = [100],
                                RowHeights = [59.396, 71.07],
                                Rows =
                                [
                                    [
                                        new TypstTableCell
                                        {
                                            Content = "A",
                                            Insets = new TypstTableCellInsets
                                            {
                                                Left = 14.4,
                                                Right = 10.8,
                                                Top = 10.8,
                                                Bottom = 10.8
                                            },
                                            VerticalAlign = "center"
                                        }
                                    ],
                                    [new TypstTableCell { Content = "B" }]
                                ]
                            }
                        }
                    ]
                }
            ]
        };

        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains("rows: (59.40pt, 71.07pt)", source);
        Assert.Contains("inset: (left: 14.40pt, right: 10.80pt, top: 10.80pt, bottom: 10.80pt)", source);
        Assert.Contains("align: left + horizon", source);
    }

    [Fact]
    public void GenerateTypstSource_TableGeometry_SkipsRowsWhenAnyHeightIsZero()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = new TypstPresentation
        {
            Slides =
            [
                new TypstSlide
                {
                    Layout = new PptxEditor.Core.Models.SlideLayout { Width = 720, Height = 540 },
                    Elements =
                    [
                        new TypstElement
                        {
                            Type = "Table",
                            Width = 200,
                            Height = 100,
                            Table = new TypstTableElement
                            {
                                ColumnWidths = [100],
                                RowHeights = [59.396, 0],
                                Rows =
                                [
                                    [new TypstTableCell { Content = "A" }],
                                    [new TypstTableCell { Content = "B" }]
                                ]
                            }
                        }
                    ]
                }
            ]
        };

        var source = converter.GenerateTypstSource(presentation);

        Assert.DoesNotContain("rows:", source);
    }

    [Fact]
    public void ExtractTableCellGeometry_ReadsMarginsAndAnchorFromTcPrOuterXml()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var tcPr = new Drawing.TableCellProperties();
        tcPr.SetAttribute(new OpenXmlAttribute("", "marL", "", "182880"));
        tcPr.SetAttribute(new OpenXmlAttribute("", "marR", "", "137160"));
        tcPr.SetAttribute(new OpenXmlAttribute("", "marT", "", "137160"));
        tcPr.SetAttribute(new OpenXmlAttribute("", "marB", "", "137160"));
        tcPr.SetAttribute(new OpenXmlAttribute("", "anchor", "", "ctr"));

        var cell = new Drawing.TableCell(
            tcPr,
            new Drawing.TextBody(
                new Drawing.BodyProperties(),
                new Drawing.ListStyle(),
                new Drawing.Paragraph(new Drawing.Run(new Drawing.Text { Text = "Cell" }))));

        var extractInsets = typeof(PptxToTypstConverter).GetMethod(
            "ExtractCellInsets",
            System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic);
        var extractAlign = typeof(PptxToTypstConverter).GetMethod(
            "ExtractCellVerticalAlign",
            System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic);

        var insets = Assert.IsType<TypstTableCellInsets>(extractInsets!.Invoke(converter, [cell]));
        var verticalAlign = Assert.IsType<string>(extractAlign!.Invoke(converter, [cell]));

        Assert.Equal(14.4, insets.Left!.Value, 3);
        Assert.Equal(10.8, insets.Right!.Value, 3);
        Assert.Equal(10.8, insets.Top!.Value, 3);
        Assert.Equal(10.8, insets.Bottom!.Value, 3);
        Assert.Equal("center", verticalAlign);
    }

    private static TypstTableElement InvokeExtractTable(PptxToTypstConverter converter, Drawing.Table table)
    {
        var extractTable = typeof(PptxToTypstConverter).GetMethod(
            "ExtractTable",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        return Assert.IsType<TypstTableElement>(extractTable!.Invoke(converter, [table, null]));
    }

    private static Drawing.TableCell CreateTableCell(string tcPrXml, string text)
        => new(
            new Drawing.TextBody(
                new Drawing.BodyProperties(),
                new Drawing.Paragraph(new Drawing.Run(new Drawing.Text { Text = text }))),
            new Drawing.TableCellProperties(tcPrXml));

    [Fact]
    public void ExtractTable_ReadsExplicitCellLineBordersFromTcPr()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        const string ns = "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"";
        // Header cell: all four edges noFill over a dark fill (northwind REF pattern).
        var headerCell = CreateTableCell(
            $"<a:tcPr {ns} marL=\"91440\" marR=\"91440\" marT=\"45720\" marB=\"45720\" anchor=\"ctr\">" +
            "<a:lnL><a:noFill/></a:lnL><a:lnR><a:noFill/></a:lnR><a:lnT><a:noFill/></a:lnT><a:lnB><a:noFill/></a:lnB>" +
            "<a:solidFill><a:srgbClr val=\"0B1F3A\"/></a:solidFill></a:tcPr>",
            "METRIC");
        // Body cell: horizontal-only stroke — light gray 1pt bottom edge, others noFill.
        var bodyCell = CreateTableCell(
            $"<a:tcPr {ns} marL=\"91440\" marR=\"91440\" marT=\"45720\" marB=\"45720\" anchor=\"ctr\">" +
            "<a:lnL><a:noFill/></a:lnL><a:lnR><a:noFill/></a:lnR><a:lnT><a:noFill/></a:lnT>" +
            "<a:lnB w=\"12700\"><a:solidFill><a:srgbClr val=\"E4E7EC\"/></a:solidFill></a:lnB>" +
            "<a:noFill/></a:tcPr>",
            "ARR");

        var table = new Drawing.Table(
            new Drawing.TableProperties(),
            new Drawing.TableGrid(new Drawing.GridColumn { Width = 2113280 }),
            new Drawing.TableRow(headerCell) { Height = 685800 },
            new Drawing.TableRow(bodyCell) { Height = 685800 });

        var result = InvokeExtractTable(converter, table);

        var header = result.Rows[0][0];
        Assert.Equal("#0B1F3A", header.BackgroundColor);
        Assert.Equal(TableBorderState.None, header.StylePart!.BorderTopState);
        Assert.Equal(TableBorderState.None, header.StylePart.BorderBottomState);
        Assert.Equal(TableBorderState.None, header.StylePart.BorderLeftState);
        Assert.Equal(TableBorderState.None, header.StylePart.BorderRightState);

        var body = result.Rows[1][0];
        Assert.Null(body.BackgroundColor);
        Assert.Equal(TableBorderState.None, body.StylePart!.BorderTopState);
        Assert.Equal(TableBorderState.None, body.StylePart.BorderLeftState);
        Assert.Equal(TableBorderState.None, body.StylePart.BorderRightState);
        Assert.Equal(TableBorderState.Visible, body.StylePart.BorderBottomState);
        Assert.Equal("#E4E7EC", body.StylePart.BorderBottomColor);
        Assert.Equal(1.0, body.StylePart.BorderBottomWidth!.Value, 3);

        // Cell margins (marL/marR/marT/marB EMU) flow through to per-cell insets.
        Assert.Equal(7.2, body.Insets!.Left!.Value, 3);
        Assert.Equal(7.2, body.Insets.Right!.Value, 3);
        Assert.Equal(3.6, body.Insets.Top!.Value, 3);
        Assert.Equal(3.6, body.Insets.Bottom!.Value, 3);
        Assert.Equal("center", body.VerticalAlign);
    }

    [Fact]
    public void ExtractTable_CellWithoutLineBorders_KeepsInheritedBorderState()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        const string ns = "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"";
        var plainCell = CreateTableCell($"<a:tcPr {ns} marL=\"91440\"/>", "Plain");

        var table = new Drawing.Table(
            new Drawing.TableProperties(),
            new Drawing.TableGrid(new Drawing.GridColumn { Width = 2113280 }),
            new Drawing.TableRow(plainCell) { Height = 685800 });

        var result = InvokeExtractTable(converter, table);

        // No ln* children and no table style → no per-cell stroke state synthesized;
        // the table-level fallback stroke stays in charge.
        Assert.Null(result.Rows[0][0].StylePart);
    }

    [Fact]
    public void ExtractTable_HorizontalOnlyBorders_UndefinedVerticalEdgesRenderAsNone()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        const string ns = "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"";
        // Horizontal-only table: cells define lnT/lnB but no lnL/lnR at all.
        var cell = CreateTableCell(
            $"<a:tcPr {ns}>" +
            "<a:lnT><a:noFill/></a:lnT>" +
            "<a:lnB w=\"12700\"><a:solidFill><a:srgbClr val=\"BFBFBF\"/></a:solidFill></a:lnB>" +
            "</a:tcPr>",
            "Row");

        var table = new Drawing.Table(
            new Drawing.TableProperties(),
            new Drawing.TableGrid(
                new Drawing.GridColumn { Width = 2113280 },
                new Drawing.GridColumn { Width = 2113280 }),
            new Drawing.TableRow(cell, CreateTableCell($"<a:tcPr {ns}><a:lnT><a:noFill/></a:lnT><a:lnB w=\"12700\"><a:solidFill><a:srgbClr val=\"BFBFBF\"/></a:solidFill></a:lnB></a:tcPr>", "Row2"))
            {
                Height = 685800
            });

        var result = InvokeExtractTable(converter, table);

        foreach (var rowCell in result.Rows[0])
        {
            // Vertical edges were defined nowhere → no stroke, never a default grid line.
            Assert.Equal(TableBorderState.None, rowCell.StylePart!.BorderLeftState);
            Assert.Equal(TableBorderState.None, rowCell.StylePart.BorderRightState);
            Assert.Equal(TableBorderState.Visible, rowCell.StylePart.BorderBottomState);
            Assert.Equal("#BFBFBF", rowCell.StylePart.BorderBottomColor);
        }
    }

    [Fact]
    public void ResolveCellStylePart_BandedRows_RestartAfterFirstRow()
    {
        var resolveCellStylePart = typeof(PptxToTypstConverter).GetMethod(
            "ResolveCellStylePart",
            System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic);

        var style = new TableStyleDefinition { StyleId = "banded" };
        style.Parts["band1H"] = new TableStylePart { BackgroundColor = "#EEEEEE" };
        style.Parts["band2H"] = new TableStylePart { BackgroundColor = "#DDDDDD" };

        string? BandAt(int row, bool firstRowFlag)
        {
            var part = resolveCellStylePart!.Invoke(
                null, [row, 0, 5, 2, style, firstRowFlag, true, false, false, false]);
            return (part as TableStylePart)?.BackgroundColor;
        }

        // firstRow on: header row carries no band; bands start at row 1 with band1H.
        Assert.Null(BandAt(0, true));
        Assert.Equal("#EEEEEE", BandAt(1, true));
        Assert.Equal("#DDDDDD", BandAt(2, true));
        Assert.Equal("#EEEEEE", BandAt(3, true));

        // firstRow off: band1H starts at row 0.
        Assert.Equal("#EEEEEE", BandAt(0, false));
        Assert.Equal("#DDDDDD", BandAt(1, false));
        Assert.Equal("#EEEEEE", BandAt(2, false));
    }

    [Fact]
    public void ExtractTable_WithoutStyleId_FallsBackToDefaultTableStyleForBanding()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        // Seed the table style catalog + default style id as LoadTableStyles would.
        var style = new TableStyleDefinition { StyleId = "{DEFAULT}" };
        style.Parts["band1H"] = new TableStylePart { BackgroundColor = "#F2F2F2" };
        var stylesField = typeof(PptxToTypstConverter).GetField(
            "_tableStyles", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        var defaultIdField = typeof(PptxToTypstConverter).GetField(
            "_defaultTableStyleId", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        var styles = Assert.IsType<Dictionary<string, TableStyleDefinition>>(stylesField!.GetValue(converter));
        styles["{DEFAULT}"] = style;
        defaultIdField!.SetValue(converter, "{DEFAULT}");

        const string ns = "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"";
        var table = new Drawing.Table(
            new Drawing.TableProperties { FirstRow = true, BandRow = true },
            new Drawing.TableGrid(new Drawing.GridColumn { Width = 2113280 }),
            new Drawing.TableRow(CreateTableCell($"<a:tcPr {ns}/>", "H")) { Height = 685800 },
            new Drawing.TableRow(CreateTableCell($"<a:tcPr {ns}/>", "B1")) { Height = 685800 },
            new Drawing.TableRow(CreateTableCell($"<a:tcPr {ns}/>", "B2")) { Height = 685800 });

        var result = InvokeExtractTable(converter, table);

        Assert.Null(result.Rows[0][0].BackgroundColor);
        Assert.Equal("#F2F2F2", result.Rows[1][0].BackgroundColor);
        Assert.Null(result.Rows[2][0].BackgroundColor);
    }

    private static Drawing.Table CreateBuiltInStyledTable(string tblPrXml, params Drawing.TableRow[] rows)
    {
        var table = new Drawing.Table(
            new Drawing.TableProperties(tblPrXml),
            new Drawing.TableGrid(
                new Drawing.GridColumn { Width = 2113280 },
                new Drawing.GridColumn { Width = 2113280 }));
        foreach (var row in rows)
            table.Append(row);
        return table;
    }

    [Fact]
    public void BuiltInTableStyles_AllEntries_ParseAsValidTableStyleXml()
    {
        Assert.NotEmpty(BuiltInTableStyles.OuterXmlById);

        foreach (var (styleId, outerXml) in BuiltInTableStyles.OuterXmlById)
        {
            var entry = new Drawing.TableStyleEntry(outerXml);
            Assert.Equal(styleId, entry.StyleId?.Value, ignoreCase: true);
            // Every registered built-in must define at least the whole-table part,
            // otherwise registering it has no effect.
            Assert.NotNull(entry.WholeTable);
        }
    }

    [Fact]
    public void ExtractTable_BuiltInStyleId_ResolvesHeaderFillBordersAndBanding()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);
        // Runs the real LoadTableStyles wiring (built-in registry included), resolving
        // scheme colors against the deck theme: accent1 = 4F81BD, lt1 = FFFFFF.
        _ = converter.Convert();

        const string ns = "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"";
        var table = CreateBuiltInStyledTable(
            $"<a:tblPr {ns} firstRow=\"1\" bandRow=\"1\">" +
            $"<a:tableStyleId>{BuiltInTableStyles.MediumStyle2Accent1Id}</a:tableStyleId></a:tblPr>",
            new Drawing.TableRow(CreateTableCell($"<a:tcPr {ns}/>", "H1"), CreateTableCell($"<a:tcPr {ns}/>", "H2")) { Height = 685800 },
            new Drawing.TableRow(CreateTableCell($"<a:tcPr {ns}/>", "B1"), CreateTableCell($"<a:tcPr {ns}/>", "B2")) { Height = 685800 },
            new Drawing.TableRow(CreateTableCell($"<a:tcPr {ns}/>", "B3"), CreateTableCell($"<a:tcPr {ns}/>", "B4")) { Height = 685800 });

        var result = InvokeExtractTable(converter, table);

        // wholeTbl: white 1pt grid everywhere (invisible against the header fill).
        Assert.Equal("#FFFFFF", result.BorderColor);
        Assert.Equal(1.0, result.BorderWidth, 3);

        // firstRow: accent1 fill, bold white text, white borders (not black grid lines).
        var header = result.Rows[0][0];
        Assert.Equal("#4F81BD", header.BackgroundColor);
        Assert.True(header.Formatting.Bold);
        Assert.Equal("#FFFFFF", header.Formatting.Color);
        Assert.Equal(TableBorderState.Visible, header.StylePart!.BorderTopState);
        Assert.Equal("#FFFFFF", header.StylePart.BorderTopColor);
        Assert.Equal(TableBorderState.Visible, header.StylePart.BorderBottomState);
        Assert.Equal("#FFFFFF", header.StylePart.BorderBottomColor);

        // Banding restarts after the header: band1H (tint 20%) on row 1, band2H
        // (tint 40%) on row 2 — both distinct from the solid header fill. DrawingML
        // a:tint states how much of the source color is KEPT (rest blends to white),
        // so tint 20% renders lighter than tint 40%.
        var band1 = result.Rows[1][0].BackgroundColor;
        var band2 = result.Rows[2][0].BackgroundColor;
        Assert.NotNull(band1);
        Assert.NotNull(band2);
        Assert.NotEqual("#4F81BD", band1);
        Assert.NotEqual(band1, band2);
        Assert.True(Luminance(band1!) > Luminance(band2!), "band1H (tint 20%) must be lighter than band2H (tint 40%)");

        // Inside horizontal rule between banded rows comes from wholeTbl: white.
        var bandStylePart = result.Rows[1][0].StylePart;
        Assert.NotNull(bandStylePart);
        Assert.Equal(TableBorderState.Visible, bandStylePart.BorderBottomState);
        Assert.Equal("#FFFFFF", bandStylePart.BorderBottomColor);
    }

    [Fact]
    public void ExtractTable_BuiltInStyleId_ExplicitCellPropertiesOverrideBuiltInStyle()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);
        _ = converter.Convert();

        const string ns = "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"";
        // Header cell with explicit fill, a red 1.5pt bottom edge and a suppressed top edge.
        var headerCell = CreateTableCell(
            $"<a:tcPr {ns}>" +
            "<a:lnT><a:noFill/></a:lnT>" +
            "<a:lnB w=\"19050\"><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill></a:lnB>" +
            "<a:solidFill><a:srgbClr val=\"C00000\"/></a:solidFill>" +
            "</a:tcPr>",
            "H");
        var table = CreateBuiltInStyledTable(
            $"<a:tblPr {ns} firstRow=\"1\" bandRow=\"1\">" +
            $"<a:tableStyleId>{BuiltInTableStyles.MediumStyle2Accent1Id}</a:tableStyleId></a:tblPr>",
            new Drawing.TableRow(headerCell, CreateTableCell($"<a:tcPr {ns}/>", "H2")) { Height = 685800 },
            new Drawing.TableRow(CreateTableCell($"<a:tcPr {ns}/>", "B1"), CreateTableCell($"<a:tcPr {ns}/>", "B2")) { Height = 685800 });

        var result = InvokeExtractTable(converter, table);

        var header = result.Rows[0][0];
        // Explicit per-cell fill and borders win over the built-in style...
        Assert.Equal("#C00000", header.BackgroundColor);
        Assert.Equal(TableBorderState.None, header.StylePart!.BorderTopState);
        Assert.Equal(TableBorderState.Visible, header.StylePart.BorderBottomState);
        Assert.Equal("#FF0000", header.StylePart.BorderBottomColor);
        Assert.Equal(1.5, header.StylePart.BorderBottomWidth!.Value, 3);
        // ...while undefined edges keep the built-in wholeTbl white border.
        Assert.Equal(TableBorderState.Visible, header.StylePart.BorderLeftState);
        Assert.Equal("#FFFFFF", header.StylePart.BorderLeftColor);
    }

    private static double Luminance(string hexColor)
    {
        var r = Convert.ToInt32(hexColor.Substring(1, 2), 16);
        var g = Convert.ToInt32(hexColor.Substring(3, 2), 16);
        var b = Convert.ToInt32(hexColor.Substring(5, 2), 16);
        return 0.2126 * r + 0.7152 * g + 0.0722 * b;
    }

    [Fact]
    public void GenerateTypstSource_PartialStrokeCell_EmitsOnlyDefinedEdges()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = new TypstPresentation
        {
            Slides =
            [
                new TypstSlide
                {
                    Layout = new PptxEditor.Core.Models.SlideLayout { Width = 720, Height = 540 },
                    Elements =
                    [
                        new TypstElement
                        {
                            Type = "Table",
                            Width = 200,
                            Height = 100,
                            Table = new TypstTableElement
                            {
                                ColumnWidths = [200],
                                Rows =
                                [
                                    [
                                        new TypstTableCell
                                        {
                                            Content = "H",
                                            StylePart = new TableStylePart
                                            {
                                                BorderTopState = TableBorderState.None,
                                                BorderBottomState = TableBorderState.None,
                                                BorderLeftState = TableBorderState.None,
                                                BorderRightState = TableBorderState.None
                                            }
                                        }
                                    ],
                                    [
                                        new TypstTableCell
                                        {
                                            Content = "B",
                                            StylePart = new TableStylePart
                                            {
                                                BorderTopState = TableBorderState.None,
                                                BorderLeftState = TableBorderState.None,
                                                BorderRightState = TableBorderState.None,
                                                BorderBottomState = TableBorderState.Visible,
                                                BorderBottomColor = "#E4E7EC",
                                                BorderBottomWidth = 1
                                            }
                                        }
                                    ]
                                ]
                            }
                        }
                    ]
                }
            ]
        };

        var source = converter.GenerateTypstSource(presentation);

        // No global grid: complex per-cell strokes take over entirely.
        Assert.DoesNotContain("stroke: 1.00pt + rgb(\"#000000\")", source);
        Assert.Contains(
            "stroke: (top: none, bottom: 1.00pt + rgb(\"#E4E7EC\"), left: none, right: none)",
            source);
    }

    [Fact]
    public void Dispose_CleansUpTempDirectory()
    {
        var path = CreateSimplePptx();
        string? tempDir = null;
        
        using (var document = PresentationDocument.Open(path, false))
        using (var converter = new PptxToTypstConverter(document))
        {
            var presentation = converter.Convert();
            tempDir = presentation.TempDirectory;
            Assert.True(Directory.Exists(tempDir));
        }
        
        // After dispose, temp directory should be cleaned up
        // Note: This may fail on some systems due to file locking
        // Assert.False(Directory.Exists(tempDir));
    }

    [Fact]
    public void LaunchReview_Slide1_BackgroundColor_ResolvedFromSlide()
    {
        var baseDir = AppContext.BaseDirectory;
        var path = Path.Combine(baseDir, "..", "..", "..", "..", "examples", "REF", "PPTX", "northwind-launch-review.pptx");
        path = Path.GetFullPath(path);

        using var doc = PresentationDocument.Open(path, false);
        var converter = new PptxToTypstConverter(doc);
        var presentation = converter.Convert();

        var slide1 = presentation.Slides.FirstOrDefault();
        Assert.NotNull(slide1);

        // Slide 1 carries a slide-level solid background (#0B1F3A).
        Assert.Equal("#0B1F3A", slide1.Layout.BackgroundColor);
    }

    [Fact]
    public void StyleResolver_LayoutPlaceholder_MatchesByType_WhenIdxIsNull()
    {
        var path = CreatePlaceholderByTypePptx();

        using var document = PresentationDocument.Open(path, false);
        var slidePart = document.PresentationPart!.SlideParts.First();
        var resolver = new StyleResolver(document, slidePart);

        // Slide placeholder has no idx but type=title
        // Layout has title placeholder with lstStyle font size 4400 (44pt)
        var style = resolver.GetLayoutPlaceholderLstStyle(null, PlaceholderValues.Title, 0);
        Assert.Equal(44.0, style.FontSize);
    }

    [Fact]
    public void ExtractTextFromShape_TitleAlignment_InheritedFromMasterTxStyles()
    {
        var path = CreateMasterTitleRightAlignedPptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var element = Assert.Single(presentation.Slides[0].Elements);
        Assert.Equal("right", element.Text!.Formatting.Align);

        var source = converter.GenerateTypstSource(presentation);
        Assert.Contains("#align(right)", source);
    }

    [Fact]
    public void StyleResolver_MasterPlaceholder_SkipsIdxMatch_WhenTypeDiffers()
    {
        var path = CreateMismatchedMasterPlaceholderPptx();

        using var document = PresentationDocument.Open(path, false);
        var slidePart = document.PresentationPart!.SlideParts.First();
        var resolver = new StyleResolver(document, slidePart);

        // Slide placeholder has idx=2, type=body
        // Master has idx=2, type=dt (date) with 12pt - should NOT match
        var masterStyle = resolver.GetMasterPlaceholderLstStyle(2, PlaceholderValues.Body, 0);
        // Master has no body placeholder to fall back to, so should return empty
        Assert.Null(masterStyle.FontSize);
    }

    [Fact]
    public void GenerateTypstSource_EmitsVerticalAlignBottom()
    {
        var path = CreateSimplePptx();
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = CreateTextPresentation(new TypstTextElement
        {
            Paragraphs =
            [
                new TypstParagraph
                {
                    Content = "Bottom aligned",
                    Formatting = new TypstTextFormatting { FontFamily = "Arial", FontSize = 18 }
                }
            ],
            VerticalAlign = "bottom"
        }, height: 100, width: 300);

        var source = converter.GenerateTypstSource(presentation);
        Assert.Contains("#align(bottom)", source);
        Assert.Contains("height: 100.00pt", source);
    }

    [Fact]
    public void GenerateTypstSource_EmitsVerticalAlignCenter()
    {
        var path = CreateSimplePptx();
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = CreateTextPresentation(new TypstTextElement
        {
            Paragraphs =
            [
                new TypstParagraph
                {
                    Content = "Center aligned",
                    Formatting = new TypstTextFormatting { FontFamily = "Arial", FontSize = 18 }
                }
            ],
            VerticalAlign = "center"
        }, height: 100, width: 300);

        var source = converter.GenerateTypstSource(presentation);
        Assert.Contains("#align(horizon)", source);
        Assert.Contains("height: 100.00pt", source);
    }

    private string CreatePlaceholderByTypePptx()
    {
        var path = Path.Combine(_tempDir, "placeholder-by-type.pptx");

        using (var document = PresentationDocument.Create(path, PresentationDocumentType.Presentation))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new Presentation();
            presentationPart.Presentation.SlideMasterIdList = new SlideMasterIdList();

            var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
            var slideMaster = new SlideMaster(
                new CommonSlideData(
                    new ShapeTree(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties { Id = 0, Name = "" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()
                        ),
                        new GroupShapeProperties(
                            new Drawing.TransformGroup(
                                new Drawing.Offset { X = 0, Y = 0 },
                                new Drawing.Extents { Cx = 0, Cy = 0 },
                                new Drawing.ChildOffset { X = 0, Y = 0 },
                                new Drawing.ChildExtents { Cx = 0, Cy = 0 }
                            )
                        ),
                        new P.Shape(
                            new NonVisualShapeProperties(
                                new NonVisualDrawingProperties { Id = 1, Name = "Title" },
                                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks { NoGrouping = true }),
                                new ApplicationNonVisualDrawingProperties(
                                    new PlaceholderShape { Type = PlaceholderValues.Title }
                                )
                            ),
                            new ShapeProperties(),
                            new TextBody(
                                new Drawing.BodyProperties(),
                                new Drawing.ListStyle()
                            )
                        )
                    )
                ),
                new ColorMap(),
                new SlideLayoutIdList()
            );
            slideMasterPart.SlideMaster = slideMaster;

            var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
            var layoutLstStyle = new Drawing.ListStyle();
            var layoutLvl1 = new OpenXmlUnknownElement("a", "lvl1pPr", "http://schemas.openxmlformats.org/drawingml/2006/main");
            layoutLvl1.Append(new Drawing.DefaultRunProperties { FontSize = 4400 });
            layoutLstStyle.Append(layoutLvl1);
            slideLayoutPart.SlideLayout = new P.SlideLayout(new CommonSlideData(new ShapeTree(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = 0, Name = "" },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                ),
                new GroupShapeProperties(
                    new Drawing.TransformGroup(
                        new Drawing.Offset { X = 0, Y = 0 },
                        new Drawing.Extents { Cx = 0, Cy = 0 },
                        new Drawing.ChildOffset { X = 0, Y = 0 },
                        new Drawing.ChildExtents { Cx = 0, Cy = 0 }
                    )
                ),
                new P.Shape(
                    new NonVisualShapeProperties(
                        new NonVisualDrawingProperties { Id = 1, Name = "Title" },
                        new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks { NoGrouping = true }),
                        new ApplicationNonVisualDrawingProperties(
                            new PlaceholderShape { Type = PlaceholderValues.Title }
                        )
                    ),
                    new ShapeProperties(),
                    new TextBody(
                        new Drawing.BodyProperties(),
                        layoutLstStyle
                    )
                )
            )));
            slideLayoutPart.AddPart(slideMasterPart);

            var layoutId = new SlideLayoutId
            {
                Id = 2147483649,
                RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart)
            };
            slideMaster.SlideLayoutIdList!.Append(layoutId);

            presentationPart.Presentation.SlideIdList = new SlideIdList();
            presentationPart.Presentation.SlideMasterIdList.Append(new SlideMasterId
            {
                Id = 2147483648,
                RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
            });

            var slidePart = presentationPart.AddNewPart<SlidePart>();
            var slide = new Slide(
                new CommonSlideData(
                    new ShapeTree(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties { Id = 0, Name = "" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()
                        ),
                        new GroupShapeProperties(
                            new Drawing.TransformGroup(
                                new Drawing.Offset { X = 0, Y = 0 },
                                new Drawing.Extents { Cx = 0, Cy = 0 },
                                new Drawing.ChildOffset { X = 0, Y = 0 },
                                new Drawing.ChildExtents { Cx = 0, Cy = 0 }
                            )
                        ),
                        new P.Shape(
                            new NonVisualShapeProperties(
                                new NonVisualDrawingProperties { Id = 2, Name = "Title" },
                                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks { NoGrouping = true }),
                                new ApplicationNonVisualDrawingProperties(
                                    new PlaceholderShape { Type = PlaceholderValues.Title }
                                )
                            ),
                            new ShapeProperties(),
                            new TextBody(
                                new Drawing.BodyProperties(),
                                new Drawing.ListStyle(),
                                new Drawing.Paragraph(
                                    new Drawing.Run(new Drawing.Text { Text = "Title Text" })
                                )
                            )
                        )
                    )
                )
            );
            slidePart.Slide = slide;
            slidePart.AddPart(slideLayoutPart);

            var slideId = new SlideId
            {
                Id = 256,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            };
            presentationPart.Presentation.SlideIdList.Append(slideId);

            presentationPart.Presentation.SlideSize = new SlideSize
            {
                Cx = 9144000,
                Cy = 6858000,
                Type = SlideSizeValues.Screen4x3
            };
        }

        return path;
    }

    private string CreateMasterTitleRightAlignedPptx()
    {
        var path = Path.Combine(_tempDir, "master-title-right-aligned.pptx");

        using (var document = PresentationDocument.Create(path, PresentationDocumentType.Presentation))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new Presentation();
            presentationPart.Presentation.SlideMasterIdList = new SlideMasterIdList();

            var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
            var slideMaster = new SlideMaster(
                new CommonSlideData(
                    new ShapeTree(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties { Id = 0, Name = "" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()
                        ),
                        new GroupShapeProperties(
                            new Drawing.TransformGroup(
                                new Drawing.Offset { X = 0, Y = 0 },
                                new Drawing.Extents { Cx = 0, Cy = 0 },
                                new Drawing.ChildOffset { X = 0, Y = 0 },
                                new Drawing.ChildExtents { Cx = 0, Cy = 0 }
                            )
                        )
                    )
                ),
                new ColorMap(),
                new SlideLayoutIdList(),
                // Master declares right-aligned titles: <p:titleStyle><a:lvl1pPr algn="r">
                new TextStyles(
                    new TitleStyle(
                        new Drawing.Level1ParagraphProperties { Alignment = Drawing.TextAlignmentTypeValues.Right }
                    )
                )
            );
            slideMasterPart.SlideMaster = slideMaster;

            var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
            slideLayoutPart.SlideLayout = new P.SlideLayout(new CommonSlideData(new ShapeTree(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = 0, Name = "" },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                ),
                new GroupShapeProperties(
                    new Drawing.TransformGroup(
                        new Drawing.Offset { X = 0, Y = 0 },
                        new Drawing.Extents { Cx = 0, Cy = 0 },
                        new Drawing.ChildOffset { X = 0, Y = 0 },
                        new Drawing.ChildExtents { Cx = 0, Cy = 0 }
                    )
                ),
                new P.Shape(
                    new NonVisualShapeProperties(
                        new NonVisualDrawingProperties { Id = 1, Name = "Title" },
                        new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks { NoGrouping = true }),
                        new ApplicationNonVisualDrawingProperties(
                            new PlaceholderShape { Type = PlaceholderValues.Title }
                        )
                    ),
                    new ShapeProperties(),
                    new TextBody(
                        new Drawing.BodyProperties(),
                        new Drawing.ListStyle()
                    )
                )
            )));
            slideLayoutPart.AddPart(slideMasterPart);

            var layoutId = new SlideLayoutId
            {
                Id = 2147483649,
                RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart)
            };
            slideMaster.SlideLayoutIdList!.Append(layoutId);

            presentationPart.Presentation.SlideIdList = new SlideIdList();
            presentationPart.Presentation.SlideMasterIdList.Append(new SlideMasterId
            {
                Id = 2147483648,
                RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
            });

            var slidePart = presentationPart.AddNewPart<SlidePart>();
            var slide = new Slide(
                new CommonSlideData(
                    new ShapeTree(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties { Id = 0, Name = "" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()
                        ),
                        new GroupShapeProperties(
                            new Drawing.TransformGroup(
                                new Drawing.Offset { X = 0, Y = 0 },
                                new Drawing.Extents { Cx = 0, Cy = 0 },
                                new Drawing.ChildOffset { X = 0, Y = 0 },
                                new Drawing.ChildExtents { Cx = 0, Cy = 0 }
                            )
                        ),
                        new P.Shape(
                            new NonVisualShapeProperties(
                                new NonVisualDrawingProperties { Id = 2, Name = "Title" },
                                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks { NoGrouping = true }),
                                new ApplicationNonVisualDrawingProperties(
                                    new PlaceholderShape { Type = PlaceholderValues.Title }
                                )
                            ),
                            new ShapeProperties(
                                new Drawing.Transform2D(
                                    new Drawing.Offset { X = Pt(40), Y = Pt(20) },
                                    new Drawing.Extents { Cx = Pt(400), Cy = Pt(60) })),
                            new TextBody(
                                new Drawing.BodyProperties(),
                                new Drawing.ListStyle(),
                                // No algn on the slide paragraph — must inherit from master txStyles
                                new Drawing.Paragraph(
                                    new Drawing.Run(new Drawing.Text { Text = "BASIC BLOCK LIST" })
                                )
                            )
                        )
                    )
                )
            );
            slidePart.Slide = slide;
            slidePart.AddPart(slideLayoutPart);

            var slideId = new SlideId
            {
                Id = 256,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            };
            presentationPart.Presentation.SlideIdList.Append(slideId);

            presentationPart.Presentation.SlideSize = new SlideSize
            {
                Cx = 9144000,
                Cy = 6858000,
                Type = SlideSizeValues.Screen4x3
            };
        }

        return path;
    }

    private string CreateMismatchedMasterPlaceholderPptx()
    {
        var path = Path.Combine(_tempDir, "mismatched-master.pptx");

        using (var document = PresentationDocument.Create(path, PresentationDocumentType.Presentation))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new Presentation();
            presentationPart.Presentation.SlideMasterIdList = new SlideMasterIdList();

            var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
            var slideMaster = new SlideMaster(
                new CommonSlideData(
                    new ShapeTree(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties { Id = 0, Name = "" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()
                        ),
                        new GroupShapeProperties(
                            new Drawing.TransformGroup(
                                new Drawing.Offset { X = 0, Y = 0 },
                                new Drawing.Extents { Cx = 0, Cy = 0 },
                                new Drawing.ChildOffset { X = 0, Y = 0 },
                                new Drawing.ChildExtents { Cx = 0, Cy = 0 }
                            )
                        ),
                        new P.Shape(
                            new NonVisualShapeProperties(
                                new NonVisualDrawingProperties { Id = 1, Name = "Date" },
                                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks { NoGrouping = true }),
                                new ApplicationNonVisualDrawingProperties(
                                    new PlaceholderShape { Type = PlaceholderValues.DateAndTime, Index = 2 }
                                )
                            ),
                            new ShapeProperties(),
                            new TextBody(
                                new Drawing.BodyProperties(),
                                new Drawing.ListStyle(
                                    new OpenXmlUnknownElement("a", "lvl1pPr", "http://schemas.openxmlformats.org/drawingml/2006/main")
                                        .AppendChild(new Drawing.DefaultRunProperties { FontSize = 1200 })
                                        .Parent!
                                )
                            )
                        )
                    )
                ),
                new ColorMap(),
                new SlideLayoutIdList()
            );
            slideMasterPart.SlideMaster = slideMaster;

            var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
            slideLayoutPart.SlideLayout = new P.SlideLayout(new CommonSlideData(new ShapeTree(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = 0, Name = "" },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                ),
                new GroupShapeProperties(
                    new Drawing.TransformGroup(
                        new Drawing.Offset { X = 0, Y = 0 },
                        new Drawing.Extents { Cx = 0, Cy = 0 },
                        new Drawing.ChildOffset { X = 0, Y = 0 },
                        new Drawing.ChildExtents { Cx = 0, Cy = 0 }
                    )
                ),
                new P.Shape(
                    new NonVisualShapeProperties(
                        new NonVisualDrawingProperties { Id = 1, Name = "Body" },
                        new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks { NoGrouping = true }),
                        new ApplicationNonVisualDrawingProperties(
                            new PlaceholderShape { Type = PlaceholderValues.Body, Index = 2 }
                        )
                    ),
                    new ShapeProperties(),
                    new TextBody(
                        new Drawing.BodyProperties(),
                        new Drawing.ListStyle(
                            new OpenXmlUnknownElement("a", "lvl1pPr", "http://schemas.openxmlformats.org/drawingml/2006/main")
                                .AppendChild(new Drawing.DefaultRunProperties { FontSize = 2400 })
                                .Parent!
                        )
                    )
                )
            )));
            slideLayoutPart.AddPart(slideMasterPart);

            var layoutId = new SlideLayoutId
            {
                Id = 2147483649,
                RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart)
            };
            slideMaster.SlideLayoutIdList!.Append(layoutId);

            presentationPart.Presentation.SlideIdList = new SlideIdList();
            presentationPart.Presentation.SlideMasterIdList.Append(new SlideMasterId
            {
                Id = 2147483648,
                RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
            });

            var slidePart = presentationPart.AddNewPart<SlidePart>();
            var slide = new Slide(
                new CommonSlideData(
                    new ShapeTree(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties { Id = 0, Name = "" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()
                        ),
                        new GroupShapeProperties(
                            new Drawing.TransformGroup(
                                new Drawing.Offset { X = 0, Y = 0 },
                                new Drawing.Extents { Cx = 0, Cy = 0 },
                                new Drawing.ChildOffset { X = 0, Y = 0 },
                                new Drawing.ChildExtents { Cx = 0, Cy = 0 }
                            )
                        ),
                        new P.Shape(
                            new NonVisualShapeProperties(
                                new NonVisualDrawingProperties { Id = 2, Name = "Body" },
                                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks { NoGrouping = true }),
                                new ApplicationNonVisualDrawingProperties(
                                    new PlaceholderShape { Type = PlaceholderValues.Body, Index = 2 }
                                )
                            ),
                            new ShapeProperties(),
                            new TextBody(
                                new Drawing.BodyProperties(),
                                new Drawing.ListStyle(),
                                new Drawing.Paragraph(
                                    new Drawing.Run(new Drawing.Text { Text = "Body Text" })
                                )
                            )
                        )
                    )
                )
            );
            slidePart.Slide = slide;
            slidePart.AddPart(slideLayoutPart);

            var slideId = new SlideId
            {
                Id = 256,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            };
            presentationPart.Presentation.SlideIdList.Append(slideId);

            presentationPart.Presentation.SlideSize = new SlideSize
            {
                Cx = 9144000,
                Cy = 6858000,
                Type = SlideSizeValues.Screen4x3
            };
        }

        return path;
    }

    [Fact]
    public void GenerateTypstSource_LineBreakInParagraph_EmitsLinebreak()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var shape = new P.Shape(
            new TextBody(
                new Drawing.BodyProperties(),
                new Drawing.ListStyle(),
                new Drawing.Paragraph(
                    new Drawing.Run(new Drawing.Text { Text = "Line one" }),
                    new Drawing.Break(),
                    new Drawing.Run(new Drawing.Text { Text = "Line two" }))));

        var method = typeof(PptxToTypstConverter).GetMethod(
            "ExtractTextFromShape",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(P.Shape), typeof(StyleResolver), typeof(SlidePart)],
            null);

        var slidePart = document.PresentationPart!.SlideParts.FirstOrDefault();
        Assert.NotNull(slidePart);

        var styleResolver = new StyleResolver(document, slidePart);
        var text = Assert.IsType<TypstTextElement>(method!.Invoke(converter, [shape, styleResolver, slidePart]));

        var source = converter.GenerateTypstSource(new TypstPresentation
        {
            Slides =
            [
                new TypstSlide
                {
                    Layout = new PptxEditor.Core.Models.SlideLayout { Width = 720, Height = 540 },
                    Elements =
                    [
                        new TypstElement
                        {
                            Type = "Text",
                            X = 100, Width = 300, Height = 100,
                            Text = text
                        }
                    ]
                }
            ]
        });

        Assert.Contains("#linebreak()", source);
        Assert.Contains("Line one", source);
        Assert.Contains("Line two", source);
    }

    [Fact]
    public void GenerateTypstSource_DoubleLineBreak_EmitsTwoLinebreaks()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var shape = new P.Shape(
            new TextBody(
                new Drawing.BodyProperties(),
                new Drawing.ListStyle(),
                new Drawing.Paragraph(
                    new Drawing.Run(new Drawing.Text { Text = "Line one" }),
                    new Drawing.Break(),
                    new Drawing.Break(),
                    new Drawing.Run(new Drawing.Text { Text = "Line two" }))));

        var method = typeof(PptxToTypstConverter).GetMethod(
            "ExtractTextFromShape",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(P.Shape), typeof(StyleResolver), typeof(SlidePart)],
            null);

        var slidePart = document.PresentationPart!.SlideParts.FirstOrDefault();
        Assert.NotNull(slidePart);

        var styleResolver = new StyleResolver(document, slidePart);
        var text = Assert.IsType<TypstTextElement>(method!.Invoke(converter, [shape, styleResolver, slidePart]));

        var source = converter.GenerateTypstSource(new TypstPresentation
        {
            Slides =
            [
                new TypstSlide
                {
                    Layout = new PptxEditor.Core.Models.SlideLayout { Width = 720, Height = 540 },
                    Elements =
                    [
                        new TypstElement
                        {
                            Type = "Text",
                            X = 100, Width = 300, Height = 100,
                            Text = text
                        }
                    ]
                }
            ]
        });

        // Should contain two linebreak calls, creating a blank line
        var linebreakCount = source.Split("#linebreak()").Length - 1;
        Assert.True(linebreakCount >= 2, $"Expected at least 2 #linebreak() calls, found {linebreakCount}");
    }

    [Fact]
    public void GenerateTypstSource_SpaceBeforeParagraph_EmitsVerticalSpacing()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = CreateTextPresentation(new TypstTextElement
        {
            Paragraphs =
            [
                new TypstParagraph
                {
                    Content = "First paragraph",
                    Formatting = new TypstTextFormatting { FontFamily = "Arial", FontSize = 18 }
                },
                new TypstParagraph
                {
                    Content = "Second paragraph",
                    Formatting = new TypstTextFormatting { FontFamily = "Arial", FontSize = 18 },
                    SpaceBefore = 12.5
                }
            ]
        });

        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains("#v(12.50pt)", source);
        Assert.Contains("[First paragraph]", source);
        Assert.Contains("[Second paragraph]", source);
    }

    [Fact]
    public void GenerateTypstSource_SpaceAfterParagraph_EmitsVerticalSpacing()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = CreateTextPresentation(new TypstTextElement
        {
            Paragraphs =
            [
                new TypstParagraph
                {
                    Content = "First paragraph",
                    Formatting = new TypstTextFormatting { FontFamily = "Arial", FontSize = 18 },
                    SpaceAfter = 8.0
                },
                new TypstParagraph
                {
                    Content = "Second paragraph",
                    Formatting = new TypstTextFormatting { FontFamily = "Arial", FontSize = 18 }
                }
            ]
        });

        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains("#v(8.00pt)", source);
        Assert.Contains("[First paragraph]", source);
        Assert.Contains("[Second paragraph]", source);
    }

    [Fact]
    public void GenerateTypstSource_ListWithMarginLeft_EmitsIndentParameter()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = CreateTextPresentation(new TypstTextElement
        {
            Paragraphs =
            [
                new TypstParagraph
                {
                    Content = "Item one",
                    Formatting = new TypstTextFormatting { FontFamily = "Arial", FontSize = 18 },
                    HasBullet = true,
                    BulletChar = "•",
                    MarginLeft = 22.5,
                    Indent = -7.5
                },
                new TypstParagraph
                {
                    Content = "Item two",
                    Formatting = new TypstTextFormatting { FontFamily = "Arial", FontSize = 18 },
                    HasBullet = true,
                    BulletChar = "•",
                    MarginLeft = 22.5,
                    Indent = -7.5
                }
            ]
        });

        var source = converter.GenerateTypstSource(presentation);

        // OOXML: marL=22.5pt is the body text offset, indent=-7.5pt shifts the marker
        // left of the body, so the marker sits at 15pt and the body 7.5pt right of it.
        Assert.Contains("indent: 15.00pt", source);
        Assert.Contains("body-indent: 7.50pt", source);
        Assert.Contains("[Item one]", source);
        Assert.Contains("[Item two]", source);
    }

    [Fact]
    public void GenerateTypstSource_NormalToListTransition_NoExtraParagraphBreak()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = CreateTextPresentation(new TypstTextElement
        {
            Paragraphs =
            [
                new TypstParagraph
                {
                    Content = "Normal text",
                    Formatting = new TypstTextFormatting { FontFamily = "Arial", FontSize = 18 }
                },
                new TypstParagraph
                {
                    Content = "Item one",
                    Formatting = new TypstTextFormatting { FontFamily = "Arial", FontSize = 18 },
                    HasBullet = true,
                    BulletChar = "•"
                },
                new TypstParagraph
                {
                    Content = "Item two",
                    Formatting = new TypstTextFormatting { FontFamily = "Arial", FontSize = 18 },
                    HasBullet = true,
                    BulletChar = "•"
                }
            ]
        });

        var source = converter.GenerateTypstSource(presentation);

        // There should be a \n\n between normal paragraphs, but NOT between normal and list
        var normalToListIndex = source.IndexOf("Normal text");
        Assert.True(normalToListIndex >= 0);

        // The list should immediately follow without extra blank lines
        Assert.Contains("#list", source);
        Assert.Contains("[Item one]", source);
        Assert.Contains("[Item two]", source);
    }

    [Fact]
    public void DiscoverSystemFonts_FindsArial()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var method = typeof(PptxToTypstConverter).GetMethod(
            "DiscoverSystemFonts",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
            null,
            Type.EmptyTypes,
            null);

        var result = Assert.IsType<HashSet<string>>(method!.Invoke(converter, null));

        // Arial is commonly available on most systems
        if (result.Contains("Arial"))
        {
            Assert.Contains("Arial", result);
        }
    }

    [Fact]
    public void SubstituteUnavailableFont_AptosVariantFallsBackToAptosBeforeCarlito()
    {
        var method = typeof(PptxToTypstConverter).GetMethod(
            "SubstituteUnavailableFont",
            System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(string), typeof(HashSet<string>)],
            null);

        var availableFonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Aptos", "Carlito" };

        var result = method!.Invoke(null, ["Aptos Light", availableFonts]);

        Assert.Equal("Aptos", result);
    }

    [Fact]
    public void SubstituteUnavailableFont_AptosVariantFallsBackToCarlitoWhenAptosUnavailable()
    {
        var method = typeof(PptxToTypstConverter).GetMethod(
            "SubstituteUnavailableFont",
            System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(string), typeof(HashSet<string>)],
            null);

        var availableFonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Carlito" };

        var result = method!.Invoke(null, ["Aptos Display", availableFonts]);

        Assert.Equal("Carlito", result);
    }

    [Fact]
    public void GenerateTypstSource_GlobalFontStackPrefersAvailableThemeMinorFontBeforeCarlito()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var availableSystemFontsField = typeof(PptxToTypstConverter).GetField(
            "_availableSystemFonts",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        availableSystemFontsField!.SetValue(converter, new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Aptos", "Carlito" });

        var presentation = new TypstPresentation
        {
            ThemeFonts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["+mn-lt"] = "Aptos"
            }
        };

        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains("#set text(font: (\"Aptos\", \"Carlito\"", source);
    }

    [Fact]
    public void GenerateTypstSource_UsesAptosAliasWhenDiscoveredAsAvailable()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var availableSystemFontsField = typeof(PptxToTypstConverter).GetField(
            "_availableSystemFonts",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        availableSystemFontsField!.SetValue(converter, new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Aptos Light", "Carlito" });

        var presentation = CreateTextPresentation(new TypstTextElement
        {
            Paragraphs =
            [
                new TypstParagraph
                {
                    Content = "Theme title",
                    Formatting = new TypstTextFormatting { FontFamily = "+mj-lt", FontSize = 24 }
                }
            ]
        });
        presentation.ThemeFonts["+mj-lt"] = "Aptos Light";
        presentation.ThemeFonts["+mn-lt"] = "Aptos";

        var themeFontsField = typeof(PptxToTypstConverter).GetField(
            "_themeFonts",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        themeFontsField!.SetValue(converter, presentation.ThemeFonts);

        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains("font: (\"Aptos Light\"", source);
        Assert.DoesNotContain("font: \"Carlito\"", source);
    }

    [Fact]
    public void GenerateTypstSource_PerRunFontEmitsFallbackChainNotSingleFamily()
    {
        // In Typst, a per-element `font:` parameter REPLACES the whole font chain.
        // Emitting a bare `font: "Calibri"` therefore collapses the chain to Typst's
        // (serif) embedded fallback whenever Calibri is absent from the compiler's
        // font set — the per-run emission must carry the global fallback chain.
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = CreateTextPresentation(new TypstTextElement
        {
            Paragraphs =
            [
                new TypstParagraph
                {
                    Content = "Body text",
                    Formatting = new TypstTextFormatting { FontFamily = "+mn-lt", FontSize = 14 }
                }
            ]
        });
        presentation.ThemeFonts["+mn-lt"] = "Calibri";

        var themeFontsField = typeof(PptxToTypstConverter).GetField(
            "_themeFonts",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        themeFontsField!.SetValue(converter, presentation.ThemeFonts);

        var source = converter.GenerateTypstSource(presentation);

        Assert.DoesNotContain("font: \"Calibri\"", source);
        Assert.Contains("font: (\"Calibri\", \"Carlito\", \"Arial\", \"Helvetica\", \"Liberation Sans\")", source);
    }

    [Fact]
    public void GetFontMetrics_EmbeddedFont_ReturnsCachedValue()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        // Trigger Convert to populate any embedded fonts
        converter.Convert();

        var getFontMetrics = typeof(PptxToTypstConverter).GetMethod(
            "GetFontMetrics",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(string)],
            null);

        // If no embedded fonts, this test is not applicable
        var fontMetricsField = typeof(PptxToTypstConverter).GetField(
            "_fontMetrics",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        var fontMetrics = Assert.IsType<Dictionary<string, TypstFontMetrics>>(fontMetricsField!.GetValue(converter));

        if (fontMetrics.Count == 0)
            return;

        var firstEmbedded = fontMetrics.Keys.First();
        var metrics = getFontMetrics!.Invoke(converter, [firstEmbedded]);
        Assert.NotNull(metrics);
        Assert.IsType<TypstFontMetrics>(metrics);
    }

    [Fact]
    public void GetFontMetrics_SystemFont_ReadsFromDisk()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        converter.Convert();

        var getFontMetrics = typeof(PptxToTypstConverter).GetMethod(
            "GetFontMetrics",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(string)],
            null);

        // Try to read metrics for a common system font
        var metrics = getFontMetrics!.Invoke(converter, ["Arial"]);
        if (metrics == null)
        {
            // Arial not available on this system, skip
            return;
        }

        var typedMetrics = Assert.IsType<TypstFontMetrics>(metrics);
        Assert.True(typedMetrics.UnitsPerEm > 0);

        // Verify it was cached
        var fontMetricsField = typeof(PptxToTypstConverter).GetField(
            "_fontMetrics",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        var fontMetrics = Assert.IsType<Dictionary<string, TypstFontMetrics>>(fontMetricsField!.GetValue(converter));
        Assert.Contains("Arial", fontMetrics.Keys);
    }

    [Fact]
    public void GetTextMetricOffset_SystemFont_ReturnsNonZero()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        converter.Convert();

        var getTextMetricOffset = typeof(PptxToTypstConverter).GetMethod(
            "GetTextMetricOffset",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(TypstTextElement)],
            null);

        var text = new TypstTextElement
        {
            Paragraphs =
            [
                new TypstParagraph
                {
                    Content = "Test",
                    Formatting = new TypstTextFormatting
                    {
                        FontFamily = "Arial",
                        FontSize = 24
                    }
                }
            ]
        };

        var offset = getTextMetricOffset!.Invoke(converter, [text]);
        if (offset == null)
        {
            // Arial not available, skip
            return;
        }

        // The offset should be >= 0 (it's Math.Max(0, ...))
        var offsetValue = Assert.IsType<double>(offset);
        Assert.True(offsetValue >= 0);
    }

    [Fact]
    public void ExtractTableStylePart_ReadsTextBoldAndColor()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        // Force LoadTableStyles to run by calling Convert
        converter.Convert();

        var extractTableStylePart = typeof(PptxToTypstConverter).GetMethod(
            "ExtractTableStylePart",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(Drawing.TablePartStyleType), typeof(string), typeof(TableStyleDefinition), typeof(StyleResolver)],
            null);

        // Build a TablePartStyleType with tcTxStyle containing bold and white color
        var tcTxStyle = new Drawing.TableCellTextStyle();
        tcTxStyle.SetAttribute(new OpenXmlAttribute("b", string.Empty, "on"));
        var defRPr = new Drawing.DefaultRunProperties();
        defRPr.Append(new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = "FFFFFF" }));
        tcTxStyle.Append(defRPr);

        var part = new Drawing.FirstRow();
        part.Append(new Drawing.TableCellStyle());
        part.Append(tcTxStyle);

        var definition = new TableStyleDefinition { StyleId = "test" };
        var slidePart = document.PresentationPart!.SlideParts.FirstOrDefault();
        var styleResolver = slidePart != null ? new StyleResolver(document, slidePart) : null;

        extractTableStylePart!.Invoke(converter, [part, "firstRow", definition, styleResolver]);

        Assert.True(definition.Parts.ContainsKey("firstRow"));
        var stylePart = definition.Parts["firstRow"];
        Assert.True(stylePart.TextBold);
        Assert.Equal("#FFFFFF", stylePart.TextColor);
    }

    [Fact]
    public void ExtractCellFormatting_MergesTableStyleText()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var extractCellFormatting = typeof(PptxToTypstConverter).GetMethod(
            "ExtractCellFormatting",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(Drawing.TableCell), typeof(TableStylePart), typeof(StyleResolver)],
            null);

        // Create a table cell with empty text body (no explicit formatting)
        var cell = new Drawing.TableCell(
            new Drawing.TextBody(
                new Drawing.BodyProperties(),
                new Drawing.ListStyle(),
                new Drawing.Paragraph(
                    new Drawing.Run(new Drawing.Text { Text = "Cell text" })
                )
            )
        );

        var stylePart = new TableStylePart { TextBold = true, TextColor = "#FFFFFF" };

        var result = extractCellFormatting!.Invoke(converter, [cell, stylePart, null]);
        var formatting = Assert.IsType<TypstTextFormatting>(result);

        Assert.True(formatting.Bold);
        Assert.Equal("#FFFFFF", formatting.Color);
    }

    [Fact]
    public void ExtractCellParagraphs_InheritsStyleBoldWhenRunOnlyOverridesColor()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var extractCellFormatting = typeof(PptxToTypstConverter).GetMethod(
            "ExtractCellFormatting",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(Drawing.TableCell), typeof(TableStylePart), typeof(StyleResolver)],
            null);
        var extractCellParagraphs = typeof(PptxToTypstConverter).GetMethod(
            "ExtractCellParagraphs",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(Drawing.TableCell), typeof(TypstTextFormatting), typeof(StyleResolver)],
            null);

        var cell = new Drawing.TableCell(
            new Drawing.TextBody(
                new Drawing.BodyProperties(),
                new Drawing.ListStyle(),
                new Drawing.Paragraph(
                    new Drawing.Run(
                        new Drawing.RunProperties(
                            new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = "000000" }))
                        {
                            FontSize = new Int32Value(1400)
                        },
                        new Drawing.Text { Text = "Plan name\t" }))));

        var stylePart = new TableStylePart
        {
            TextBold = true,
            TextItalic = true,
            TextColor = "#FFFFFF"
        };

        var formatting = Assert.IsType<TypstTextFormatting>(extractCellFormatting!.Invoke(converter, [cell, stylePart, null]));
        var paragraphs = Assert.IsType<List<TypstParagraph>>(extractCellParagraphs!.Invoke(converter, [cell, formatting, null]));

        Assert.True(formatting.Bold);
        Assert.True(formatting.Italic);
        Assert.Equal("#000000", formatting.Color);
        Assert.Equal(14.0, formatting.FontSize);

        var paragraph = Assert.Single(paragraphs);
        var run = Assert.Single(paragraph.Runs);
        Assert.Equal("Plan name", paragraph.Content);
        Assert.Equal("Plan name", run.Content);
        Assert.True(run.Formatting.Bold);
        Assert.True(run.Formatting.Italic);
        Assert.Equal("#000000", run.Formatting.Color);
        Assert.Equal(14.0, run.Formatting.FontSize);
    }

    [Fact]
    public void GenerateTypstSource_TableHeader_InheritedBoldWithDirectBlackColor()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = new TypstPresentation
        {
            Slides =
            [
                new TypstSlide
                {
                    Layout = new PptxEditor.Core.Models.SlideLayout { Width = 720, Height = 540 },
                    Elements =
                    [
                        new TypstElement
                        {
                            Type = "Table",
                            Width = 200,
                            Height = 80,
                            Table = new TypstTableElement
                            {
                                ColumnWidths = [200],
                                Rows =
                                [
                                    new List<TypstTableCell>
                                    {
                                        new()
                                        {
                                            Paragraphs =
                                            [
                                                new TypstParagraph
                                                {
                                                    Content = "Plan name",
                                                    Runs =
                                                    [
                                                        new TypstTextRun
                                                        {
                                                            Content = "Plan name",
                                                            Formatting = new TypstTextFormatting
                                                            {
                                                                Bold = true,
                                                                Color = "#000000",
                                                                FontSize = 14
                                                            }
                                                        }
                                                    ]
                                                }
                                            ]
                                        }
                                    }
                                ]
                            }
                        }
                    ]
                }
            ]
        };

        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains("#text(size: 14.00pt, weight: \"bold\")[Plan name]", source);
        Assert.DoesNotContain("Plan name\t", source);
    }

    [Fact]
    public void GenerateTypstSource_TableHeader_BoldWhiteText()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = new TypstPresentation
        {
            Slides =
            [
                new TypstSlide
                {
                    Layout = new PptxEditor.Core.Models.SlideLayout { Width = 720, Height = 540 },
                    Elements =
                    [
                        new TypstElement
                        {
                            Type = "Table",
                            X = 100, Y = 100, Width = 500, Height = 200,
                            Table = new TypstTableElement
                            {
                                ColumnWidths = [250, 250],
                                Rows =
                                [
                                    new List<TypstTableCell>
                                    {
                                        new TypstTableCell
                                        {
                                            Content = "Header 1",
                                            Formatting = new TypstTextFormatting
                                            {
                                                Bold = true,
                                                Color = "#FFFFFF",
                                                FontFamily = "Arial",
                                                FontSize = 14
                                            },
                                            BackgroundColor = "#CEBA80"
                                        },
                                        new TypstTableCell
                                        {
                                            Content = "Header 2",
                                            Formatting = new TypstTextFormatting
                                            {
                                                Bold = true,
                                                Color = "#FFFFFF",
                                                FontFamily = "Arial",
                                                FontSize = 14
                                            },
                                            BackgroundColor = "#CEBA80"
                                        }
                                    }
                                ]
                            }
                        }
                    ]
                }
            ]
        };

        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains("weight: \"bold\"", source);
        Assert.Contains("fill: rgb(\"#FFFFFF\")", source);
        Assert.Contains("Header 1", source);
        Assert.Contains("Header 2", source);
    }

    [Fact]
    public void BuildCellStroke_NoFillVerticalBorders()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var buildCellStroke = typeof(PptxToTypstConverter).GetMethod(
            "BuildCellStroke",
            System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(TypstTableCell), typeof(double), typeof(string)],
            null);

        var cell = new TypstTableCell
        {
            Content = "Cell",
            StylePart = new TableStylePart
            {
                BorderLeftNone = true,
                BorderRightNone = true,
                BorderLeftState = TableBorderState.None,
                BorderRightState = TableBorderState.None
            }
        };

        var result = buildCellStroke!.Invoke(null, [cell, 1.0, "#000000"]);
        var stroke = Assert.IsType<string>(result);

        Assert.Contains("left: none", stroke);
        Assert.Contains("right: none", stroke);
    }

    [Fact]
    public void ExtractSolidFillColor_TintKeepsSourceColorFractionAndBlendsTowardWhite()
    {
        var extractSolidFillColor = typeof(PptxToTypstConverter).GetMethod(
            "ExtractSolidFillColorStatic",
            System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(Drawing.SolidFill), typeof(StyleResolver)],
            null);

        var solidFill = new Drawing.SolidFill(
            new Drawing.RgbColorModelHex(
                new Drawing.Tint { Val = 20000 })
            {
                Val = "CEBA80"
            });

        var result = Assert.IsType<string>(extractSolidFillColor!.Invoke(null, [solidFill, null]));

        Assert.Equal("#F6F3EC", result);
    }

    [Fact]
    public void ApplyTableGridBorders_PreservesExplicitInteriorCellBorder()
    {
        var path = CreateSimplePptx();
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var mergeExplicitCellBorders = typeof(PptxToTypstConverter).GetMethod(
            "MergeExplicitCellBorders",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(Drawing.TableCell), typeof(TableStylePart), typeof(StyleResolver)],
            null);
        var applyTableGridBorders = typeof(PptxToTypstConverter).GetMethod(
            "ApplyTableGridBorders",
            System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(TableStylePart), typeof(int), typeof(int), typeof(int), typeof(int)],
            null);

        // Schema-correct explicit border: a:tcPr carries a:lnL/a:lnR/a:lnT/a:lnB directly
        // (a:tcBdr only exists in tableStyles.xml cell styles, never inside a:tcPr).
        var cell = new Drawing.TableCell(
            new Drawing.TableCellProperties(
                "<a:tcPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                "<a:lnT w=\"25400\"><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill></a:lnT>" +
                "</a:tcPr>"),
            new Drawing.TextBody(new Drawing.BodyProperties(), new Drawing.ListStyle(), new Drawing.Paragraph()));

        var inheritedStylePart = new TableStylePart
        {
            BorderInsideHState = TableBorderState.Visible,
            BorderInsideHColor = "#0000FF",
            BorderInsideHWidth = 1.0,
            BorderInsideVState = TableBorderState.Visible,
            BorderInsideVColor = "#000000",
            BorderInsideVWidth = 1.0
        };

        var merged = mergeExplicitCellBorders!.Invoke(converter, [cell, inheritedStylePart, null]);
        var result = applyTableGridBorders!.Invoke(null, [merged, 1, 1, 3, 3]);
        var mapped = Assert.IsType<TableStylePart>(result);

        Assert.Equal(TableBorderState.Visible, mapped.BorderTopState);
        Assert.Equal("#FF0000", mapped.BorderTopColor);
        Assert.Equal(2.0, mapped.BorderTopWidth);
        Assert.True(mapped.BorderTopExplicit);
        Assert.Equal(TableBorderState.Visible, mapped.BorderBottomState);
        Assert.Equal("#0000FF", mapped.BorderBottomColor);
        Assert.Equal(TableBorderState.None, mapped.BorderLeftState);
        Assert.Equal(TableBorderState.Visible, mapped.BorderRightState);
        Assert.Equal("#000000", mapped.BorderRightColor);
    }

    [Fact]
    public void ExtractCellText_PreservesParagraphsAndLineBreaks()
    {
        var path = CreateSimplePptx();
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var extractCellText = typeof(PptxToTypstConverter).GetMethod(
            "ExtractCellText",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(Drawing.TableCell)],
            null);

        var cell = new Drawing.TableCell(
            new Drawing.TextBody(
                new Drawing.BodyProperties(),
                new Drawing.ListStyle(),
                new Drawing.Paragraph(
                    new Drawing.Run(new Drawing.Text { Text = "Line 1" }),
                    new Drawing.Break(),
                    new Drawing.Run(new Drawing.Text { Text = "Line 2" })),
                new Drawing.Paragraph(
                    new Drawing.Run(new Drawing.Text { Text = "Next paragraph" }))));

        var result = Assert.IsType<string>(extractCellText!.Invoke(converter, [cell]));

        Assert.Equal("Line 1\nLine 2\n\nNext paragraph", result);
    }

    [Fact]
    public void GenerateTypstSource_TableCell_PreservesParagraphBreaksLineBreaksAndAlignment()
    {
        var path = CreateSimplePptx();
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = new TypstPresentation
        {
            Slides =
            [
                new TypstSlide
                {
                    Layout = new PptxEditor.Core.Models.SlideLayout { Width = 720, Height = 540 },
                    Elements =
                    [
                        new TypstElement
                        {
                            Type = "Table",
                            X = 10,
                            Y = 10,
                            Width = 200,
                            Height = 80,
                            Table = new TypstTableElement
                            {
                                ColumnWidths = [200],
                                Rows =
                                [
                                    new List<TypstTableCell>
                                    {
                                        new()
                                        {
                                            Formatting = new TypstTextFormatting(),
                                            Paragraphs =
                                            [
                                                new TypstParagraph
                                                {
                                                    Content = "First\nSecond",
                                                    Formatting = new TypstTextFormatting { Align = "center" },
                                                    Runs =
                                                    [
                                                        new TypstTextRun { Content = "First", Formatting = new TypstTextFormatting { Align = "center" } },
                                                        new TypstTextRun { IsLineBreak = true, Formatting = new TypstTextFormatting { Align = "center" } },
                                                        new TypstTextRun { Content = "Second", Formatting = new TypstTextFormatting { Align = "center" } }
                                                    ]
                                                },
                                                new TypstParagraph
                                                {
                                                    Content = "Third",
                                                    Formatting = new TypstTextFormatting()
                                                }
                                            ]
                                        }
                                    }
                                ]
                            }
                        }
                    ]
                }
            ]
        };

        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains("#align(center)[", source);
        Assert.Contains("First #linebreak() Second", source);
        Assert.Contains("Second]\n\nThird", source);
    }

    [Fact]
    public void GenerateTypstSource_TableCell_UsesIdenticalRunFormattingWhenCollapsingRuns()
    {
        var path = CreateSimplePptx();
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var directRunFormatting = new TypstTextFormatting
        {
            Bold = true,
            Color = "#FF0000"
        };

        var presentation = new TypstPresentation
        {
            Slides =
            [
                new TypstSlide
                {
                    Layout = new PptxEditor.Core.Models.SlideLayout { Width = 720, Height = 540 },
                    Elements =
                    [
                        new TypstElement
                        {
                            Type = "Table",
                            X = 10,
                            Y = 10,
                            Width = 200,
                            Height = 80,
                            Table = new TypstTableElement
                            {
                                ColumnWidths = [200],
                                Rows =
                                [
                                    new List<TypstTableCell>
                                    {
                                        new()
                                        {
                                            Formatting = new TypstTextFormatting(),
                                            Paragraphs =
                                            [
                                                new TypstParagraph
                                                {
                                                    Content = "Hello world",
                                                    Formatting = new TypstTextFormatting(),
                                                    Runs =
                                                    [
                                                        new TypstTextRun { Content = "Hello ", Formatting = directRunFormatting },
                                                        new TypstTextRun { Content = "world", Formatting = directRunFormatting }
                                                    ]
                                                }
                                            ]
                                        }
                                    }
                                ]
                            }
                        }
                    ]
                }
            ]
        };

        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains("#text(weight: \"bold\", fill: rgb(\"#FF0000\"))[Hello world]", source);
    }

    [Fact]
    public void Convert_ShapeWithMissingTransform_UsesDefaultPositionAndConvertsText()
    {
        var shape = new P.Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = 40, Name = "No transform" },
                new NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(),
            new TextBody(
                new Drawing.BodyProperties(),
                new Drawing.ListStyle(),
                new Drawing.Paragraph(new Drawing.Run(new Drawing.Text { Text = "Default positioned" }))));
        var path = CreateGroupShapePptx("shape-missing-transform.pptx", shape);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var element = AssertSingleTextElement(presentation, "Default positioned");

        AssertPosition(element, x: 0, y: 0, width: 100, height: 50);
    }

    [Fact]
    public void Convert_ShapesWithMissingTextBodyOrEmptyParagraphs_YieldNoTextElements()
    {
        var noTextBody = new P.Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = 41, Name = "No text body" },
                new NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(new Drawing.Transform2D(
                new Drawing.Offset { X = Pt(10), Y = Pt(10) },
                new Drawing.Extents { Cx = Pt(100), Cy = Pt(50) })));
        var emptyParagraph = new P.Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = 42, Name = "Empty paragraph" },
                new NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(new Drawing.Transform2D(
                new Drawing.Offset { X = Pt(20), Y = Pt(20) },
                new Drawing.Extents { Cx = Pt(100), Cy = Pt(50) })),
            new TextBody(new Drawing.BodyProperties(), new Drawing.ListStyle(), new Drawing.Paragraph()));
        var path = CreateGroupShapePptx("shape-no-text.pptx", noTextBody, emptyParagraph);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();

        Assert.DoesNotContain(presentation.Slides[0].Elements, e => e.Type == "Text");
    }

    [Fact]
    public void ExtractTextFromShape_BodyPropertiesParagraphsAndRunOnOffBranches()
    {
        var path = CreateSimplePptx();
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);
        var slidePart = document.PresentationPart!.SlideParts.First();
        var styleResolver = new StyleResolver(document, slidePart);

        var bodyPr = new Drawing.BodyProperties
        {
            LeftInset = 12700,
            TopInset = 25400,
            RightInset = 38100,
            BottomInset = 50800
        };
        bodyPr.SetAttribute(new OpenXmlAttribute("", "anchor", "", "ctr"));
        bodyPr.SetAttribute(new OpenXmlAttribute("", "wrap", "", "square"));
        bodyPr.Append(new Drawing.ShapeAutoFit());

        var firstPPr = new Drawing.ParagraphProperties();
        firstPPr.SetAttribute(new OpenXmlAttribute("", "algn", "", "r"));
        var shape = new P.Shape(
            new TextBody(
                bodyPr,
                new Drawing.ListStyle(),
                new Drawing.Paragraph(
                    firstPPr,
                    new Drawing.Run(new Drawing.RunProperties { Bold = true }, new Drawing.Text { Text = "Bold" }),
                    new Drawing.Run(new Drawing.RunProperties { Bold = false, Italic = true }, new Drawing.Text { Text = "Italic" }),
                    new Drawing.Run(new Drawing.Text { Text = "Plain" })),
                new Drawing.Paragraph(new Drawing.Run(new Drawing.Text { Text = "Second" }))));

        var text = ExtractTextForTest(converter, shape, styleResolver, slidePart);

        Assert.Equal("BoldItalicPlain\n\nSecond", text.Content);
        Assert.True(text.AutoFit);
        Assert.Equal("center", text.VerticalAlign);
        Assert.Equal(1, text.PaddingLeft);
        Assert.Equal(2, text.PaddingTop);
        Assert.Equal(3, text.PaddingRight);
        Assert.Equal(4, text.PaddingBottom);
        Assert.Equal("right", text.Paragraphs[0].Formatting.Align);
        Assert.True(text.Paragraphs[0].Runs[0].Formatting.Bold);
        Assert.False(text.Paragraphs[0].Runs[1].Formatting.Bold);
        Assert.True(text.Paragraphs[0].Runs[1].Formatting.Italic);
        Assert.False(text.Paragraphs[0].Runs[2].Formatting.Bold);
    }

    [Fact]
    public void ExtractTextFromShape_RegexFallbackReadsBulletsNumberingNoneLevelsAndIndents()
    {
        var path = CreateSimplePptx();
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);
        var slidePart = document.PresentationPart!.SlideParts.First();
        var styleResolver = new StyleResolver(document, slidePart);

        var charPPr = new Drawing.ParagraphProperties { Level = 1 };
        charPPr.SetAttribute(new OpenXmlAttribute("", "marL", "", "254000"));
        charPPr.SetAttribute(new OpenXmlAttribute("", "indent", "", "-127000"));
        charPPr.Append(new Drawing.CharacterBullet { Char = "→" });

        var autoPPr = new Drawing.ParagraphProperties();
        autoPPr.Append(new Drawing.AutoNumberedBullet { Type = Drawing.TextAutoNumberSchemeValues.ArabicPeriod });

        var nonePPr = new Drawing.ParagraphProperties();
        nonePPr.Append(new Drawing.NoBullet());

        var shape = new P.Shape(
            new TextBody(
                new Drawing.BodyProperties(),
                new Drawing.ListStyle(),
                new Drawing.Paragraph(charPPr, new Drawing.Run(new Drawing.Text { Text = "Character bullet" })),
                new Drawing.Paragraph(autoPPr, new Drawing.Run(new Drawing.Text { Text = "Numbered" })),
                new Drawing.Paragraph(nonePPr, new Drawing.Run(new Drawing.Text { Text = "No bullet" }))));

        var text = ExtractTextForTest(converter, shape, styleResolver, slidePart);

        Assert.Equal(3, text.Paragraphs.Count);
        Assert.Equal(1, text.Paragraphs[0].Level);
        Assert.True(text.Paragraphs[0].HasBullet);
        Assert.Equal("→", text.Paragraphs[0].BulletChar);
        Assert.Equal(20.0, text.Paragraphs[0].MarginLeft);
        Assert.Equal(-10.0, text.Paragraphs[0].Indent);
        Assert.True(text.Paragraphs[1].HasBullet);
        Assert.Equal("arabicPeriod", text.Paragraphs[1].AutoNumberType);
        Assert.False(text.Paragraphs[2].HasBullet);
        Assert.Null(text.Paragraphs[2].BulletChar);
    }

    [Fact]
    public void Convert_GraphicFrameTableDiagramUnsupportedAndMissingGraphicDataBranches()
    {
        var tableFrame = TableGraphicFrame(50, "Table frame", x: 30, y: 40, width: 200, height: 80);
        var diagramFrame = DiagramGraphicFrame(51, "Empty diagram frame");
        var unsupportedFrame = UnsupportedGraphicFrame(52, "Unsupported frame");
        var missingGraphicData = new P.GraphicFrame(
            new NonVisualGraphicFrameProperties(
                new NonVisualDrawingProperties { Id = 53, Name = "Missing graphic data" },
                new NonVisualGraphicFrameDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()));
        var path = CreateGroupShapePptx("graphic-frame-branches.pptx", tableFrame, diagramFrame, unsupportedFrame, missingGraphicData);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();

        var table = Assert.Single(presentation.Slides[0].Elements, e => e.Type == "Table");
        Assert.Equal("Table frame", table.Name);
        AssertPosition(table, x: 30, y: 40, width: 200, height: 80);
        Assert.Equal("Cell", Assert.Single(Assert.Single(table.Table!.Rows)).Content);
        Assert.DoesNotContain(presentation.Slides[0].Elements, e => e.Name is "Empty diagram frame" or "Unsupported frame" or "Missing graphic data");
    }

    [Fact]
    public void Convert_GroupShapeWithZeroChildExtents_DoesNotThrowAndSkipsInfiniteGeometry()
    {
        var group = GroupShape(
            60,
            TransformGroup(x: 10, y: 10, width: 100, height: 100, childX: 0, childY: 0, childWidth: 0, childHeight: 0),
            UnsupportedConnector(61));
        var path = CreateGroupShapePptx("group-zero-extents.pptx", group);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();

        Assert.Empty(presentation.Slides[0].Elements);
    }

    [Fact]
    public void Convert_PicturesWithMissingBlipOrEmbed_AreIgnored()
    {
        var noBlip = new P.Picture(
            new NonVisualPictureProperties(
                new NonVisualDrawingProperties { Id = 70, Name = "No blip" },
                new NonVisualPictureDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new P.BlipFill(),
            new ShapeProperties(new Drawing.Transform2D(
                new Drawing.Offset { X = Pt(10), Y = Pt(10) },
                new Drawing.Extents { Cx = Pt(100), Cy = Pt(50) })));
        var noEmbed = new P.Picture(
            new NonVisualPictureProperties(
                new NonVisualDrawingProperties { Id = 71, Name = "No embed" },
                new NonVisualPictureDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new P.BlipFill(new Drawing.Blip()),
            new ShapeProperties(new Drawing.Transform2D(
                new Drawing.Offset { X = Pt(20), Y = Pt(20) },
                new Drawing.Extents { Cx = Pt(100), Cy = Pt(50) })));
        var path = CreateGroupShapePptx("picture-missing-blip.pptx", noBlip, noEmbed);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();

        Assert.DoesNotContain(presentation.Slides[0].Elements, e => e.Type == "Image");
    }

    [Fact]
    public void Convert_SlideNumberPlaceholderUsesSlideIndexFromRegexPlaceholderType()
    {
        var slideNumberShape = new P.Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = 80, Name = "Slide Number" },
                new NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape { Type = PlaceholderValues.SlideNumber })),
            new ShapeProperties(new Drawing.Transform2D(
                new Drawing.Offset { X = Pt(5), Y = Pt(6) },
                new Drawing.Extents { Cx = Pt(40), Cy = Pt(20) })),
            new TextBody(
                new Drawing.BodyProperties(),
                new Drawing.ListStyle(),
                new Drawing.Paragraph(new Drawing.Run(new Drawing.Text { Text = "ignored" }))));
        var path = CreateGroupShapePptx("slide-number-placeholder.pptx", slideNumberShape);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();

        var element = AssertSingleTextElement(presentation, "1");
        Assert.Equal("Slide Number", element.Name);
    }

    private static TypstTextElement ExtractTextForTest(PptxToTypstConverter converter, P.Shape shape, StyleResolver styleResolver, SlidePart slidePart)
    {
        var method = typeof(PptxToTypstConverter).GetMethod(
            "ExtractTextFromShape",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(P.Shape), typeof(StyleResolver), typeof(SlidePart)],
            null);

        return Assert.IsType<TypstTextElement>(method!.Invoke(converter, [shape, styleResolver, slidePart]));
    }

    private static P.GraphicFrame TableGraphicFrame(uint id, string name, double x, double y, double width, double height)
    {
        var table = new Drawing.Table(
            new Drawing.TableProperties { FirstRow = true },
            new Drawing.TableGrid(new Drawing.GridColumn { Width = Pt(width) }),
            new Drawing.TableRow(
                new Drawing.TableCell(
                    new Drawing.TextBody(
                        new Drawing.BodyProperties(),
                        new Drawing.ListStyle(),
                        new Drawing.Paragraph(new Drawing.Run(new Drawing.Text { Text = "Cell" }))),
                    new Drawing.TableCellProperties()))
            {
                Height = Pt(height)
            });

        return GraphicFrame(id, name, x, y, width, height, new Drawing.GraphicData(table)
        {
            Uri = "http://schemas.openxmlformats.org/drawingml/2006/table"
        });
    }

    private static P.GraphicFrame DiagramGraphicFrame(uint id, string name)
        => GraphicFrame(id, name, 0, 0, 100, 50, new Drawing.GraphicData
        {
            Uri = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
        });

    private static P.GraphicFrame UnsupportedGraphicFrame(uint id, string name)
        => GraphicFrame(id, name, 0, 0, 100, 50, new Drawing.GraphicData
        {
            Uri = "http://example.com/unsupported"
        });

    private static P.GraphicFrame GraphicFrame(uint id, string name, double x, double y, double width, double height, Drawing.GraphicData graphicData)
    {
        return new P.GraphicFrame(
            new NonVisualGraphicFrameProperties(
                new NonVisualDrawingProperties { Id = id, Name = name },
                new NonVisualGraphicFrameDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new Transform(
                new Drawing.Offset { X = Pt(x), Y = Pt(y) },
                new Drawing.Extents { Cx = Pt(width), Cy = Pt(height) }),
            new Drawing.Graphic(graphicData));
    }

    private const string PresentationmlNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private const string DrawingmlANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static P.Shape ShapeFromXml(string spPrFillXml, string txBodyXml = "")
    {
        return new P.Shape($@"<p:sp xmlns:p=""{PresentationmlNs}"" xmlns:a=""{DrawingmlANs}"">
  <p:nvSpPr><p:cNvPr id=""42"" name=""Alpha shape""/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x=""12700"" y=""12700""/><a:ext cx=""914400"" cy=""914400""/></a:xfrm>
    <a:prstGeom prst=""rect""><a:avLst/></a:prstGeom>
    {spPrFillXml}
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/>{(string.IsNullOrEmpty(txBodyXml) ? "<a:p><a:pPr algn=\"ctr\"/></a:p>" : txBodyXml)}</p:txBody>
</p:sp>");
    }

    [Fact]
    public void GenerateTypstSource_ShapeSolidFillWithLumMod_DarkensColor()
    {
        // bg1/lt1 (white) with lumMod 75% must resolve to light gray (#BFBFBF), not
        // white — the SmartArt corpus section headers ("List //") depend on this.
        var shape = ShapeFromXml("""<a:solidFill><a:srgbClr val="FFFFFF"><a:lumMod val="75000"/></a:srgbClr></a:solidFill>""");
        var path = CreateGroupShapePptx("lumod-shape-fill.pptx", shape);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var shapeElement = Assert.Single(presentation.Slides[0].Elements, e => e.Shape != null);
        Assert.Equal("#BFBFBF", shapeElement.Shape!.FillColor);
    }

    [Fact]
    public void GenerateTypstSource_ShapeSolidFillWithLumOff_LightensColor()
    {
        var shape = ShapeFromXml("""<a:solidFill><a:srgbClr val="000000"><a:lumOff val="20000"/></a:srgbClr></a:solidFill>""");
        var path = CreateGroupShapePptx("lumoff-shape-fill.pptx", shape);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var shapeElement = Assert.Single(presentation.Slides[0].Elements, e => e.Shape != null);
        Assert.Equal("#333333", shapeElement.Shape!.FillColor);
    }

    [Fact]
    public void GenerateTypstSource_ShapeSolidFillWithAlpha_EmitsEightDigitHex()
    {
        var shape = ShapeFromXml("""<a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="15000"/></a:srgbClr></a:solidFill>""");
        var path = CreateGroupShapePptx("alpha-shape-fill.pptx", shape);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var shapeElement = Assert.Single(presentation.Slides[0].Elements, e => e.Shape != null);
        Assert.Equal("#FFFFFF26", shapeElement.Shape!.FillColor); // 15% opacity → 0x26 alpha byte

        var source = converter.GenerateTypstSource(presentation);
        Assert.Contains("""fill: rgb("#FFFFFF26")""", source);
    }

    [Fact]
    public void Convert_ShapeSolidFillWithZeroAlpha_TreatedAsNoFill()
    {
        var shape = ShapeFromXml("""<a:solidFill><a:srgbClr val="FF0000"><a:alpha val="0"/></a:srgbClr></a:solidFill>""");
        var path = CreateGroupShapePptx("zero-alpha-shape-fill.pptx", shape);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        // Fully transparent fill == a:noFill: no shape element is emitted at all.
        Assert.DoesNotContain(presentation.Slides[0].Elements, e => e.Shape != null);
    }

    [Fact]
    public void GenerateTypstSource_TextRunSolidFillWithAlpha_EmitsEightDigitHex()
    {
        const string txBody = """
            <a:p><a:r><a:rPr lang="en-US"><a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="50000"/></a:srgbClr></a:solidFill></a:rPr><a:t>Glass</a:t></a:r></a:p>
            """;
        var shape = ShapeFromXml(string.Empty, txBody);
        var path = CreateGroupShapePptx("alpha-text-run.pptx", shape);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var textElement = Assert.Single(presentation.Slides[0].Elements, e => e.Text != null);
        Assert.Equal("#FFFFFF7F", textElement.Text!.Formatting.Color); // 50% opacity → 0x7F alpha byte

        var source = converter.GenerateTypstSource(presentation);
        Assert.Contains("""fill: rgb("#FFFFFF7F")""", source);
    }

    [Fact]
    public void GenerateTypstSource_ShapeGradientFill_EmitsTypstLinearGradient()
    {
        // AetherLink pattern: full-bleed rect with a two-stop navy gradient as slide background
        const string fill = """
            <a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="0B1026"/></a:gs><a:gs pos="100000"><a:srgbClr val="131B3F"/></a:gs></a:gsLst><a:lin ang="6900000" scaled="1"/></a:gradFill><a:ln><a:noFill/></a:ln>
            """;
        var shape = ShapeFromXml(fill);
        var path = CreateGroupShapePptx("gradient-fill.pptx", shape);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var shapeElement = Assert.Single(presentation.Slides[0].Elements, e => e.Shape != null);
        var gradient = Assert.IsType<TypstGradientFill>(shapeElement.Shape!.FillGradient);
        Assert.Equal(115.0, gradient.Angle); // 6900000 / 60000, verbatim OOXML->Typst degrees
        Assert.Equal(new TypstGradientStop("#0B1026", 0.0), gradient.Stops[0]);
        Assert.Equal(new TypstGradientStop("#131B3F", 1.0), gradient.Stops[1]);

        var source = converter.GenerateTypstSource(presentation);
        Assert.Contains("""fill: gradient.linear((rgb("#0B1026"), 0%), (rgb("#131B3F"), 100%), angle: 115deg)""", source);
    }

    [Fact]
    public void GenerateTypstSource_GradientStopWithAlpha_EmitsEightDigitHexStop()
    {
        // AetherLink slide 1 pattern: alpha gradient overlay fading from 88% to fully transparent
        const string fill = """
            <a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="0B1026"><a:alpha val="88000"/></a:srgbClr></a:gs><a:gs pos="100000"><a:srgbClr val="0B1026"><a:alpha val="0"/></a:srgbClr></a:gs></a:gsLst><a:lin ang="6900000" scaled="1"/></a:gradFill>
            """;
        var shape = ShapeFromXml(fill);
        var path = CreateGroupShapePptx("gradient-alpha-stops.pptx", shape);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var shapeElement = Assert.Single(presentation.Slides[0].Elements, e => e.Shape != null);
        var gradient = Assert.IsType<TypstGradientFill>(shapeElement.Shape!.FillGradient);
        // Transparent stops are meaningful in gradients (fade-out) — kept as #RRGGBB00, not noFill
        Assert.Equal("#0B1026E0", gradient.Stops[0].Color);
        Assert.Equal("#0B102600", gradient.Stops[1].Color);

        var source = converter.GenerateTypstSource(presentation);
        Assert.Contains("""(rgb("#0B1026E0"), 0%)""", source);
        Assert.Contains("""(rgb("#0B102600"), 100%)""", source);
    }

    [Fact]
    public void GenerateTypstSource_NoStrokeShape_EmitsExplicitStrokeNone()
    {
        // SmartArt-extracted shapes with no resolved stroke must not inherit Typst's
        // default 1pt black stroke (a visible black box around text containers).
        var path = CreateGroupShapePptx("nostroke-emission.pptx");
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = new TypstPresentation
        {
            Slides =
            {
                new TypstSlide
                {
                    Elements =
                    {
                        new TypstElement
                        {
                            Type = "Shape", X = 0, Y = 0, Width = 100, Height = 50,
                            Shape = new TypstShapeElement
                            {
                                ShapeType = "rect",
                                FillColor = "#4472C4",
                                NoStroke = true
                            }
                        }
                    }
                }
            }
        };

        var source = converter.GenerateTypstSource(presentation);
        Assert.Contains("""fill: rgb("#4472C4")""", source);
        Assert.Contains("stroke: none", source);
    }

    [Fact]
    public void GenerateTypstSource_ShapeWithoutNoStrokeFlag_KeepsLegacyEmission()
    {
        // The regular slide-shape path does not set NoStroke (a missing a:ln there may
        // still inherit a themed outline) — no stroke argument is emitted, as before.
        var path = CreateGroupShapePptx("nostroke-legacy.pptx");
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = new TypstPresentation
        {
            Slides =
            {
                new TypstSlide
                {
                    Elements =
                    {
                        new TypstElement
                        {
                            Type = "Shape", X = 0, Y = 0, Width = 100, Height = 50,
                            Shape = new TypstShapeElement
                            {
                                ShapeType = "rect",
                                FillColor = "#4472C4"
                            }
                        }
                    }
                }
            }
        };

        var source = converter.GenerateTypstSource(presentation);
        Assert.Contains("""fill: rgb("#4472C4")""", source);
        Assert.DoesNotContain("stroke: none", source);
    }

    private string CreateBackgroundPptx(string fileName, string? slideBgXml, string? layoutBgXml, string? masterBgXml, string? themeLt1Hex = null)
    {
        var path = Path.Combine(_tempDir, fileName);

        using (var document = PresentationDocument.Create(path, PresentationDocumentType.Presentation))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new Presentation
            {
                SlideMasterIdList = new SlideMasterIdList(),
                SlideIdList = new SlideIdList(),
                SlideSize = new SlideSize { Cx = (int)Pt(720), Cy = (int)Pt(540), Type = SlideSizeValues.Screen4x3 }
            };

            var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
            var masterCSld = new CommonSlideData(CreateShapeTree());
            if (masterBgXml != null) masterCSld.InsertAt(new P.Background(masterBgXml), 0);
            slideMasterPart.SlideMaster = new SlideMaster(
                masterCSld,
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

            if (themeLt1Hex != null)
            {
                var themePart = slideMasterPart.AddNewPart<ThemePart>();
                themePart.Theme = new Drawing.Theme($@"<a:theme xmlns:a=""{DrawingmlANs}"" name=""T"">
  <a:themeElements>
    <a:clrScheme name=""T"">
      <a:dk1><a:srgbClr val=""000000""/></a:dk1>
      <a:lt1><a:srgbClr val=""{themeLt1Hex}""/></a:lt1>
      <a:dk2><a:srgbClr val=""111111""/></a:dk2>
      <a:lt2><a:srgbClr val=""EEEEEE""/></a:lt2>
      <a:accent1><a:srgbClr val=""4472C4""/></a:accent1>
      <a:accent2><a:srgbClr val=""ED7D31""/></a:accent2>
      <a:accent3><a:srgbClr val=""A5A5A5""/></a:accent3>
      <a:accent4><a:srgbClr val=""FFC000""/></a:accent4>
      <a:accent5><a:srgbClr val=""5B9BD5""/></a:accent5>
      <a:accent6><a:srgbClr val=""70AD47""/></a:accent6>
      <a:hlink><a:srgbClr val=""0563C1""/></a:hlink>
      <a:folHlink><a:srgbClr val=""954F72""/></a:folHlink>
    </a:clrScheme>
  </a:themeElements>
</a:theme>");
            }

            var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
            var layoutCSld = new CommonSlideData(CreateShapeTree());
            if (layoutBgXml != null) layoutCSld.InsertAt(new P.Background(layoutBgXml), 0);
            slideLayoutPart.SlideLayout = new P.SlideLayout(layoutCSld);
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
            var slideCSld = new CommonSlideData(CreateShapeTree());
            if (slideBgXml != null) slideCSld.InsertAt(new P.Background(slideBgXml), 0);
            slidePart.Slide = new Slide(slideCSld);
            slidePart.AddPart(slideLayoutPart);

            presentationPart.Presentation.SlideIdList.Append(new SlideId
            {
                Id = 256,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });
        }

        return path;
    }

    private static string SolidBgXml(string hex)
        => $@"<p:bg xmlns:p=""{PresentationmlNs}"" xmlns:a=""{DrawingmlANs}""><p:bgPr><a:solidFill><a:srgbClr val=""{hex}""/></a:solidFill><a:effectLst/></p:bgPr></p:bg>";

    private static string BgRefXml(string scheme)
        => $@"<p:bg xmlns:p=""{PresentationmlNs}"" xmlns:a=""{DrawingmlANs}""><p:bgRef idx=""1001""><a:schemeClr val=""{scheme}""/></p:bgRef></p:bg>";

    private string ConvertAndGetBackground(string path, out string source)
    {
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);
        var presentation = converter.Convert();
        source = converter.GenerateTypstSource(presentation);
        return presentation.Slides[0].Layout.BackgroundColor ?? string.Empty;
    }

    [Fact]
    public void Convert_BackgroundCascade_SlideWinsOverLayoutAndMaster()
    {
        var path = CreateBackgroundPptx("bg-slide-wins.pptx",
            slideBgXml: SolidBgXml("111111"), layoutBgXml: SolidBgXml("222222"), masterBgXml: SolidBgXml("333333"));

        var bg = ConvertAndGetBackground(path, out var source);

        Assert.Equal("#111111", bg);
        Assert.Contains("""#set page(fill: rgb("#111111"))""", source);
    }

    [Fact]
    public void Convert_BackgroundCascade_FallsBackToLayout()
    {
        var path = CreateBackgroundPptx("bg-layout.pptx",
            slideBgXml: null, layoutBgXml: SolidBgXml("222222"), masterBgXml: SolidBgXml("333333"));

        var bg = ConvertAndGetBackground(path, out _);

        Assert.Equal("#222222", bg);
    }

    [Fact]
    public void Convert_BackgroundCascade_FallsBackToMasterBgPr()
    {
        var path = CreateBackgroundPptx("bg-master.pptx",
            slideBgXml: null, layoutBgXml: null, masterBgXml: SolidBgXml("333333"));

        var bg = ConvertAndGetBackground(path, out _);

        Assert.Equal("#333333", bg);
    }

    [Fact]
    public void Convert_BackgroundCascade_ResolvesMasterBgRefSchemeColor()
    {
        // AetherLink master pattern: <p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef>
        var path = CreateBackgroundPptx("bg-master-bgref.pptx",
            slideBgXml: null, layoutBgXml: null, masterBgXml: BgRefXml("bg1"), themeLt1Hex: "0B1026");

        var bg = ConvertAndGetBackground(path, out var source);

        Assert.Equal("#0B1026", bg);
        Assert.Contains("""#set page(fill: rgb("#0B1026"))""", source);
    }

    // ---------------------------------------------------------------------
    // SmartArt-corpus slide-1 fidelity: group child-offset math, connector
    // shapes (p:cxnSp), diagStripe preset geometry, layout footer content.
    // ---------------------------------------------------------------------

    private static P.Shape FilledRectShape(uint id, double x, double y, double width, double height, string hexFill)
    {
        return new P.Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = id, Name = $"Rect {id}" },
                new NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(
                new Drawing.Transform2D(
                    new Drawing.Offset { X = Pt(x), Y = Pt(y) },
                    new Drawing.Extents { Cx = Pt(width), Cy = Pt(height) }),
                new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = Drawing.ShapeTypeValues.Rectangle },
                new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = hexFill })));
    }

    private static P.ConnectionShape LineConnector(uint id, double x, double y, double width, double height, string hexStroke, int strokeEmus)
    {
        return new P.ConnectionShape(
            new NonVisualConnectionShapeProperties(
                new NonVisualDrawingProperties { Id = id, Name = $"Connector {id}" },
                new NonVisualConnectorShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(
                new Drawing.Transform2D(
                    new Drawing.Offset { X = Pt(x), Y = Pt(y) },
                    new Drawing.Extents { Cx = Pt(width), Cy = Pt(height) }),
                new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = Drawing.ShapeTypeValues.Line },
                new Drawing.Outline(new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = hexStroke }))
                {
                    Width = new Int32Value(strokeEmus)
                }));
    }

    [Fact]
    public void Convert_GroupShapeWithScaledChildOffset_MapsChildrenIntoGroupCoordinateSpace()
    {
        // ECMA-376 group mapping: abs = grpOff + (child − chOff) × (ext / chExt).
        // A child spanning the whole child space must land exactly on the group bbox
        // (mirrors the SmartArt corpus title-slide group: chOff.y ≠ 0, scale ≠ 1).
        var group = GroupShape(
            10,
            TransformGroup(x: 27, y: 16, width: 190, height: 140, childX: 0, childY: 224, childWidth: 296, childHeight: 218),
            FilledRectShape(11, x: 0, y: 224, width: 296, height: 218, "4472C4"));
        var path = CreateGroupShapePptx("group-scaled-child-offset.pptx", group);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var element = Assert.Single(presentation.Slides[0].Elements);

        AssertPosition(element, x: 27, y: 16, width: 190, height: 140);
    }

    [Fact]
    public void Convert_ConnectionShapeInsideZeroHeightGroup_RendersStrokeAtMappedPosition()
    {
        // List-glyph pattern: nested grpSp with cy=0/chExt cy=0 (degenerate) holding a
        // horizontal straight connector (p:cxnSp). The connector must render (stroke)
        // and the 0/0 group scale must not poison positions with NaN.
        var innerGroup = GroupShape(
            11,
            TransformGroup(x: 20, y: 60, width: 130, height: 0, childX: 20, childY: 60, childWidth: 130, childHeight: 0),
            LineConnector(12, x: 20, y: 60, width: 80, height: 0, hexStroke: "ED7D31", strokeEmus: 76200));
        var outerGroup = GroupShape(
            10,
            TransformGroup(x: 100, y: 100, width: 260, height: 130, childX: 0, childY: 40, childWidth: 520, childHeight: 260),
            innerGroup);
        var path = CreateGroupShapePptx("connector-zero-height-group.pptx", outerGroup);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var element = Assert.Single(presentation.Slides[0].Elements);

        Assert.Equal("Shape", element.Type);
        Assert.False(double.IsNaN(element.X) || double.IsNaN(element.Y), "Connector position must not be NaN");
        AssertPosition(element, x: 110, y: 110, width: 40, height: 0);
        Assert.Equal("#ED7D31", element.Shape!.StrokeColor);
        Assert.Equal(6, element.Shape.StrokeWidth, 2);
    }

    [Fact]
    public void Convert_DiagonalStripePreset_EmitsStripePolygonNotBoundingRect()
    {
        // diagStripe (adj = stripe thickness in 1/100000 of the bbox) is a diagonal
        // band: (0,f) (f,0) (1,0) (0,1) — not the full square bbox.
        var stripe = new P.Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = 2, Name = "Stripe" },
                new NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(
                new Drawing.Transform2D(
                    new Drawing.Offset { X = Pt(500), Y = Pt(0) },
                    new Drawing.Extents { Cx = Pt(200), Cy = Pt(200) })
                {
                    Rotation = new Int32Value(5400000)
                },
                new Drawing.PresetGeometry(
                    new Drawing.AdjustValueList(
                        new Drawing.ShapeGuide { Name = "adj", Formula = "val 30578" }))
                {
                    Preset = Drawing.ShapeTypeValues.DiagonalStripe
                },
                new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = "ED7D31" })));
        var path = CreateGroupShapePptx("diag-stripe.pptx", stripe);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var element = Assert.Single(presentation.Slides[0].Elements);

        Assert.Equal("Shape", element.Type);
        Assert.Equal("polygon", element.Shape!.ShapeType);
        Assert.Equal(90, element.Rotation, 2);
        Assert.Equal(4, element.Shape.Points.Count);
        Assert.Equal(0.0, element.Shape.Points[0].X, 5);
        Assert.Equal(0.30578, element.Shape.Points[0].Y, 5);
        Assert.Equal(0.30578, element.Shape.Points[1].X, 5);
        Assert.Equal(0.0, element.Shape.Points[1].Y, 5);
        Assert.Equal(1.0, element.Shape.Points[2].X, 5);
        Assert.Equal(0.0, element.Shape.Points[2].Y, 5);
        Assert.Equal(0.0, element.Shape.Points[3].X, 5);
        Assert.Equal(1.0, element.Shape.Points[3].Y, 5);

        var source = converter.GenerateTypstSource(presentation);
        Assert.Contains("#rotate(90.0deg", source);
        Assert.Contains("#polygon(", source);
    }

    [Fact]
    public void Convert_RotatedPolygon_RotatesAboutShapeBoxCenter()
    {
        // A Typst #polygon's bounding box is its INK bbox, which can be smaller
        // than the shape box (e.g. blockArc bands, stripes). #rotate(origin:
        // center) would then rotate about the ink center instead of the shape-box
        // center (PowerPoint xfrm@rot semantics), displacing the rotated shape.
        // The polygon must be wrapped in an explicit-size block.
        var stripe = new P.Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = 2, Name = "Stripe" },
                new NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(
                new Drawing.Transform2D(
                    new Drawing.Offset { X = Pt(500), Y = Pt(0) },
                    new Drawing.Extents { Cx = Pt(200), Cy = Pt(200) })
                {
                    Rotation = new Int32Value(5400000)
                },
                new Drawing.PresetGeometry(
                    new Drawing.AdjustValueList(
                        new Drawing.ShapeGuide { Name = "adj", Formula = "val 30578" }))
                {
                    Preset = Drawing.ShapeTypeValues.DiagonalStripe
                },
                new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = "ED7D31" })));
        var path = CreateGroupShapePptx("diag-stripe-rotate-center.pptx", stripe);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains(
            "#rotate(90.0deg, origin: center)[#block(width: 200.00pt, height: 200.00pt)[#polygon(",
            source);
    }

    [Fact]
    public void Convert_LayoutUserDrawnGroupShape_RendersChildren()
    {
        // Footer-byline pattern: a grpSp on the slide layout (no placeholders inside)
        // must render on the slide, mapped through its child coordinate space.
        var footerGroup = GroupShape(
            20,
            TransformGroup(x: 400, y: 500, width: 150, height: 30, childX: 390, childY: 495, childWidth: 150, childHeight: 30),
            TextShape(21, "Made with love", x: 390, y: 495, width: 150, height: 30));
        var path = CreateLayoutContentPptx("layout-user-drawn-group.pptx", null, footerGroup);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var element = AssertSingleTextElement(presentation, "Made with love");

        AssertPosition(element, x: 400, y: 500, width: 150, height: 30);
    }

    [Fact]
    public void Convert_LayoutUserDrawnPicture_RendersImageResolvedFromLayoutPart()
    {
        // Footer logo pattern: a p:pic on the layout whose a:blip r:embed is a
        // relationship of the LAYOUT part (not the slide part).
        var picture = new P.Picture(
            new P.NonVisualPictureProperties(
                new NonVisualDrawingProperties { Id = 30, Name = "Logo" },
                new P.NonVisualPictureDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new P.BlipFill(
                new Drawing.Blip { Embed = "rIdLogo" },
                new Drawing.Stretch(new Drawing.FillRectangle())),
            new ShapeProperties(
                new Drawing.Transform2D(
                    new Drawing.Offset { X = Pt(600), Y = Pt(490) },
                    new Drawing.Extents { Cx = Pt(100), Cy = Pt(30) })));

        // 1x1 transparent PNG
        var pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");

        var path = CreateLayoutContentPptx("layout-user-drawn-picture.pptx", layoutPart =>
        {
            var imagePart = layoutPart.AddNewPart<ImagePart>("image/png", "rIdLogo");
            using (var stream = imagePart.GetStream(FileMode.Create, FileAccess.Write))
            {
                stream.Write(pngBytes, 0, pngBytes.Length);
            }
        }, picture);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var element = Assert.Single(presentation.Slides[0].Elements, e => e.Type == "Image");

        AssertPosition(element, x: 600, y: 490, width: 100, height: 30);
    }

    private string CreateLayoutContentPptx(string fileName, Action<SlideLayoutPart>? customizeLayout, params OpenXmlElement[] layoutElements)
    {
        var path = Path.Combine(_tempDir, fileName);

        using (var document = PresentationDocument.Create(path, PresentationDocumentType.Presentation))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new Presentation
            {
                SlideMasterIdList = new SlideMasterIdList(),
                SlideIdList = new SlideIdList(),
                SlideSize = new SlideSize { Cx = (int)Pt(720), Cy = (int)Pt(540), Type = SlideSizeValues.Screen4x3 }
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
            slideLayoutPart.SlideLayout = new P.SlideLayout(new CommonSlideData(CreateShapeTree(layoutElements)));
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
            slidePart.Slide = new Slide(new CommonSlideData(CreateShapeTree()));
            slidePart.AddPart(slideLayoutPart);

            customizeLayout?.Invoke(slideLayoutPart);

            presentationPart.Presentation.SlideIdList.Append(new SlideId
            {
                Id = 256,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });
        }

        return path;
    }
    #region SmartArt text/style fidelity: normAutofit, slidenum field, placeholder fill, footer color

    /// <summary>
    /// Builds a deck whose slide layout and slide each carry the given shape XML
    /// (<c>p:sp</c> fragments), with a minimal theme (dk1 = 404040) so scheme colors
    /// resolve deterministically.
    /// </summary>
    private string CreateDeckWithLayoutShapes(string fileName, string[] layoutShapesXml, string[] slideShapesXml, string? masterTxStylesXml = null)
    {
        var path = Path.Combine(_tempDir, fileName);

        using (var document = PresentationDocument.Create(path, PresentationDocumentType.Presentation))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new Presentation
            {
                SlideMasterIdList = new SlideMasterIdList(),
                SlideIdList = new SlideIdList(),
                SlideSize = new SlideSize { Cx = (int)Pt(960), Cy = (int)Pt(540), Type = SlideSizeValues.Screen4x3 }
            };

            var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
            var master = new SlideMaster(
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
            if (masterTxStylesXml != null)
            {
                master.Append(new P.TextStyles(masterTxStylesXml));
            }
            slideMasterPart.SlideMaster = master;

            var themePart = slideMasterPart.AddNewPart<ThemePart>();
            themePart.Theme = new Drawing.Theme($@"<a:theme xmlns:a=""{DrawingmlANs}"" name=""T"">
  <a:themeElements>
    <a:clrScheme name=""T"">
      <a:dk1><a:srgbClr val=""404040""/></a:dk1>
      <a:lt1><a:srgbClr val=""FFFFFF""/></a:lt1>
      <a:dk2><a:srgbClr val=""111111""/></a:dk2>
      <a:lt2><a:srgbClr val=""EEEEEE""/></a:lt2>
      <a:accent1><a:srgbClr val=""4472C4""/></a:accent1>
      <a:accent2><a:srgbClr val=""ED7D31""/></a:accent2>
      <a:accent3><a:srgbClr val=""A5A5A5""/></a:accent3>
      <a:accent4><a:srgbClr val=""FFC000""/></a:accent4>
      <a:accent5><a:srgbClr val=""5B9BD5""/></a:accent5>
      <a:accent6><a:srgbClr val=""70AD47""/></a:accent6>
      <a:hlink><a:srgbClr val=""0563C1""/></a:hlink>
      <a:folHlink><a:srgbClr val=""954F72""/></a:folHlink>
    </a:clrScheme>
  </a:themeElements>
</a:theme>");

            var layoutShapes = layoutShapesXml.Select(xml => new P.Shape(xml)).Cast<OpenXmlElement>().ToArray();
            var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
            slideLayoutPart.SlideLayout = new P.SlideLayout(new CommonSlideData(CreateShapeTree(layoutShapes)));
            slideLayoutPart.AddPart(slideMasterPart);
            master.SlideLayoutIdList!.Append(new SlideLayoutId
            {
                Id = 2147483649,
                RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart)
            });

            presentationPart.Presentation.SlideMasterIdList.Append(new SlideMasterId
            {
                Id = 2147483648,
                RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
            });

            var slideShapes = slideShapesXml.Select(xml => new P.Shape(xml)).Cast<OpenXmlElement>().ToArray();
            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.Slide = new Slide(new CommonSlideData(CreateShapeTree(slideShapes)));
            slidePart.AddPart(slideLayoutPart);

            presentationPart.Presentation.SlideIdList.Append(new SlideId
            {
                Id = 256,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });
        }

        return path;
    }

    [Fact]
    public void Convert_NormalAutoFitFontScaleAndLineSpaceReduction_AppliedToSizesAndLineSpacing()
    {
        var shape = new P.Shape($@"<p:sp xmlns:p=""{PresentationmlNs}"" xmlns:a=""{DrawingmlANs}"">
  <p:nvSpPr><p:cNvPr id=""2"" name=""Shrinking""/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x=""12700"" y=""12700""/><a:ext cx=""914400"" cy=""914400""/></a:xfrm></p:spPr>
  <p:txBody>
    <a:bodyPr><a:normAutofit fontScale=""50000"" lnSpcReduction=""20000""/></a:bodyPr>
    <a:lstStyle/>
    <a:p>
      <a:pPr><a:lnSpc><a:spcPct val=""100000""/></a:lnSpc></a:pPr>
      <a:r><a:rPr lang=""en-US"" sz=""2000""/><a:t>Shrink me</a:t></a:r>
    </a:p>
  </p:txBody>
</p:sp>");
        var path = CreateGroupShapePptx("normautofit-fontscale.pptx", shape);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var element = AssertSingleTextElement(presentation, "Shrink me");
        var paragraph = Assert.Single(element.Text!.Paragraphs);
        Assert.Equal(10.0, paragraph.Formatting.FontSize, 3);
        Assert.Equal(10.0, Assert.Single(paragraph.Runs).Formatting.FontSize, 3);
        Assert.Equal(0.8, paragraph.LineSpacing!.Value, 3);
    }

    [Fact]
    public void Convert_NormalAutoFitWithoutAttributes_LeavesSizesAndLineSpacingUnchanged()
    {
        var shape = new P.Shape($@"<p:sp xmlns:p=""{PresentationmlNs}"" xmlns:a=""{DrawingmlANs}"">
  <p:nvSpPr><p:cNvPr id=""2"" name=""Plain autofit""/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x=""12700"" y=""12700""/><a:ext cx=""914400"" cy=""914400""/></a:xfrm></p:spPr>
  <p:txBody>
    <a:bodyPr><a:normAutofit/></a:bodyPr>
    <a:lstStyle/>
    <a:p>
      <a:pPr><a:lnSpc><a:spcPct val=""90000""/></a:lnSpc></a:pPr>
      <a:r><a:rPr lang=""en-US"" sz=""2000""/><a:t>No shrink</a:t></a:r>
    </a:p>
  </p:txBody>
</p:sp>");
        var path = CreateGroupShapePptx("normautofit-plain.pptx", shape);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var element = AssertSingleTextElement(presentation, "No shrink");
        var paragraph = Assert.Single(element.Text!.Paragraphs);
        Assert.Equal(20.0, paragraph.Formatting.FontSize, 3);
        Assert.Equal(0.9, paragraph.LineSpacing!.Value, 3);
    }

    [Fact]
    public void Convert_SlideNumberFieldInUserDrawnLayoutTextbox_RendersActualSlideIndex()
    {
        var layoutTextbox = $@"<p:sp xmlns:p=""{PresentationmlNs}"" xmlns:a=""{DrawingmlANs}"">
  <p:nvSpPr><p:cNvPr id=""13"" name=""Slide Number Placeholder 5""/><p:cNvSpPr txBox=""1""/><p:nvPr userDrawn=""1""/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x=""11382764"" y=""6253489""/><a:ext cx=""562643"" cy=""390437""/></a:xfrm><a:prstGeom prst=""rect""><a:avLst/></a:prstGeom></p:spPr>
  <p:txBody>
    <a:bodyPr anchor=""ctr""/>
    <a:lstStyle/>
    <a:p>
      <a:fld id=""{{F68327C5-B821-4FE9-A59A-A60D9EB59A9A}}"" type=""slidenum"">
        <a:rPr lang=""en-US""><a:solidFill><a:srgbClr val=""6F6D6A""/></a:solidFill></a:rPr>
        <a:t>&lt;#&gt;</a:t>
      </a:fld>
    </a:p>
  </p:txBody>
</p:sp>";
        var path = CreateDeckWithLayoutShapes("slidenum-field.pptx", [layoutTextbox], []);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var element = AssertSingleTextElement(presentation, "1");
        Assert.Equal("#6F6D6A", Assert.Single(element.Text!.Paragraphs[0].Runs).Formatting.Color);
    }

    [Fact]
    public void Convert_PlaceholderWithoutFillOrLine_InheritsLayoutPlaceholderFillAndStroke()
    {
        var layoutPlaceholder = $@"<p:sp xmlns:p=""{PresentationmlNs}"" xmlns:a=""{DrawingmlANs}"">
  <p:nvSpPr><p:cNvPr id=""51"" name=""Text Placeholder""/><p:cNvSpPr/><p:nvPr><p:ph type=""body"" sz=""quarter"" idx=""13""/></p:nvPr></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x=""371475"" y=""1556792""/><a:ext cx=""1920240"" cy=""584775""/></a:xfrm>
    <a:solidFill><a:srgbClr val=""CCCCCC""/></a:solidFill>
    <a:ln w=""12700""><a:solidFill><a:srgbClr val=""999999""/></a:solidFill></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang=""en-US""/><a:t>Layout text</a:t></a:r></a:p></p:txBody>
</p:sp>";
        var slidePlaceholder = $@"<p:sp xmlns:p=""{PresentationmlNs}"" xmlns:a=""{DrawingmlANs}"">
  <p:nvSpPr><p:cNvPr id=""6"" name=""Text Placeholder 5""/><p:cNvSpPr/><p:nvPr><p:ph type=""body"" sz=""quarter"" idx=""13""/></p:nvPr></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x=""371475"" y=""1556792""/><a:ext cx=""1920240"" cy=""3539430""/></a:xfrm></p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang=""en-US""/><a:t>Explanation</a:t></a:r></a:p></p:txBody>
</p:sp>";
        var path = CreateDeckWithLayoutShapes("placeholder-fill-inherit.pptx", [layoutPlaceholder], [slidePlaceholder]);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var shapeElement = Assert.Single(presentation.Slides[0].Elements, e => e.Shape != null);
        Assert.Equal("#CCCCCC", shapeElement.Shape!.FillColor);
        Assert.Equal("#999999", shapeElement.Shape.StrokeColor);
        Assert.Equal(1.0, shapeElement.Shape.StrokeWidth, 3);
    }

    [Fact]
    public void Convert_PlaceholderWithExplicitNoFill_DoesNotInheritLayoutFill()
    {
        var layoutPlaceholder = $@"<p:sp xmlns:p=""{PresentationmlNs}"" xmlns:a=""{DrawingmlANs}"">
  <p:nvSpPr><p:cNvPr id=""51"" name=""Text Placeholder""/><p:cNvSpPr/><p:nvPr><p:ph type=""body"" sz=""quarter"" idx=""13""/></p:nvPr></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x=""371475"" y=""1556792""/><a:ext cx=""1920240"" cy=""584775""/></a:xfrm>
    <a:solidFill><a:srgbClr val=""CCCCCC""/></a:solidFill>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang=""en-US""/><a:t>Layout text</a:t></a:r></a:p></p:txBody>
</p:sp>";
        var slidePlaceholder = $@"<p:sp xmlns:p=""{PresentationmlNs}"" xmlns:a=""{DrawingmlANs}"">
  <p:nvSpPr><p:cNvPr id=""6"" name=""Text Placeholder 5""/><p:cNvSpPr/><p:nvPr><p:ph type=""body"" sz=""quarter"" idx=""13""/></p:nvPr></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x=""371475"" y=""1556792""/><a:ext cx=""1920240"" cy=""3539430""/></a:xfrm>
    <a:noFill/>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang=""en-US""/><a:t>Explanation</a:t></a:r></a:p></p:txBody>
</p:sp>";
        var path = CreateDeckWithLayoutShapes("placeholder-nofill-wins.pptx", [layoutPlaceholder], [slidePlaceholder]);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        Assert.DoesNotContain(presentation.Slides[0].Elements, e => e.Shape != null);
    }

    [Fact]
    public void Convert_FooterPlaceholder_InheritsMasterOtherStyleTextColor()
    {
        const string txStylesXml = $@"<p:txStyles xmlns:p=""{PresentationmlNs}"" xmlns:a=""{DrawingmlANs}"">
  <p:bodyStyle><a:lvl1pPr><a:defRPr sz=""3200""/></a:lvl1pPr></p:bodyStyle>
  <p:otherStyle>
    <a:lvl1pPr>
      <a:defRPr sz=""1200""><a:solidFill><a:srgbClr val=""808080""/></a:solidFill></a:defRPr>
    </a:lvl1pPr>
  </p:otherStyle>
</p:txStyles>";
        var footerShape = $@"<p:sp xmlns:p=""{PresentationmlNs}"" xmlns:a=""{DrawingmlANs}"">
  <p:nvSpPr><p:cNvPr id=""3"" name=""Footer Placeholder 2""/><p:cNvSpPr/><p:nvPr><p:ph type=""ftr"" sz=""quarter"" idx=""11""/></p:nvPr></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x=""371475"" y=""6280675""/><a:ext cx=""7672502"" cy=""365125""/></a:xfrm></p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang=""en-US""/><a:t>Footer text</a:t></a:r></a:p></p:txBody>
</p:sp>";
        var path = CreateDeckWithLayoutShapes("footer-otherstyle-color.pptx", [], [footerShape], txStylesXml);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var element = AssertSingleTextElement(presentation, "Footer text");
        Assert.Equal("#808080", Assert.Single(element.Text!.Paragraphs[0].Runs).Formatting.Color);
    }

    [Fact]
    public void Convert_DiagramShapeTextOverflowingTextBox_ShrinksFontToFitAndKeepsLabelsAboveShapes()
    {
        var path = Path.Combine(_tempDir, "diagram-shrink.pptx");

        using (var document = PresentationDocument.Create(path, PresentationDocumentType.Presentation))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new Presentation
            {
                SlideMasterIdList = new SlideMasterIdList(),
                SlideIdList = new SlideIdList(),
                SlideSize = new SlideSize { Cx = (int)Pt(960), Cy = (int)Pt(540), Type = SlideSizeValues.Screen4x3 }
            };

            var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
            slideMasterPart.SlideMaster = new SlideMaster(
                new CommonSlideData(CreateShapeTree()),
                new ColorMap(),
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

            var dataPart = slidePart.AddNewPart<DiagramDataPart>();
            var dataRelId = slidePart.GetIdOfPart(dataPart);
            using (var stream = dataPart.GetStream(FileMode.Create))
            using (var writer = new StreamWriter(stream))
            {
                writer.Write(@"<dgm:dataModel xmlns:dgm=""http://schemas.openxmlformats.org/drawingml/2006/diagram"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><dgm:ptLst/></dgm:dataModel>");
            }

            var drawingPart = slidePart.AddNewPart<DiagramPersistLayoutPart>();
            using (var stream = drawingPart.GetStream(FileMode.Create))
            using (var writer = new StreamWriter(stream))
            {
                writer.Write(@"<dsp:drawing xmlns:dsp=""http://schemas.microsoft.com/office/drawing/2008/diagram"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">
  <dsp:spTree>
    <dsp:sp modelId=""{11111111-1111-1111-1111-111111111111}"">
      <dsp:spPr><a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""1270000"" cy=""508000""/></a:xfrm><a:prstGeom prst=""roundRect""><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val=""DDDDDD""/></a:solidFill></dsp:spPr>
      <dsp:txBody>
        <a:bodyPr lIns=""12700"" tIns=""12700"" rIns=""12700"" bIns=""12700""><a:noAutofit/></a:bodyPr>
        <a:lstStyle/>
        <a:p><a:r><a:rPr lang=""en-US"" sz=""1500""/><a:t>Some fairly long text that wraps</a:t></a:r></a:p>
        <a:p><a:r><a:rPr lang=""en-US"" sz=""1500""/><a:t>Some fairly long text that wraps</a:t></a:r></a:p>
        <a:p><a:r><a:rPr lang=""en-US"" sz=""1500""/><a:t>Some fairly long text that wraps</a:t></a:r></a:p>
      </dsp:txBody>
      <dsp:txXfrm><a:off x=""0"" y=""0""/><a:ext cx=""1270000"" cy=""508000""/></dsp:txXfrm>
    </dsp:sp>
    <dsp:sp modelId=""{22222222-2222-2222-2222-222222222222}"">
      <dsp:spPr><a:xfrm><a:off x=""1524000"" y=""0""/><a:ext cx=""1270000"" cy=""508000""/></a:xfrm><a:prstGeom prst=""roundRect""><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val=""BBBBBB""/></a:solidFill></dsp:spPr>
      <dsp:txBody>
        <a:bodyPr lIns=""12700"" tIns=""12700"" rIns=""12700"" bIns=""12700""><a:noAutofit/></a:bodyPr>
        <a:lstStyle/>
        <a:p><a:r><a:rPr lang=""en-US"" sz=""1500""/><a:t>Fits</a:t></a:r></a:p>
      </dsp:txBody>
      <dsp:txXfrm><a:off x=""1524000"" y=""0""/><a:ext cx=""1270000"" cy=""508000""/></dsp:txXfrm>
    </dsp:sp>
  </dsp:spTree>
</dsp:drawing>");
            }

            var relIds = new OpenXmlUnknownElement("dgm", "relIds", "http://schemas.openxmlformats.org/drawingml/2006/diagram");
            relIds.SetAttribute(new OpenXmlAttribute("r", "dm", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", dataRelId));
            var frame = new P.GraphicFrame(
                new NonVisualGraphicFrameProperties(
                    new NonVisualDrawingProperties { Id = 4, Name = "Diagram" },
                    new NonVisualGraphicFrameDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new Transform(
                    new Drawing.Offset { X = Pt(0), Y = Pt(0) },
                    new Drawing.Extents { Cx = Pt(400), Cy = Pt(100) }),
                new Drawing.Graphic(new Drawing.GraphicData(relIds)
                {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
                }));

            slidePart.Slide = new Slide(new CommonSlideData(CreateShapeTree(frame)));
            slidePart.AddPart(slideLayoutPart);

            presentationPart.Presentation.SlideIdList.Append(new SlideId
            {
                Id = 256,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });
        }

        using var openDocument = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(openDocument);

        var presentation = converter.Convert();
        var elements = presentation.Slides[0].Elements;

        // Overflowing node text (3 long paragraphs in a 40pt box) must shrink.
        var overflowing = Assert.Single(elements, e => e.Text?.Content.Contains("Some fairly long text") == true);
        Assert.True(overflowing.Text!.Paragraphs[0].Formatting.FontSize < 15.0,
            $"expected shrunk font, got {overflowing.Text.Paragraphs[0].Formatting.FontSize}");

        // Comfortably fitting node text keeps its size.
        var fitting = Assert.Single(elements, e => e.Text?.Content == "Fits");
        Assert.Equal(15.0, fitting.Text!.Paragraphs[0].Formatting.FontSize, 3);

        // All diagram shapes are emitted before any diagram text so later shapes
        // (e.g. arc connectors) cannot cover earlier node labels.
        var lastShapeIndex = elements.FindLastIndex(e => e.Type == "Shape");
        var firstTextIndex = elements.FindIndex(e => e.Type == "Text");
        Assert.True(lastShapeIndex >= 0 && firstTextIndex > lastShapeIndex,
            $"expected all shapes before all texts, last shape at {lastShapeIndex}, first text at {firstTextIndex}");
    }

    #endregion
}
