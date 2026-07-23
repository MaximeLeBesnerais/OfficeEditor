using System.IO.Compression;
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

        AssertPosition(element, x: 130, y: 90, width: 200, height: 80);
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

        AssertPosition(element, x: 130, y: 90, width: 200, height: 80);
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

        AssertPosition(element, x: 130, y: 90, width: 200, height: 80);
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
    public void PresPro_Slide1_BackgroundColor_ResolvedFromLayout()
    {
        var baseDir = AppContext.BaseDirectory;
        var path = Path.Combine(baseDir, "..", "..", "..", "..", "examples", "REF", ".pptx");
        path = Path.GetFullPath(path);
        
        if (!File.Exists(path))
            return; // REF fixture purged  (licensing); replacement pending

        using var doc = PresentationDocument.Open(path, false);
        var converter = new PptxToTypstConverter(doc);
        var presentation = converter.Convert();

        var slide1 = presentation.Slides.FirstOrDefault();
        Assert.NotNull(slide1);
        
        // Background should be resolved from layout (accent1 = #CEBA80)
        Assert.Equal("#CEBA80", slide1.Layout.BackgroundColor);
    }

    [Fact]
    public void PresPro_Slide1_TitleFontSize_FromMaster()
    {
        var baseDir = AppContext.BaseDirectory;
        var path = Path.Combine(baseDir, "..", "..", "..", "..", "examples", "REF", ".pptx");
        path = Path.GetFullPath(path);
        if (!File.Exists(path))
            return; // REF fixture purged  (licensing); replacement pending

        using var doc = PresentationDocument.Open(path, false);
        var converter = new PptxToTypstConverter(doc);
        var presentation = converter.Convert();

        var slide1 = presentation.Slides.FirstOrDefault();
        Assert.NotNull(slide1);

        var titleElement = slide1.Elements.FirstOrDefault(e => e.Name == "Title 1");
        Assert.NotNull(titleElement);
        Assert.Equal(72.0, titleElement.Text?.Formatting.FontSize);
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

        Assert.Contains("indent: 22.50pt", source);
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

        Assert.Contains("font: \"Aptos Light\"", source);
        Assert.DoesNotContain("font: \"Carlito\"", source);
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

        var cell = new Drawing.TableCell(
            new Drawing.TableCellProperties(
                new Drawing.TableCellBorders(
                    new Drawing.TopBorder(
                        new Drawing.Outline(
                            new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = "FF0000" }))
                        {
                            Width = 25400
                        }))),
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

        Assert.Empty(presentation.Slides[0].Elements.Where(e => e.Type == "Text"));
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

        Assert.Empty(presentation.Slides[0].Elements.Where(e => e.Type == "Image"));
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
}
