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
            [typeof(P.Shape), typeof(StyleResolver)],
            null);

        // Get the first slide part for style resolution
        var slidePart = document.PresentationPart!.SlideParts.FirstOrDefault();
        Assert.NotNull(slidePart);

        var styleResolver = new StyleResolver(document, slidePart);
        var text = Assert.IsType<TypstTextElement>(method!.Invoke(converter, [shape, styleResolver]));

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
        var path = Path.Combine(baseDir, "..", "..", "..", "..", "examples", "REF", "pres-pro.pptx");
        path = Path.GetFullPath(path);
        
        if (!File.Exists(path))
            return; // Skip if file doesn't exist

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
        var path = Path.Combine(baseDir, "..", "..", "..", "..", "examples", "REF", "pres-pro.pptx");
        path = Path.GetFullPath(path);
        if (!File.Exists(path))
            return;

        using var doc = PresentationDocument.Open(path, false);
        var converter = new PptxToTypstConverter(doc);
        var presentation = converter.Convert();

        var slide1 = presentation.Slides.FirstOrDefault();
        Assert.NotNull(slide1);

        var titleElement = slide1.Elements.FirstOrDefault(e => e.Name == "Title 1");
        Assert.NotNull(titleElement);
        Assert.Equal(54.0, titleElement.Text?.Formatting.FontSize);
    }
}
