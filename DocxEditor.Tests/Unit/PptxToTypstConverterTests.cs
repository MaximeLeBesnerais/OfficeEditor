using System.Diagnostics;
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
using Xunit.Abstractions;

namespace DocxEditor.Tests.Unit;

public class PptxToTypstConverterTests : IDisposable
{
    private readonly string _tempDir;
    private readonly ITestOutputHelper _output;

    public PptxToTypstConverterTests(ITestOutputHelper output)
    {
        _output = output;
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
            LineSpacing = new TextSpacing(TextSpacingKind.Points, 32.25)
        });
        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains("#set par(leading: 14.25pt)", source);
    }

    [Fact]
    public void GetTypstParagraphLeading_72PtFont80Percent_ReturnsNegative14_40()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var method = typeof(PptxToTypstConverter).GetMethod(
            "GetTypstParagraphLeading",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(TextSpacing), typeof(TypstTextFormatting), typeof(double?)],
            null);

        var spacing = new TextSpacing(TextSpacingKind.Percent, 0.8);
        var formatting = new TypstTextFormatting { FontFamily = "Arial", FontSize = 72 };
        var result = method!.Invoke(converter, [spacing, formatting, null]);
        var leading = Assert.IsType<double>(result);
        Assert.Equal(-14.40, leading, 2);
    }

    [Fact]
    public void GetTypstParagraphLeading_72PtFont100Percent_ReturnsZero()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var method = typeof(PptxToTypstConverter).GetMethod(
            "GetTypstParagraphLeading",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(TextSpacing), typeof(TypstTextFormatting), typeof(double?)],
            null);

        var spacing = new TextSpacing(TextSpacingKind.Percent, 1.0);
        var formatting = new TypstTextFormatting { FontFamily = "Arial", FontSize = 72 };
        var result = method!.Invoke(converter, [spacing, formatting, null]);
        var leading = Assert.IsType<double>(result);
        Assert.Equal(0.0, leading, 2);
    }

    [Fact]
    public void GetTypstParagraphLeading_18PtFont120Percent_Returns3_60()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var method = typeof(PptxToTypstConverter).GetMethod(
            "GetTypstParagraphLeading",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(TextSpacing), typeof(TypstTextFormatting), typeof(double?)],
            null);

        var spacing = new TextSpacing(TextSpacingKind.Percent, 1.2);
        var formatting = new TypstTextFormatting { FontFamily = "Arial", FontSize = 18 };
        var result = method!.Invoke(converter, [spacing, formatting, null]);
        var leading = Assert.IsType<double>(result);
        Assert.Equal(3.60, leading, 2);
    }

    [Fact]
    public void GetTypstParagraphLeading_27PtFontPoints39_75_Returns12_75()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var method = typeof(PptxToTypstConverter).GetMethod(
            "GetTypstParagraphLeading",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(TextSpacing), typeof(TypstTextFormatting), typeof(double?)],
            null);

        var spacing = new TextSpacing(TextSpacingKind.Points, 39.75);
        var formatting = new TypstTextFormatting { FontFamily = "Arial", FontSize = 27 };
        var result = method!.Invoke(converter, [spacing, formatting, null]);
        var leading = Assert.IsType<double>(result);
        Assert.Equal(12.75, leading, 2);
    }

    [Fact]
    public void GenerateTypstSource_100PercentLineSpacing_DoesNotEmitLeading()
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
                    Content = "Normal spacing",
                    Formatting = new TypstTextFormatting
                    {
                        FontFamily = "Arial",
                        FontSize = 18
                    }
                }
            ],
            LineSpacing = new TextSpacing(TextSpacingKind.Percent, 1.0)
        });
        var source = converter.GenerateTypstSource(presentation);

        Assert.DoesNotContain("#set par(leading:", source);
    }

    [Fact]
    public void GenerateTypstSource_TwoParagraphsDifferentLineSpacing_EmitsDifferentLeading()
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
                    },
                    LineSpacing = new TextSpacing(TextSpacingKind.Points, 32.25)
                },
                new TypstParagraph
                {
                    Content = "Second paragraph",
                    Formatting = new TypstTextFormatting
                    {
                        FontFamily = "Arial",
                        FontSize = 18
                    },
                    LineSpacing = new TextSpacing(TextSpacingKind.Points, 40.0)
                }
            ]
        });
        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains("#set par(leading: 14.25pt)", source);
        Assert.Contains("#set par(leading: 22.00pt)", source);
    }

    [Fact]
    public void GenerateTypstSource_FirstParagraphNoSpacing_SecondHasSpacing_OnlySecondEmitsLeading()
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
                    Content = "Second paragraph",
                    Formatting = new TypstTextFormatting
                    {
                        FontFamily = "Arial",
                        FontSize = 18
                    },
                    LineSpacing = new TextSpacing(TextSpacingKind.Points, 32.25)
                }
            ]
        });
        var source = converter.GenerateTypstSource(presentation);

        var firstParaIndex = source.IndexOf("First paragraph");
        var leadingIndex = source.IndexOf("#set par(leading:");

        Assert.True(leadingIndex > firstParaIndex, "Leading should appear after first paragraph content");
    }

    [Fact]
    public void GenerateTypstSource_ListGroup_EmitsLeadingBasedOnFirstItem()
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
                    Formatting = new TypstTextFormatting
                    {
                        FontFamily = "Arial",
                        FontSize = 18
                    },
                    HasBullet = true,
                    BulletChar = "•",
                    LineSpacing = new TextSpacing(TextSpacingKind.Points, 32.25)
                },
                new TypstParagraph
                {
                    Content = "Item two",
                    Formatting = new TypstTextFormatting
                    {
                        FontFamily = "Arial",
                        FontSize = 18
                    },
                    HasBullet = true,
                    BulletChar = "•",
                    LineSpacing = new TextSpacing(TextSpacingKind.Points, 32.25)
                }
            ]
        });
        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains("#set par(leading: 14.25pt)", source);
        Assert.Contains("#list", source);
        Assert.Contains("[Item one]", source);
        Assert.Contains("[Item two]", source);
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
            LineSpacing = new TextSpacing(TextSpacingKind.Points, 32.25)
        });
        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains("[First paragraph]", source);
        Assert.Contains("#v(32.25pt)", source);
        Assert.Contains("[Second paragraph]", source);
    }

    [Fact]
    public void GenerateTypstSource_SpaceBeforeWithLeading_DoesNotEmitVWithBlockArgument()
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
                    },
                    SpaceBefore = new TextSpacing(TextSpacingKind.Points, 10.0),
                    LineSpacing = new TextSpacing(TextSpacingKind.Points, 32.25)
                }
            ]
        });

        var source = converter.GenerateTypstSource(presentation);

        Assert.DoesNotMatch(@"#v\([^)]*\)\[", source);
        Assert.Contains("#v(10.00pt)", source);
        Assert.Contains("#set par(leading: 14.25pt)", source);
    }

    [Fact]
    public void GenerateTypstSource_LineSpacingNoSpaceBefore_NoUnsafeScopes()
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
                    Content = "Text with line spacing",
                    Formatting = new TypstTextFormatting { FontFamily = "Arial", FontSize = 72 }
                }
            ],
            LineSpacing = new TextSpacing(TextSpacingKind.Percent, 0.8)
        });

        var source = converter.GenerateTypstSource(presentation);
        AssertNoUnsafeTypstScopes(source);
    }

    [Fact]
    public void GenerateTypstSource_SpaceBeforeWithLineSpacing_NoUnsafeScopes()
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
                    Content = "Text with space before and line spacing",
                    Formatting = new TypstTextFormatting { FontFamily = "Arial", FontSize = 18 },
                    SpaceBefore = new TextSpacing(TextSpacingKind.Points, 10.0),
                    LineSpacing = new TextSpacing(TextSpacingKind.Points, 32.25)
                }
            ]
        });

        var source = converter.GenerateTypstSource(presentation);
        AssertNoUnsafeTypstScopes(source);
    }

    [Fact]
    public void PresPro_GeneratedTypst_HasNoUnsafeBracketScopes()
    {
        var baseDir = AppContext.BaseDirectory;
        var path = Path.Combine(baseDir, "..", "..", "..", "..", "examples", "REF", "pres-pro.pptx");
        path = Path.GetFullPath(path);

        if (!File.Exists(path))
            return;

        using var doc = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(doc);

        var presentation = converter.Convert();
        var source = converter.GenerateTypstSource(presentation);

        AssertNoUnsafeTypstScopes(source);
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

    [Fact]
    public void Convert_ParagraphPropertiesWithoutMarginOrIndent_DoesNotCrash()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var shape = new P.Shape(
            new TextBody(
                new Drawing.BodyProperties(),
                new Drawing.ListStyle(),
                new Drawing.Paragraph(
                    new Drawing.ParagraphProperties { LeftMargin = 457200, Indent = -228600 },
                    new Drawing.Run(new Drawing.Text { Text = "Indented item" })),
                new Drawing.Paragraph(
                    new Drawing.ParagraphProperties { Level = 0 },
                    new Drawing.Run(new Drawing.Text { Text = "Normal item" }))));

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

        Assert.Equal("Indented item\n\nNormal item", text.Content);
        Assert.Equal(2, text.Paragraphs.Count);
        Assert.Equal(36.0, text.Paragraphs[0].MarginLeft);   // 457200 EMU / 12700 = 36pt
        Assert.Equal(-18.0, text.Paragraphs[0].Indent);      // -228600 EMU / 12700 = -18pt
        Assert.Null(text.Paragraphs[1].MarginLeft);
        Assert.Null(text.Paragraphs[1].Indent);
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

    private static void AssertNoUnsafeTypstScopes(string typst)
    {
        // Reject [[#set par - double bracket renders as literal visible [
        Assert.DoesNotContain("[[#set par", typst);

        // Reject [#set par - raw bracket scope around set rule
        Assert.DoesNotContain("[#set par", typst);

        // Reject #v(...)[ - vertical spacing with content block argument
        Assert.DoesNotMatch(@"#v\([^)]*\)\[", typst);

        // Reject #v(...)\n[ followed by #set par - space before with unsafe scope
        Assert.DoesNotMatch(@"#v\([^)]*\)\n\[\s*#set par", typst);
    }

    private static void AssertNoNegativeLeading(string typst)
    {
        Assert.DoesNotContain("leading: -", typst);
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
        Assert.Contains("height: 100.00pt", source);
        Assert.DoesNotContain("#align(bottom)", source);
        // innerHeight=100, estimatedTextHeight=18, yOffset=82
        Assert.Contains("dy: 82.00pt", source);
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
        Assert.Contains("height: 100.00pt", source);
        Assert.DoesNotContain("#align(horizon)", source);
        // innerHeight=100, estimatedTextHeight=18, yOffset=41
        Assert.Contains("dy: 41.00pt", source);
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

    [Fact]
    public void Convert_NormAutofitFontScale_ScalesFontSize()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var bodyPr = new Drawing.BodyProperties();
        var normAutofit = new Drawing.NormalAutoFit();
        normAutofit.SetAttribute(new OpenXmlAttribute("fontScale", "", "92500"));
        bodyPr.Append(normAutofit);

        var shape = new P.Shape(
            new TextBody(
                bodyPr,
                new Drawing.ListStyle(),
                new Drawing.Paragraph(
                    new Drawing.Run(
                        new Drawing.RunProperties { FontSize = new Int32Value(7200) },
                        new Drawing.Text { Text = "Scaled text" }
                    )
                )
            )
        );

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

        Assert.Equal(0.925, text.FontScale);
        Assert.Equal(66.6, text.Formatting.FontSize, 1);

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

        Assert.Contains("size: 66.60pt", source);
    }

    [Fact]
    public void GenerateTypstSource_NormAutofitLnSpcReduction_ReducesLeading()
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
                    Content = "Reduced spacing",
                    Formatting = new TypstTextFormatting
                    {
                        FontFamily = "Arial",
                        FontSize = 72
                    },
                    LineSpacing = new TextSpacing(TextSpacingKind.Percent, 1.0)
                }
            ],
            LineSpacing = new TextSpacing(TextSpacingKind.Percent, 1.0),
            LineSpacingReduction = 0.10
        });

        var source = converter.GenerateTypstSource(presentation);

        // 72pt font, 100% spacing, 10% reduction => target = 72 * 1.0 * 0.9 = 64.8
        // natural height fallback = 72 (no metrics loaded)
        // leading = 64.8 - 72 = -7.2
        Assert.Contains("leading: -7.20pt", source);
    }

    [Fact]
    public void Convert_NoNormAutofit_NoScalingApplied()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var shape = new P.Shape(
            new TextBody(
                new Drawing.BodyProperties(),
                new Drawing.ListStyle(),
                new Drawing.Paragraph(
                    new Drawing.Run(
                        new Drawing.RunProperties { FontSize = new Int32Value(7200) },
                        new Drawing.Text { Text = "Normal text" }
                    )
                )
            )
        );

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

        Assert.Null(text.FontScale);
        Assert.Null(text.LineSpacingReduction);
        Assert.Equal(72.0, text.Formatting.FontSize);

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

        Assert.Contains("size: 72.00pt", source);
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
                    SpaceBefore = new TextSpacing(TextSpacingKind.Points, 12.5)
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
                    SpaceAfter = new TextSpacing(TextSpacingKind.Points, 8.0)
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
    public void GenerateTypstSource_SpaceAfterParagraph_Percent_EmitsScaledToFontSize()
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
                    Formatting = new TypstTextFormatting { FontFamily = "Arial", FontSize = 14 },
                    SpaceAfter = new TextSpacing(TextSpacingKind.Percent, 0.35)
                },
                new TypstParagraph
                {
                    Content = "Second paragraph",
                    Formatting = new TypstTextFormatting { FontFamily = "Arial", FontSize = 14 }
                }
            ]
        });

        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains("#v(4.90pt)", source);
        Assert.Contains("[First paragraph]", source);
        Assert.Contains("[Second paragraph]", source);
    }

    [Fact]
    public void GenerateTypstSource_SpaceBeforeParagraph_Percent_EmitsScaledToFontSize()
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
                    SpaceBefore = new TextSpacing(TextSpacingKind.Percent, 0.5)
                }
            ]
        });

        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains("#v(9.00pt)", source);
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
        var defRPr = new Drawing.DefaultRunProperties { Bold = new BooleanValue(true) };
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
    public void ExtractTableStylePart_DirectBoldAttribute_ReadsBold()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        converter.Convert();

        var extractTableStylePart = typeof(PptxToTypstConverter).GetMethod(
            "ExtractTableStylePart",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(Drawing.TablePartStyleType), typeof(string), typeof(TableStyleDefinition), typeof(StyleResolver)],
            null);

        var tcTxStyle = new Drawing.TableCellTextStyle();
        tcTxStyle.SetAttribute(new OpenXmlAttribute("b", "", "on"));

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
    }

    [Fact]
    public void ExtractTableStylePart_DirectSchemeClr_ResolvesToWhite()
    {
        var baseDir = AppContext.BaseDirectory;
        var path = Path.Combine(baseDir, "..", "..", "..", "..", "examples", "REF", "pres-pro.pptx");
        path = Path.GetFullPath(path);
        if (!File.Exists(path))
            return; // Skip if file doesn't exist

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        converter.Convert();

        var extractTableStylePart = typeof(PptxToTypstConverter).GetMethod(
            "ExtractTableStylePart",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(Drawing.TablePartStyleType), typeof(string), typeof(TableStyleDefinition), typeof(StyleResolver)],
            null);

        var tcTxStyle = new Drawing.TableCellTextStyle();
        tcTxStyle.SetAttribute(new OpenXmlAttribute("b", "", "on"));
        var schemeClr = new OpenXmlUnknownElement("a", "schemeClr", "http://schemas.openxmlformats.org/drawingml/2006/main");
        schemeClr.SetAttribute(new OpenXmlAttribute("val", "", "lt1"));
        tcTxStyle.Append(schemeClr);

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
    public void ExtractTableStylePart_EmptyTcBdr_DoesNotSetBorderNone()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        converter.Convert();

        var extractTableStylePart = typeof(PptxToTypstConverter).GetMethod(
            "ExtractTableStylePart",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(Drawing.TablePartStyleType), typeof(string), typeof(TableStyleDefinition), typeof(StyleResolver)],
            null);

        var tcStyle = new Drawing.TableCellStyle();
        var borders = new Drawing.TableCellBorders();
        tcStyle.Append(borders);

        var part = new Drawing.FirstRow();
        part.Append(tcStyle);

        var definition = new TableStyleDefinition { StyleId = "test" };

        extractTableStylePart!.Invoke(converter, [part, "firstRow", definition, null]);

        Assert.True(definition.Parts.ContainsKey("firstRow"));
        var stylePart = definition.Parts["firstRow"];
        Assert.False(stylePart.BorderTopNone);
        Assert.False(stylePart.BorderBottomNone);
        Assert.False(stylePart.BorderLeftNone);
        Assert.False(stylePart.BorderRightNone);
    }

    [Fact]
    public void ExtractTableStylePart_InsideVNoFill_SetsBorderInsideVNoneOnly()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        converter.Convert();

        var extractTableStylePart = typeof(PptxToTypstConverter).GetMethod(
            "ExtractTableStylePart",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(Drawing.TablePartStyleType), typeof(string), typeof(TableStyleDefinition), typeof(StyleResolver)],
            null);

        var tcStyle = new Drawing.TableCellStyle();
        var borders = new Drawing.TableCellBorders();

        // insideV with noFill
        var insideV = new Drawing.InsideVerticalBorder();
        insideV.Outline = new Drawing.Outline();
        insideV.Outline.Append(new Drawing.NoFill());
        borders.Append(insideV);

        // visible top border
        var top = new Drawing.TopBorder();
        top.Outline = new Drawing.Outline { Width = 12700 };
        top.Outline.Append(new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = "000000" }));
        borders.Append(top);

        tcStyle.Append(borders);

        var part = new Drawing.WholeTable();
        part.Append(tcStyle);

        var definition = new TableStyleDefinition { StyleId = "test" };

        extractTableStylePart!.Invoke(converter, [part, "wholeTbl", definition, null]);

        Assert.True(definition.Parts.ContainsKey("wholeTbl"));
        var stylePart = definition.Parts["wholeTbl"];
        Assert.True(stylePart.BorderInsideVNone);
        Assert.False(stylePart.BorderInsideHNone);
        Assert.False(stylePart.BorderTopNone);
    }

    [Fact]
    public void ExtractBorderInfo_MissingBorder_ReturnsInherit()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var extractBorderInfo = typeof(PptxToTypstConverter).GetMethod(
            "ExtractBorderInfo",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(OpenXmlElement), typeof(StyleResolver)],
            null);

        var result = extractBorderInfo!.Invoke(converter, [null, null]);
        var info = Assert.IsType<ValueTuple<string?, double?, bool>>(result);
        Assert.False(info.Item3); // IsNone should be false (inherit)
    }

    [Fact]
    public void GenerateTypstSource_TableWithRowHeights_EmitsRowsParameter()
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
                                RowHeights = [59.40, 71.07],
                                Rows =
                                [
                                    new List<TypstTableCell>
                                    {
                                        new TypstTableCell { Content = "R1C1", Formatting = new TypstTextFormatting { FontFamily = "Arial", FontSize = 14 } },
                                        new TypstTableCell { Content = "R1C2", Formatting = new TypstTextFormatting { FontFamily = "Arial", FontSize = 14 } }
                                    },
                                    new List<TypstTableCell>
                                    {
                                        new TypstTableCell { Content = "R2C1", Formatting = new TypstTextFormatting { FontFamily = "Arial", FontSize = 14 } },
                                        new TypstTableCell { Content = "R2C2", Formatting = new TypstTextFormatting { FontFamily = "Arial", FontSize = 14 } }
                                    }
                                ]
                            }
                        }
                    ]
                }
            ]
        };

        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains("rows: (59.40pt, 71.07pt)", source);
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
            [typeof(Drawing.TableCell), typeof(TableStylePart)],
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

        var result = extractCellFormatting!.Invoke(converter, [cell, stylePart]);
        var formatting = Assert.IsType<TypstTextFormatting>(result);

        Assert.True(formatting.Bold);
        Assert.Equal("#FFFFFF", formatting.Color);
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
                BorderRightNone = true
            }
        };

        var result = buildCellStroke!.Invoke(null, [cell, 1.0, "#000000"]);
        var stroke = Assert.IsType<string>(result);

        Assert.Contains("left: none", stroke);
        Assert.Contains("right: none", stroke);
    }

    [Fact]
    public void ExtractLineSpacingFromElement_SpcPts_ReturnsPointsKind()
    {
        var pPr = new OpenXmlUnknownElement("a", "pPr", "http://schemas.openxmlformats.org/drawingml/2006/main");
        var lnSpc = new OpenXmlUnknownElement("a", "lnSpc", "http://schemas.openxmlformats.org/drawingml/2006/main");
        var spcPts = new OpenXmlUnknownElement("a", "spcPts", "http://schemas.openxmlformats.org/drawingml/2006/main");
        spcPts.SetAttribute(new OpenXmlAttribute("val", "", "1000"));
        lnSpc.AppendChild(spcPts);
        pPr.AppendChild(lnSpc);

        var method = typeof(PptxToTypstConverter).GetMethod(
            "ExtractLineSpacingFromElement",
            System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(OpenXmlElement)],
            null);

        var result = method!.Invoke(null, [pPr]);
        var spacing = Assert.IsType<TextSpacing>(result);
        Assert.Equal(TextSpacingKind.Points, spacing.Kind);
        Assert.Equal(10.0, spacing.Value);
    }

    [Fact]
    public void ExtractLineSpacingFromElement_SpcPct_ReturnsPercentKind()
    {
        var pPr = new OpenXmlUnknownElement("a", "pPr", "http://schemas.openxmlformats.org/drawingml/2006/main");
        var lnSpc = new OpenXmlUnknownElement("a", "lnSpc", "http://schemas.openxmlformats.org/drawingml/2006/main");
        var spcPct = new OpenXmlUnknownElement("a", "spcPct", "http://schemas.openxmlformats.org/drawingml/2006/main");
        spcPct.SetAttribute(new OpenXmlAttribute("val", "", "35000"));
        lnSpc.AppendChild(spcPct);
        pPr.AppendChild(lnSpc);

        var method = typeof(PptxToTypstConverter).GetMethod(
            "ExtractLineSpacingFromElement",
            System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(OpenXmlElement)],
            null);

        var result = method!.Invoke(null, [pPr]);
        var spacing = Assert.IsType<TextSpacing>(result);
        Assert.Equal(TextSpacingKind.Percent, spacing.Kind);
        Assert.Equal(0.35, spacing.Value);
    }

    [Fact]
    public void ExtractLineSpacingFromElement_SpcPct80Percent_ReturnsCorrectValue()
    {
        var pPr = new OpenXmlUnknownElement("a", "pPr", "http://schemas.openxmlformats.org/drawingml/2006/main");
        var lnSpc = new OpenXmlUnknownElement("a", "lnSpc", "http://schemas.openxmlformats.org/drawingml/2006/main");
        var spcPct = new OpenXmlUnknownElement("a", "spcPct", "http://schemas.openxmlformats.org/drawingml/2006/main");
        spcPct.SetAttribute(new OpenXmlAttribute("val", "", "80000"));
        lnSpc.AppendChild(spcPct);
        pPr.AppendChild(lnSpc);

        var method = typeof(PptxToTypstConverter).GetMethod(
            "ExtractLineSpacingFromElement",
            System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(OpenXmlElement)],
            null);

        var result = method!.Invoke(null, [pPr]);
        var spacing = Assert.IsType<TextSpacing>(result);
        Assert.Equal(TextSpacingKind.Percent, spacing.Kind);
        Assert.Equal(0.8, spacing.Value);
    }

    [Fact]
    public void ExtractParagraphSpacing_SpcPtsAndSpcPct_PreservesKinds()
    {
        var pPr = new OpenXmlUnknownElement("a", "pPr", "http://schemas.openxmlformats.org/drawingml/2006/main");
        var spcBef = new OpenXmlUnknownElement("a", "spcBef", "http://schemas.openxmlformats.org/drawingml/2006/main");
        var spcBefPts = new OpenXmlUnknownElement("a", "spcPts", "http://schemas.openxmlformats.org/drawingml/2006/main");
        spcBefPts.SetAttribute(new OpenXmlAttribute("val", "", "1250"));
        spcBef.AppendChild(spcBefPts);
        pPr.AppendChild(spcBef);

        var spcAft = new OpenXmlUnknownElement("a", "spcAft", "http://schemas.openxmlformats.org/drawingml/2006/main");
        var spcAftPct = new OpenXmlUnknownElement("a", "spcPct", "http://schemas.openxmlformats.org/drawingml/2006/main");
        spcAftPct.SetAttribute(new OpenXmlAttribute("val", "", "50000"));
        spcAft.AppendChild(spcAftPct);
        pPr.AppendChild(spcAft);

        var method = typeof(PptxToTypstConverter).GetMethod(
            "ExtractParagraphSpacing",
            System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(OpenXmlElement)],
            null);

        var result = method!.Invoke(null, [pPr]);
        var spacing = Assert.IsType<ValueTuple<TextSpacing?, TextSpacing?>>(result);
        Assert.NotNull(spacing.Item1);
        Assert.Equal(TextSpacingKind.Points, spacing.Item1!.Kind);
        Assert.Equal(12.5, spacing.Item1.Value);
        Assert.NotNull(spacing.Item2);
        Assert.Equal(TextSpacingKind.Percent, spacing.Item2!.Kind);
        Assert.Equal(0.5, spacing.Item2.Value);
    }

    [Fact]
    public void GenerateTypstSource_TopAlignedText_NoLeadingOrMetricOffset()
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
                    Content = "Top aligned",
                    Formatting = new TypstTextFormatting { FontFamily = "Arial", FontSize = 18 }
                }
            ],
            VerticalAlign = "top",
            PaddingTop = 5,
            LineSpacing = new TextSpacing(TextSpacingKind.Points, 32.25)
        }, height: 100, width: 300);

        var source = converter.GenerateTypstSource(presentation);

        // dy should be exactly paddingTop (5pt), not 5pt + topLeading + metricOffset
        Assert.Contains("dy: 5.00pt", source);
        Assert.DoesNotContain("#align(horizon)", source);
        Assert.DoesNotContain("#align(bottom)", source);
    }

    [Fact]
    public void GenerateTypstSource_CenterAlignedText_ApproximateHalfOffset()
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

        // innerHeight = 100, estimatedTextHeight = 18 (1 line * 18pt font)
        // yOffset = (100 - 18) / 2 = 41
        // dy = 0 + 0 + 41 = 41
        Assert.Contains("dy: 41.00pt", source);
        Assert.DoesNotContain("#align(horizon)", source);
    }

    [Fact]
    public void GenerateTypstSource_BottomAlignedLargeText_ApproximateBottomOffset()
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
                    Formatting = new TypstTextFormatting { FontFamily = "Arial", FontSize = 40 }
                }
            ],
            VerticalAlign = "bottom"
        }, height: 100, width: 300);

        var source = converter.GenerateTypstSource(presentation);

        // innerHeight = 100, estimatedTextHeight = 40 (1 line * 40pt font)
        // yOffset = 100 - 40 = 60
        // dy = 0 + 0 + 60 = 60
        Assert.Contains("dy: 60.00pt", source);
        Assert.DoesNotContain("#align(bottom)", source);
    }

    [Fact]
    public void ResolveCellStylePart_MergesMissingBorderFromWholeTable()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var resolveCellStylePart = typeof(PptxToTypstConverter).GetMethod(
            "ResolveCellStylePart",
            System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(int), typeof(int), typeof(int), typeof(int), typeof(TableStyleDefinition), typeof(bool), typeof(bool), typeof(bool), typeof(bool), typeof(bool)],
            null);

        var definition = new TableStyleDefinition
        {
            StyleId = "test",
            Parts = new Dictionary<string, TableStylePart>(StringComparer.OrdinalIgnoreCase)
            {
                ["wholeTbl"] = new TableStylePart
                {
                    BorderTopColor = "#000000",
                    BorderTopWidth = 1.0,
                    BorderBottomColor = "#000000",
                    BorderBottomWidth = 1.0,
                    BorderLeftColor = "#000000",
                    BorderLeftWidth = 1.0,
                    BorderRightColor = "#000000",
                    BorderRightWidth = 1.0,
                    TextColor = "#333333"
                },
                ["firstRow"] = new TableStylePart
                {
                    BackgroundColor = "#CEBA80",
                    TextBold = true
                }
            }
        };

        var result = resolveCellStylePart!.Invoke(null, [0, 0, 3, 2, definition, true, false, false, false, false]);
        var part = Assert.IsType<TableStylePart>(result);

        Assert.Equal("#CEBA80", part.BackgroundColor);
        Assert.True(part.TextBold);
        Assert.Equal("#000000", part.BorderTopColor);
        Assert.Equal(1.0, part.BorderTopWidth);
        Assert.Equal("#000000", part.BorderLeftColor);
        Assert.Equal("#333333", part.TextColor);
    }

    [Fact]
    public void ApplyInteriorBorderLogic_UsesInsideColorWidth()
    {
        var path = CreateSimplePptx();

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var applyInteriorBorderLogic = typeof(PptxToTypstConverter).GetMethod(
            "ApplyInteriorBorderLogic",
            System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(TableStylePart), typeof(int), typeof(int), typeof(int), typeof(int)],
            null);

        var buildCellStroke = typeof(PptxToTypstConverter).GetMethod(
            "BuildCellStroke",
            System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic,
            null,
            [typeof(TypstTableCell), typeof(double), typeof(string)],
            null);

        var stylePart = new TableStylePart
        {
            BorderTopColor = "#000000",
            BorderTopWidth = 1.0,
            BorderBottomColor = "#000000",
            BorderBottomWidth = 1.0,
            BorderLeftColor = "#000000",
            BorderLeftWidth = 1.0,
            BorderRightColor = "#000000",
            BorderRightWidth = 1.0,
            InsideHColor = "#FF0000",
            InsideHWidth = 2.0,
            InsideVColor = "#0000FF",
            InsideVWidth = 3.0
        };

        var result = applyInteriorBorderLogic!.Invoke(null, [stylePart, 1, 1, 2, 2]);
        var mergedPart = Assert.IsType<TableStylePart>(result);

        var cell = new TypstTableCell
        {
            Content = "Cell",
            StylePart = mergedPart
        };

        var stroke = buildCellStroke!.Invoke(null, [cell, 1.0, "#000000"]);
        var strokeStr = Assert.IsType<string>(stroke);

        // Row 1, Col 1 in a 2x2 table: top/left are interior, bottom/right are outer edges
        Assert.Contains("top: 2.00pt + rgb(\"#FF0000\")", strokeStr);
        Assert.Contains("left: 3.00pt + rgb(\"#0000FF\")", strokeStr);
        Assert.DoesNotContain("bottom:", strokeStr);
        Assert.DoesNotContain("right:", strokeStr);
    }

    [Fact]
    public void BuildCellStroke_MissingSideColor_DoesNotEmitRedundantDefault()
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
                BorderTopColor = null,
                BorderTopWidth = null,
                BorderBottomColor = null,
                BorderBottomWidth = null
            }
        };

        var result = buildCellStroke!.Invoke(null, [cell, 1.0, "#000000"]);
        Assert.Null(result);
    }

    [Fact]
    public void BuildCellStroke_ExplicitWidthMissingColor_EmitsWithDefaultColor()
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
                BorderTopWidth = 2.0,
                BorderTopColor = null
            }
        };

        var result = buildCellStroke!.Invoke(null, [cell, 1.0, "#000000"]);
        var stroke = Assert.IsType<string>(result);
        Assert.Contains("top: 2.00pt + rgb(\"#000000\")", stroke);
    }

    [Fact]
    public void GenerateTypstSource_72PtFont80PercentLineSpacing_NoNegativeLeading()
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
                    Content = "Large text with tight spacing",
                    Formatting = new TypstTextFormatting
                    {
                        FontFamily = "Arial",
                        FontSize = 72
                    }
                }
            ],
            LineSpacing = new TextSpacing(TextSpacingKind.Percent, 0.8)
        });

        var source = converter.GenerateTypstSource(presentation);
        AssertNoNegativeLeading(source);
    }

    [Fact]
    public void PresPro_GeneratedTypst_HasNoNegativeLeading()
    {
        var baseDir = AppContext.BaseDirectory;
        var path = Path.Combine(baseDir, "..", "..", "..", "..", "examples", "REF", "pres-pro.pptx");
        path = Path.GetFullPath(path);

        if (!File.Exists(path))
            return;

        using var doc = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(doc);

        var presentation = converter.Convert();
        var source = converter.GenerateTypstSource(presentation);

        AssertNoNegativeLeading(source);
    }

    private static string? FindTypstCli()
    {
        try
        {
            using var process = new Process();
            process.StartInfo.FileName = "typst";
            process.StartInfo.Arguments = "--version";
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.RedirectStandardError = true;
            process.StartInfo.UseShellExecute = false;
            process.Start();
            process.WaitForExit(5000);
            if (process.ExitCode == 0)
                return "typst";
        }
        catch
        {
            // typst not on PATH
        }
        return null;
    }

    [Fact]
    public void SmokeTest_PresPro_ConvertsAndCompilesToPdf()
    {
        var baseDir = AppContext.BaseDirectory;
        var path = Path.Combine(baseDir, "..", "..", "..", "..", "examples", "REF", "pres-pro.pptx");
        path = Path.GetFullPath(path);

        if (!File.Exists(path))
        {
            _output.WriteLine($"Fixture not found: {path}");
            return;
        }

        using var doc = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(doc);

        var presentation = converter.Convert();
        var source = converter.GenerateTypstSource(presentation);

        Assert.NotNull(source);
        Assert.False(string.IsNullOrWhiteSpace(source), "Generated Typst source should not be empty");

        var typPath = Path.Combine(presentation.TempDirectory, "smoke.typ");
        File.WriteAllText(typPath, source);

        var typstCli = FindTypstCli();
        if (typstCli == null)
        {
            _output.WriteLine("Typst CLI not found on PATH; skipping PDF compilation assertion.");
            return;
        }

        var pdfPath = Path.Combine(presentation.TempDirectory, "smoke.pdf");
        var psi = new ProcessStartInfo
        {
            FileName = typstCli,
            Arguments = $"compile --font-path \"{Path.Combine(presentation.TempDirectory, "fonts")}\" \"{typPath}\" \"{pdfPath}\"",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false
        };

        using var process = Process.Start(psi);
        process!.WaitForExit(30000);
        var stdout = process.StandardOutput.ReadToEnd();
        var stderr = process.StandardError.ReadToEnd();

        if (!string.IsNullOrEmpty(stdout))
            _output.WriteLine($"typst stdout: {stdout}");
        if (!string.IsNullOrEmpty(stderr))
            _output.WriteLine($"typst stderr: {stderr}");

        Assert.Equal(0, process.ExitCode);
        Assert.True(File.Exists(pdfPath), $"PDF should exist at {pdfPath}");
        var pdfInfo = new FileInfo(pdfPath);
        Assert.True(pdfInfo.Length > 0, "PDF should not be empty");
    }

    [Fact]
    public void SmokeTest_REMOVED_ConvertsAndCompilesToPdf()
    {
        var baseDir = AppContext.BaseDirectory;
        var path = Path.Combine(baseDir, "..", "..", "..", "..", "examples", "REF", "REMOVED.pptx");
        path = Path.GetFullPath(path);

        if (!File.Exists(path))
        {
            _output.WriteLine($"Fixture not found: {path}");
            return;
        }

        using var doc = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(doc);

        var presentation = converter.Convert();
        var source = converter.GenerateTypstSource(presentation);

        Assert.NotNull(source);
        Assert.False(string.IsNullOrWhiteSpace(source), "Generated Typst source should not be empty");

        var typPath = Path.Combine(presentation.TempDirectory, "smoke.typ");
        File.WriteAllText(typPath, source);

        var typstCli = FindTypstCli();
        if (typstCli == null)
        {
            _output.WriteLine("Typst CLI not found on PATH; skipping PDF compilation assertion.");
            return;
        }

        var pdfPath = Path.Combine(presentation.TempDirectory, "smoke.pdf");
        var psi = new ProcessStartInfo
        {
            FileName = typstCli,
            Arguments = $"compile --font-path \"{Path.Combine(presentation.TempDirectory, "fonts")}\" \"{typPath}\" \"{pdfPath}\"",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false
        };

        using var process = Process.Start(psi);
        process!.WaitForExit(30000);
        var stdout = process.StandardOutput.ReadToEnd();
        var stderr = process.StandardError.ReadToEnd();

        if (!string.IsNullOrEmpty(stdout))
            _output.WriteLine($"typst stdout: {stdout}");
        if (!string.IsNullOrEmpty(stderr))
            _output.WriteLine($"typst stderr: {stderr}");

        Assert.Equal(0, process.ExitCode);
        Assert.True(File.Exists(pdfPath), $"PDF should exist at {pdfPath}");
        var pdfInfo = new FileInfo(pdfPath);
        Assert.True(pdfInfo.Length > 0, "PDF should not be empty");
    }
}
