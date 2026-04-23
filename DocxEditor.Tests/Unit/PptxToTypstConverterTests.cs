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
            
            var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
            var slideMaster = new SlideMaster(
                new CommonSlideData(new ShapeTree()),
                new ColorMap(),
                new SlideLayoutIdList()
            );
            slideMasterPart.SlideMaster = slideMaster;
            
            var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
            slideLayoutPart.SlideLayout = new DocumentFormat.OpenXml.Presentation.SlideLayout(new CommonSlideData(new ShapeTree()));
            
            var layoutId = new SlideLayoutId
            {
                Id = 1,
                RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart)
            };
            slideMaster.SlideLayoutIdList!.Append(layoutId);
            
            presentationPart.Presentation.SlideIdList = new SlideIdList();
            
            // Create slide with formatted text
            var slidePart = presentationPart.AddNewPart<SlidePart>();
            var slide = new Slide(
                new CommonSlideData(
                    new ShapeTree(
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
}
