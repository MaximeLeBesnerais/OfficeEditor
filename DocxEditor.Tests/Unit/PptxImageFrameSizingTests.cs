using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using PptxEditor.Core.Converters;
using PptxEditor.Core.Models;
using Xunit;

namespace DocxEditor.Tests.Unit;

public sealed class PptxImageFrameSizingTests
{
    private static string ReferencePath(string fileName) =>
        Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "examples", "REF", "PPTX", fileName);

    [Fact]
    public void GenerateTypstSource_LowResolutionImage_UsesFrameDimensions()
    {
        using var document = PresentationDocument.Open(ReferencePath("northwind-demo.pptx"), false);
        using var converter = new PptxToTypstConverter(document);
        var presentation = new TypstPresentation
        {
            Slides =
            [
                new TypstSlide
                {
                    Layout = new SlideLayout { Width = 720, Height = 540 },
                    Elements =
                    [
                        new TypstElement
                        {
                            Type = "Image",
                            Width = 600,
                            Height = 400,
                            Image = new TypstImageElement
                            {
                                FileName = "low-resolution.png",
                                PixelWidth = 10,
                                PixelHeight = 10
                            }
                        }
                    ]
                }
            ]
        };

        var source = converter.GenerateTypstSource(presentation);
        var size = EmittedImageSize(source, "low-resolution.png");

        Assert.Equal(600, size.Width, 2);
        Assert.Equal(400, size.Height, 2);
    }

    [Fact]
    public void Convert_AetherLinkSlide1BackgroundImage_UsesFullSlideFrame()
    {
        var (presentation, source) = ConvertReference("AetherLink-Glass-Shareholder-Overview.pptx");
        var slide = presentation.Slides[0];
        var imageElement = Assert.Single(slide.Elements,
            element => element.Type == "Image" && element.X == 0 && element.Y == 0);

        Assert.InRange(imageElement.Width, slide.Layout.Width - 2, slide.Layout.Width + 2);
        Assert.InRange(imageElement.Height, slide.Layout.Height - 2, slide.Layout.Height + 2);
        AssertImageUsesElementFrame(source, imageElement);
    }

    [Fact]
    public void Convert_LaunchReview_SolidBackgroundStillEmitsPageFill()
    {
        // northwind-launch-review.pptx slide 1 carries a slide-level solid background
        // (#0B1F3A); the page fill must survive into the emitted Typst source.
        var (_, source) = ConvertReference("northwind-launch-review.pptx");

        Assert.Contains("#set page(fill: rgb(\"#0B1F3A\"))", source);
    }

    [Fact]
    public void Convert_NativePixelMetadataDoesNotOverrideImageFrameGeometry()
    {
        var (presentation, source) = ConvertReference("AetherLink-Glass-Shareholder-Overview.pptx");
        var imageElement = Assert.Single(presentation.Slides[0].Elements,
            element => element.Type == "Image" && element.X == 0 && element.Y == 0);
        Assert.NotNull(imageElement.Image);
        var image = imageElement.Image!;

        Assert.NotNull(image.PixelWidth);
        Assert.NotNull(image.PixelHeight);
        Assert.True(image.PixelWidth!.Value * 72.0 / 150 < imageElement.Width);
        Assert.True(image.PixelHeight!.Value * 72.0 / 150 < imageElement.Height);
        AssertImageUsesElementFrame(source, imageElement);
    }

    private static (TypstPresentation Presentation, string Source) ConvertReference(string fileName)
    {
        using var document = PresentationDocument.Open(ReferencePath(fileName), false);
        using var converter = new PptxToTypstConverter(document);
        var presentation = converter.Convert();
        return (presentation, converter.GenerateTypstSource(presentation));
    }

    private static void AssertImageUsesElementFrame(string source, TypstElement element)
    {
        Assert.NotNull(element.Image);
        var size = EmittedImageSize(source, element.Image!.FileName);

        Assert.Equal(element.Width, size.Width, 2);
        Assert.Equal(element.Height, size.Height, 2);
    }

    private static (double Width, double Height) EmittedImageSize(string source, string fileName)
    {
        var pattern = $"#image\\(\\\"assets/{Regex.Escape(fileName)}\\\", width: (?<width>[0-9.]+)pt, height: (?<height>[0-9.]+)pt\\)";
        var match = Regex.Match(source, pattern);
        Assert.True(match.Success, $"Image call for '{fileName}' was not found in generated source.");

        return (
            double.Parse(match.Groups["width"].Value, System.Globalization.CultureInfo.InvariantCulture),
            double.Parse(match.Groups["height"].Value, System.Globalization.CultureInfo.InvariantCulture));
    }
}
