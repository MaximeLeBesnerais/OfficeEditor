using DocumentFormat.OpenXml.Packaging;
using PptxEditor.Core.Converters;
using PptxEditor.Core.Models;
using Xunit;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Gradient stops from OOXML <c>a:gradFill</c> may sit strictly inside [0, 1]
/// (e.g. 29%..73%) — PowerPoint flat-extends the edge colors to the shape
/// bounds. Typst instead requires the first stop at 0% and the last at 100%
/// ("first stop must have an offset of 0"), so the converter must pad the
/// edges with the boundary colors.
/// </summary>
public sealed class PptxToTypstConverterGradientTests
{
    private static string ReferencePath(string fileName) =>
        Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "examples", "REF", "PPTX", fileName);

    [Fact]
    public void GenerateTypstSource_InteriorGradientStops_PadsEdgesToFullRange()
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
                            Type = "Shape",
                            Width = 100,
                            Height = 50,
                            Shape = new TypstShapeElement
                            {
                                ShapeType = "rect",
                                FillGradient = new TypstGradientFill(90,
                                [
                                    new TypstGradientStop("#95A5A6", 0.29),
                                    new TypstGradientStop("#BAC4C5", 0.73)
                                ])
                            }
                        }
                    ]
                }
            ]
        };

        var source = converter.GenerateTypstSource(presentation);

        // Edge colors flat-extended to 0%/100% (PowerPoint semantics), interior
        // stops preserved.
        Assert.Contains(
            "gradient.linear((rgb(\"#95A5A6\"), 0%), (rgb(\"#95A5A6\"), 29%), (rgb(\"#BAC4C5\"), 73%), (rgb(\"#BAC4C5\"), 100%), angle: 90deg)",
            source);
    }

    [Fact]
    public void GenerateTypstSource_FullRangeGradientStops_EmittedUnchanged()
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
                            Type = "Shape",
                            Width = 100,
                            Height = 50,
                            Shape = new TypstShapeElement
                            {
                                ShapeType = "rect",
                                FillGradient = new TypstGradientFill(45,
                                [
                                    new TypstGradientStop("#000000", 0),
                                    new TypstGradientStop("#FFFFFF", 1)
                                ])
                            }
                        }
                    ]
                }
            ]
        };

        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains(
            "gradient.linear((rgb(\"#000000\"), 0%), (rgb(\"#FFFFFF\"), 100%), angle: 45deg)",
            source);
    }
}
