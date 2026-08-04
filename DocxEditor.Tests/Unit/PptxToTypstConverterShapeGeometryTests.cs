using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PptxEditor.Core.Converters;
using PptxEditor.Core.Models;
using Xunit;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Tests for preset-geometry shapes on the regular (non-SmartArt) shape path of
/// <see cref="PptxToTypstConverter"/>: chevron/diamond presets must emit Typst
/// polygons (Typst has no native chevron/diamond shape, and the previous fallthrough
/// degraded them to plain rects), and <c>a:xfrm rot</c> must survive into the
/// emitted <c>#rotate</c> wrapper.
/// </summary>
public sealed class PptxToTypstConverterShapeGeometryTests : IDisposable
{
    private readonly string _tempDir;

    public PptxToTypstConverterShapeGeometryTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), nameof(PptxToTypstConverterShapeGeometryTests), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        SlideOpsTestHelpers.BestEffortDelete(_tempDir);
    }

    [Fact]
    public void Convert_ChevronPreset_EmitsPolygonWithAspectAwarePointDepth()
    {
        // 200pt x 100pt chevron with default adj (50000 = 50% of the SMALLER side):
        // point depth = 50pt → normalised x-offset 0.25.
        var path = CreateDeck(PresetShape(2, Drawing.ShapeTypeValues.Chevron, cx: 2540000, cy: 1270000));

        var (shape, source) = ConvertSingleShape(path);

        Assert.Equal("polygon", shape.ShapeType);
        Assert.Equal(6, shape.Points.Count);
        Assert.Equal((0.0, 0.0), shape.Points[0]);
        Assert.Equal((0.75, 0.0), shape.Points[1]);
        Assert.Equal((1.0, 0.5), shape.Points[2]);
        Assert.Equal((0.75, 1.0), shape.Points[3]);
        Assert.Equal((0.0, 1.0), shape.Points[4]);
        Assert.Equal((0.25, 0.5), shape.Points[5]);

        // Pin the emitted Typst polygon (points scaled to the element box).
        Assert.Contains("#polygon(", source);
        Assert.Contains("(0.00pt, 0.00pt)", source);
        Assert.Contains("(150.00pt, 0.00pt)", source);
        Assert.Contains("(200.00pt, 50.00pt)", source);
        Assert.Contains("(50.00pt, 50.00pt)", source);
        Assert.Contains("#C00000", source);
    }

    [Fact]
    public void Convert_ChevronPresetWithExplicitAdj_HonorsAdjustmentValue()
    {
        // adj 25000 → depth = 25% of min(w,h) = 25pt → normalised x-offset 0.125.
        var path = CreateDeck(PresetShape(2, Drawing.ShapeTypeValues.Chevron, cx: 2540000, cy: 1270000, adj: 25000));

        var (shape, _) = ConvertSingleShape(path);

        Assert.Equal("polygon", shape.ShapeType);
        Assert.Equal((0.875, 0.0), shape.Points[1]);
        Assert.Equal((0.125, 0.5), shape.Points[5]);
    }

    [Fact]
    public void Convert_DiamondPreset_EmitsDiamondPolygon()
    {
        var path = CreateDeck(PresetShape(2, Drawing.ShapeTypeValues.Diamond, cx: 381000, cy: 381000));

        var (shape, source) = ConvertSingleShape(path);

        Assert.Equal("polygon", shape.ShapeType);
        Assert.Equal(4, shape.Points.Count);
        Assert.Equal((0.5, 0.0), shape.Points[0]);
        Assert.Equal((1.0, 0.5), shape.Points[1]);
        Assert.Equal((0.5, 1.0), shape.Points[2]);
        Assert.Equal((0.0, 0.5), shape.Points[3]);
        Assert.Contains("#polygon(", source);
    }

    [Fact]
    public void Convert_RotatedRect_EmitsRotateWrapper()
    {
        // rot is in 60000ths of a degree: 2700000 = 45°.
        var path = CreateDeck(PresetShape(2, Drawing.ShapeTypeValues.Rectangle, cx: 1270000, cy: 1270000, rot: 2700000));

        var (shape, source) = ConvertSingleShape(path);

        Assert.Equal("rect", shape.ShapeType);
        Assert.Contains("#rotate(45.0deg", source);
    }

    [Fact]
    public void GenerateTypstSource_PolygonWithoutFillOrStroke_EmitsValidPolygonSyntax()
    {
        // Regression: a fill-less, stroke-less polygon emitted "#polygon(, (x, y), …)"
        // — a leading empty argument that Typst rejects with "unexpected comma".
        // Triggered by SmartArt diagram shapes whose fills don't resolve to a solid
        // color (e.g. dsp gradient fills the extractor does not parse).
        var path = CreateDeck(PresetShape(2, Drawing.ShapeTypeValues.Chevron, cx: 2540000, cy: 1270000));
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = new TypstPresentation
        {
            Slides =
            {
                new TypstSlide
                {
                    SlideIndex = 1,
                    Elements =
                    {
                        new TypstElement
                        {
                            Type = "Shape", X = 100, Y = 50, Width = 44, Height = 9,
                            Shape = new TypstShapeElement
                            {
                                ShapeType = "polygon",
                                Points = { (0, 0), (0.5, 0), (1, 0.5), (0.5, 1), (0, 1), (0.5, 0.5) }
                            }
                        }
                    }
                }
            }
        };

        var source = converter.GenerateTypstSource(presentation);

        Assert.DoesNotContain("#polygon(,", source);
        Assert.Contains("#polygon(fill: none, ", source);
    }

    [Fact]
    public void GenerateTypstSource_InvisibleLine_OmitsInvalidNoStrokeDecorator()
    {
        var path = CreateDeck();
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = new TypstPresentation
        {
            Slides =
            {
                new TypstSlide
                {
                    Layout = new PptxEditor.Core.Models.SlideLayout { Width = 200, Height = 100 },
                    Elements =
                    {
                        new TypstElement
                        {
                            Type = "Shape", Width = 100, Height = 50,
                            Shape = new TypstShapeElement
                            {
                                ShapeType = "line",
                                NoStroke = true,
                                Points = { (0, 0), (1, 1) }
                            }
                        }
                    }
                }
            }
        };

        var source = converter.GenerateTypstSource(presentation);

        Assert.DoesNotContain("#line(", source);
        Assert.DoesNotContain("stroke: none", source);
    }

    [Fact]
    public void GenerateTypstSource_LineShape_PlacesArrowheadAtEndpoint()
    {
        var path = CreateDeck();
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = new TypstPresentation
        {
            Slides =
            {
                new TypstSlide
                {
                    Layout = new PptxEditor.Core.Models.SlideLayout { Width = 200, Height = 100 },
                    Elements =
                    {
                        new TypstElement
                        {
                            Type = "Shape", Width = 100, Height = 50,
                            Shape = new TypstShapeElement
                            {
                                ShapeType = "line",
                                StrokeColor = "#123456",
                                StrokeWidth = 2,
                                Points = { (0, 0.5), (1, 0.5) },
                                ArrowAtEnd = true
                            }
                        }
                    }
                }
            }
        };

        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains("#place(dx: 100.00pt, dy: 25.00pt)[#polygon(fill: rgb(\"#123456\"), (0pt, 0pt)", source);
    }

    [Fact]
    public void GenerateTypstSource_CustomPath_EmitsEveryContourAndClosure()
    {
        var path = CreateDeck();
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = new TypstPresentation
        {
            Slides =
            {
                new TypstSlide
                {
                    Layout = new PptxEditor.Core.Models.SlideLayout { Width = 200, Height = 100 },
                    Elements =
                    {
                        new TypstElement
                        {
                            Type = "Shape", Width = 100, Height = 50,
                            Shape = new TypstShapeElement
                            {
                                ShapeType = "path",
                                FillColor = "#123456",
                                NoStroke = true,
                                Subpaths =
                                {
                                    new() { (0, 0), (1, 0), (1, 1) },
                                    new() { (0.25, 0.25), (0.75, 0.25) }
                                },
                                ClosedSubpaths = { true, false }
                            }
                        }
                    }
                }
            }
        };

        var source = converter.GenerateTypstSource(presentation);

        Assert.Contains("curve.move((0.00pt, 0.00pt))", source);
        Assert.Contains("curve.move((25.00pt, 12.50pt))", source);
        Assert.Contains("curve.close(mode: \"straight\")", source);
        Assert.Contains("#curve(fill: rgb(\"#123456\")", source);
        Assert.DoesNotContain("stroke: none", source);
    }

    [Fact]
    public void Convert_SystemColorLastClr_UsesValidatedWhiteFallback()
    {
        var shape = PresetShape(2, Drawing.ShapeTypeValues.Rectangle, 1270000, 1270000);
        var fill = shape.ShapeProperties!.Elements<Drawing.SolidFill>().Single();
        fill.Remove();
        shape.ShapeProperties.AppendChild(new Drawing.SolidFill(new Drawing.SystemColor
        {
            Val = Drawing.SystemColorValues.Window,
            LastColor = "FFFFFF"
        }));

        var (converted, _) = ConvertSingleShape(CreateDeck(shape));

        Assert.Equal("#FFFFFF", converted.FillColor);
    }

    [Fact]
    public void Convert_InvalidSystemColorLastClr_DoesNotBecomeBlackFill()
    {
        var shape = PresetShape(2, Drawing.ShapeTypeValues.Rectangle, 1270000, 1270000);
        var fill = shape.ShapeProperties!.Elements<Drawing.SolidFill>().Single();
        fill.Remove();
        shape.ShapeProperties.AppendChild(new Drawing.SolidFill(new Drawing.SystemColor
        {
            Val = Drawing.SystemColorValues.Window,
            LastColor = "FFFF"
        }));

        using var document = PresentationDocument.Open(CreateDeck(shape), false);
        using var converter = new PptxToTypstConverter(document);

        Assert.Empty(converter.Convert().Slides[0].Elements);
    }

    [Fact]
    public void Convert_PresetWhiteFill_ResolvesToWhite()
    {
        var shape = PresetShape(2, Drawing.ShapeTypeValues.Rectangle, 1270000, 1270000);
        var fill = shape.ShapeProperties!.Elements<Drawing.SolidFill>().Single();
        fill.Remove();
        shape.ShapeProperties.AppendChild(new Drawing.SolidFill(
            new Drawing.PresetColor { Val = Drawing.PresetColorValues.White }));

        var (converted, _) = ConvertSingleShape(CreateDeck(shape));

        Assert.Equal("#FFFFFF", converted.FillColor);
    }

    [Fact]
    public void Convert_PlainRectangleWithEmptyAdjustments_HasNoCornerRadius()
    {
        var (converted, _) = ConvertSingleShape(CreateDeck(
            PresetShape(2, Drawing.ShapeTypeValues.Rectangle, 1270000, 1270000)));

        Assert.Equal(0.0, converted.CornerRadius);
    }

    private static P.Shape PresetShape(uint id, Drawing.ShapeTypeValues preset, long cx, long cy, int? rot = null, int? adj = null)
    {
        var transform = rot.HasValue
            ? new Drawing.Transform2D(
                new Drawing.Offset { X = 1000000, Y = 500000 },
                new Drawing.Extents { Cx = cx, Cy = cy })
            { Rotation = rot.Value }
            : new Drawing.Transform2D(
                new Drawing.Offset { X = 1000000, Y = 500000 },
                new Drawing.Extents { Cx = cx, Cy = cy });

        var presetGeometry = new Drawing.PresetGeometry(
            adj.HasValue
                ? new Drawing.AdjustValueList(
                    new Drawing.ShapeGuide { Name = "adj", Formula = $"val {adj.Value}" })
                : new Drawing.AdjustValueList())
        { Preset = preset };

        return new P.Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = id, Name = $"Shape {id}" },
                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(
                transform,
                presetGeometry,
                new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = "C00000" })));
    }

    private static (TypstShapeElement Shape, string Source) ConvertSingleShape(string path)
    {
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var source = converter.GenerateTypstSource(presentation);

        var element = Assert.Single(presentation.Slides[0].Elements, e => e.Type == "Shape");
        return (element.Shape!, source);
    }

    private string CreateDeck(params OpenXmlElement[] slideElements)
    {
        var path = Path.Combine(_tempDir, $"{Guid.NewGuid():N}.pptx");

        using (var document = PresentationDocument.Create(path, PresentationDocumentType.Presentation))
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
}
