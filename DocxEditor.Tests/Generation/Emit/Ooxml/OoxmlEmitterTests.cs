using System.Buffers.Binary;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using PptxEditor.Core.Generation.Emit.Ooxml;
using PptxEditor.Core.Generation.Layout;
using PptxEditor.Core.Generation.Model;
using PptxEditor.Core.Generation.Schema;
using PptxEditor.Core.Models;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace DocxEditor.Tests.Generation.Emit.Ooxml;

/// <summary>
/// OOXML emitter acceptance (P4): every Tier-1 primitive + linear gradient emits valid
/// OOXML (OpenXmlValidator-clean = opens without repair), radius adj mapping through real
/// XML, spcPct units, no autofit anywhere. Attribute read-backs use regex on OuterXml
/// (raw XML attribute reads via regex on OuterXml).
/// </summary>
public sealed class OoxmlEmitterTests : IDisposable
{
    private readonly string _testDir = CreateTempDirectory();

    public void Dispose()
    {
        try { Directory.Delete(_testDir, recursive: true); } catch { /* best effort */ }
    }

    #region Helpers (regex attribute reads on OuterXml)

    private static string? Attr(OpenXmlElement element, string attr)
    {
        var match = Regex.Match(element.OuterXml, $@"\b{Regex.Escape(attr)}\s*=\s*""([^""]*)""");
        return match.Success ? match.Groups[1].Value : null;
    }

    private static string CreateTempDirectory()
    {
        var path = Path.Combine(Path.GetTempPath(), $"ooxml-emitter-tests-{Guid.NewGuid():N}");
        Directory.CreateDirectory(path);
        return path;
    }

    private static byte[] PngWithSize(int width, int height)
    {
        var png = new byte[33];
        png[0] = 0x89; png[1] = 0x50; png[2] = 0x4E; png[3] = 0x47;
        png[4] = 0x0D; png[5] = 0x0A; png[6] = 0x1A; png[7] = 0x0A;
        png[11] = 0x0D;
        png[12] = (byte)'I'; png[13] = (byte)'H'; png[14] = (byte)'D'; png[15] = (byte)'R';
        BinaryPrimitives.WriteInt32BigEndian(png.AsSpan(16), width);
        BinaryPrimitives.WriteInt32BigEndian(png.AsSpan(20), height);
        png[24] = 0x08; png[25] = 0x06;
        return png;
    }

    private string WriteImage(string name, byte[] bytes)
    {
        var path = Path.Combine(_testDir, name);
        File.WriteAllBytes(path, bytes);
        return path;
    }

    private static ResolvedContainer RootWith(params ResolvedElement[] children) => new()
    {
        X = 0, Y = 0, Width = 960, Height = 540,
        Overflow = OverflowPolicy.Error,
        Children = children
    };

    private static LayoutResult LayoutWith(params ResolvedElement[] children) => new()
    {
        Slides = [new ResolvedSlide { WidthPt = 960, HeightPt = 540, Root = RootWith(children) }],
        Warnings = []
    };

    private static PresentationDocument EmitAndOpen(LayoutResult layout, out OoxmlEmissionResult result)
    {
        result = new OoxmlEmitter().Emit(layout);
        return PresentationDocument.Open(new MemoryStream(result.Bytes), false);
    }

    private static P.ShapeTree FirstShapeTree(PresentationDocument document) =>
        document.PresentationPart!.SlideParts.First().Slide!.CommonSlideData!.ShapeTree!;

    private static void AssertValidates(PresentationDocument document)
    {
        var errors = new OpenXmlValidator().Validate(document).ToList();
        Assert.True(errors.Count == 0,
            string.Join("\n", errors.Select(e => $"{e.Path}: {e.Description}")));
    }

    #endregion

    #region Deck scaffolding

    [Fact]
    public void Emit_NullLayout_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => new OoxmlEmitter().Emit(null!));
    }

    [Fact]
    public void Emit_NoSlides_Throws()
    {
        var layout = new LayoutResult { Slides = [], Warnings = [] };
        Assert.Throws<ArgumentException>(() => new OoxmlEmitter().Emit(layout));
    }

    [Fact]
    public void Emit_SetsSlideSizeFromResolvedSlide()
    {
        using var document = EmitAndOpen(LayoutWith(new ResolvedRect { X = 0, Y = 0, Width = 10, Height = 10 }), out _);
        var slideSize = document.PresentationPart!.Presentation!.SlideSize!;
        Assert.Equal(960 * 12700, slideSize.Cx!.Value);
        Assert.Equal(540 * 12700, slideSize.Cy!.Value);
    }

    [Fact]
    public void Emit_EmptySlide_DeckValidates()
    {
        using var document = EmitAndOpen(new LayoutResult
        {
            Slides = [new ResolvedSlide { WidthPt = 960, HeightPt = 540, Root = RootWith() }],
            Warnings = []
        }, out _);
        AssertValidates(document);
    }

    [Fact]
    public void Emit_AllPrimitivesTogether_DeckValidates()
    {
        var imagePath = WriteImage("photo.png", PngWithSize(4000, 2000));
        var layout = LayoutWith(
            new ResolvedRect
            {
                X = 10, Y = 10, Width = 200, Height = 100,
                Fill = new SolidFill("#0B3D91"), Radius = CornerRadii.All(8),
                Shadow = new ShadowSpec { Color = "#000000", Dx = 2, Dy = 2, Blur = 6, Alpha = 0.4 }
            },
            new ResolvedEllipse
            {
                X = 220, Y = 10, Width = 80, Height = 80,
                Fill = new LinearGradientFill
                {
                    Angle = 45,
                    Stops = [new GradientStop { Color = "#FF6B00", Offset = 0 }, new GradientStop { Color = "#1A1A1A", Offset = 1, Alpha = 0.5 }]
                }
            },
            new ResolvedText
            {
                X = 10, Y = 120, Width = 290, Height = 40,
                Runs = [new ResolvedTextRun { Text = "Hello", FontFamily = "Aptos", FontSizePt = 20, ColorHex = "#1A1A1A", Bold = true, Italic = false }],
                TextAlign = TextAlign.Center, Anchor = TextAnchor.Middle,
                Insets = new EdgeInsets(4, 8, 4, 8), Overflow = OverflowPolicy.Shrink
            },
            new ResolvedLine { X = 10, Y = 170, Width = 200, Height = 2, Orientation = LineOrientation.Horizontal, Stroke = new StrokeSpec { Color = "#8A94A6", WidthPt = 2 } },
            new ResolvedLine { X = 10, Y = 180, Width = 200, Height = 2, IsConnector = true, Orientation = LineOrientation.Vertical },
            new ResolvedImage { X = 320, Y = 120, Width = 100, Height = 100, Source = imagePath, Fit = ImageFitMode.Fill, Alt = "photo" },
            new ResolvedGroup
            {
                X = 0, Y = 0, Width = 960, Height = 540,
                Children = [new ResolvedRect { X = 10, Y = 300, Width = 50, Height = 50, Fill = new SolidFill("#FF6B00") }]
            });
        using var document = EmitAndOpen(layout, out var result);
        Assert.Empty(result.Warnings);
        AssertValidates(document);
    }

    #endregion

    #region Rect + radius adj values (acceptance: per-corner mapping through real XML)

    private static P.Shape SingleShape(LayoutResult layout)
    {
        var document = PresentationDocument.Open(new MemoryStream(new OoxmlEmitter().Emit(layout).Bytes), false);
        var shape = FirstShapeTree(document).Elements<P.Shape>().Single();
        document.Dispose();
        return shape;
    }

    [Theory]
    [InlineData(10, 10, 10, 10, "roundRect", "val 10000", null)]        // uniform
    [InlineData(0, 10, 0, 0, "round1Rect", "val 10000", "val 0")]       // top-right only
    [InlineData(10, 0, 10, 10, "round1Rect", "val 0", "val 10000")]     // three rounded, tr square
    [InlineData(10, 10, 0, 0, "round2SameRect", "val 10000", "val 0")]  // top pair
    [InlineData(20, 20, 5, 5, "round2SameRect", "val 20000", "val 5000")]
    [InlineData(0, 10, 0, 10, "round2DiagRect", "val 10000", "val 0")]  // diagonal
    public void Rect_RadiusEmitsExpectedPresetAndAdjValues(
        double tl, double tr, double br, double bl, string preset, string adj1, string? adj2)
    {
        // Box 200×100 → adj base = min = 100pt.
        var shape = SingleShape(LayoutWith(new ResolvedRect
        {
            X = 0, Y = 0, Width = 200, Height = 100,
            Fill = new SolidFill("#0B3D91"),
            Radius = new CornerRadii(tl, tr, br, bl)
        }));

        var geometry = shape.ShapeProperties!.Elements<Drawing.PresetGeometry>().Single();
        Assert.Equal(preset, Attr(geometry, "prst"));

        var guides = geometry.Elements<Drawing.AdjustValueList>().Single()
            .Elements<Drawing.ShapeGuide>()
            .Select(g => (Name: Attr(g, "name"), Formula: Attr(g, "fmla")))
            .ToList();
        Assert.Equal(("adj1", adj1), guides[0]);
        if (adj2 is null)
        {
            Assert.Single(guides);
        }
        else
        {
            Assert.Equal(("adj2", adj2), guides[1]);
        }
    }

    [Fact]
    public void Rect_NoRadius_EmitsPlainRect()
    {
        var shape = SingleShape(LayoutWith(new ResolvedRect { X = 5, Y = 6, Width = 200, Height = 100, Fill = new SolidFill("#0B3D91") }));
        var geometry = shape.ShapeProperties!.Elements<Drawing.PresetGeometry>().Single();
        Assert.Equal("rect", Attr(geometry, "prst"));
        Assert.Empty(geometry.Elements<Drawing.AdjustValueList>().Single().Elements<Drawing.ShapeGuide>());
    }

    [Fact]
    public void Rect_GeometryAndFillInEmuAndHex()
    {
        var shape = SingleShape(LayoutWith(new ResolvedRect
        {
            X = 5, Y = 6, Width = 200, Height = 100,
            Fill = new SolidFill("#0b3d91"),
            Stroke = new StrokeSpec { Color = "#FF6B00", WidthPt = 2.5 }
        }));
        var transform = shape.ShapeProperties!.Transform2D!;
        Assert.Equal("63500", Attr(transform, "x"));
        Assert.Equal("76200", Attr(transform, "y"));
        Assert.Equal("2540000", Attr(transform, "cx"));
        Assert.Equal("1270000", Attr(transform, "cy"));

        var fill = shape.ShapeProperties.Elements<Drawing.SolidFill>().Single();
        Assert.Equal("0B3D91", Attr(fill, "val"));

        var outline = shape.ShapeProperties.Elements<Drawing.Outline>().Single();
        Assert.Equal("31750", Attr(outline, "w")); // 2.5pt × 12700
        Assert.Equal("FF6B00", Attr(outline, "val"));
    }

    [Fact]
    public void Rect_NoFillNoStroke_EmitsExplicitNoFillAndNoOutline()
    {
        var shape = SingleShape(LayoutWith(new ResolvedRect { X = 0, Y = 0, Width = 10, Height = 10 }));
        Assert.Single(shape.ShapeProperties!.Elements<Drawing.NoFill>());
        var outline = shape.ShapeProperties.Elements<Drawing.Outline>().Single();
        Assert.Single(outline.Elements<Drawing.NoFill>());
    }

    #endregion

    #region Gradient + shadow

    [Fact]
    public void Gradient_LinearEmitsGsLstLinAngleAndAlpha()
    {
        var shape = SingleShape(LayoutWith(new ResolvedRect
        {
            X = 0, Y = 0, Width = 100, Height = 100,
            Fill = new LinearGradientFill
            {
                Angle = 90,
                Stops =
                [
                    new GradientStop { Color = "#FFFFFF", Offset = 1, Alpha = 0.5 },
                    new GradientStop { Color = "#000000", Offset = 0 }
                ]
            }
        }));

        var gradient = shape.ShapeProperties!.Elements<Drawing.GradientFill>().Single();
        var stops = gradient.Elements<Drawing.GradientStopList>().Single().Elements<Drawing.GradientStop>().ToList();
        Assert.Equal(2, stops.Count);
        Assert.Equal("0", Attr(stops[0], "pos")); // stops sorted ascending
        Assert.Equal("000000", Attr(stops[0], "val"));
        Assert.Equal("100000", Attr(stops[1], "pos")); // spcPct family: 100000 = 100%
        Assert.Equal("FFFFFF", Attr(stops[1].Elements<Drawing.RgbColorModelHex>().Single(), "val"));
        Assert.Equal("50000", Attr(stops[1].Elements<Drawing.RgbColorModelHex>().Single().Elements<Drawing.Alpha>().Single(), "val"));

        var linear = gradient.Elements<Drawing.LinearGradientFill>().Single();
        Assert.Equal("5400000", Attr(linear, "ang")); // 90° × 60000
    }

    [Fact]
    public void Gradient_AlphaIsPct1000InsideStopColor()
    {
        var shape = SingleShape(LayoutWith(new ResolvedRect
        {
            X = 0, Y = 0, Width = 100, Height = 100,
            Fill = new LinearGradientFill
            {
                Angle = 0,
                Stops =
                [
                    new GradientStop { Color = "#FF6B00", Offset = 0, Alpha = 0.25 },
                    new GradientStop { Color = "#0B3D91", Offset = 1 }
                ]
            }
        }));
        var stops = shape.ShapeProperties!.Descendants<Drawing.GradientStop>().ToList();
        var alpha = stops[0].Elements<Drawing.RgbColorModelHex>().Single().Elements<Drawing.Alpha>().Single();
        Assert.Equal("25000", Attr(alpha, "val"));
        Assert.Empty(stops[1].Elements<Drawing.RgbColorModelHex>().Single().Elements<Drawing.Alpha>());
    }

    [Fact]
    public void Shadow_NativeOuterShadowWithEmuDistBlurAndDirection()
    {
        var shape = SingleShape(LayoutWith(new ResolvedRect
        {
            X = 0, Y = 0, Width = 100, Height = 100,
            Fill = new SolidFill("#0B3D91"),
            Shadow = new ShadowSpec { Color = "#000000", Dx = 3, Dy = 4, Blur = 10, Alpha = 0.4 }
        }));

        var shadow = shape.ShapeProperties!.Elements<Drawing.EffectList>().Single()
            .Elements<Drawing.OuterShadow>().Single();
        Assert.Equal("127000", Attr(shadow, "blurRad")); // 10pt × 12700
        Assert.Equal("63500", Attr(shadow, "dist"));     // hypot(3, 4) = 5pt
        Assert.Equal("3187806", Attr(shadow, "dir"));    // atan2(4, 3) = 53.1301° × 60000

        var color = shadow.Elements<Drawing.RgbColorModelHex>().Single();
        Assert.Equal("000000", Attr(color, "val"));
        Assert.Equal("40000", Attr(color.Elements<Drawing.Alpha>().Single(), "val"));
    }

    #endregion

    #region Ellipse / line / connector

    [Fact]
    public void Ellipse_EmitsEllipsePresetWithFill()
    {
        var shape = SingleShape(LayoutWith(new ResolvedEllipse { X = 1, Y = 2, Width = 30, Height = 40, Fill = new SolidFill("#FF6B00") }));
        var geometry = shape.ShapeProperties!.Elements<Drawing.PresetGeometry>().Single();
        Assert.Equal("ellipse", Attr(geometry, "prst"));
        Assert.Equal("FF6B00", Attr(shape.ShapeProperties.Elements<Drawing.SolidFill>().Single(), "val"));
    }

    [Fact]
    public void Line_HorizontalCentersInBoxWithZeroHeight()
    {
        var shape = SingleShape(LayoutWith(new ResolvedLine
        {
            X = 10, Y = 20, Width = 100, Height = 4,
            Orientation = LineOrientation.Horizontal,
            Stroke = new StrokeSpec { Color = "#8A94A6", WidthPt = 2 }
        }));
        var geometry = shape.ShapeProperties!.Elements<Drawing.PresetGeometry>().Single();
        Assert.Equal("line", Attr(geometry, "prst"));

        var transform = shape.ShapeProperties.Transform2D!;
        Assert.Equal("127000", Attr(transform, "x"));
        Assert.Equal("279400", Attr(transform, "y"));  // (20 + 4/2) × 12700
        Assert.Equal("1270000", Attr(transform, "cx"));
        Assert.Equal("0", Attr(transform, "cy"));

        var outline = shape.ShapeProperties.Elements<Drawing.Outline>().Single();
        Assert.Equal("25400", Attr(outline, "w"));
        Assert.Equal("8A94A6", Attr(outline, "val"));
    }

    [Fact]
    public void Line_NullStroke_FallsBackToOnePointBlack()
    {
        var shape = SingleShape(LayoutWith(new ResolvedLine { X = 0, Y = 0, Width = 50, Height = 0, Orientation = LineOrientation.Horizontal }));
        var outline = shape.ShapeProperties!.Elements<Drawing.Outline>().Single();
        Assert.Equal("12700", Attr(outline, "w"));
        Assert.Equal("000000", Attr(outline, "val"));
    }

    [Fact]
    public void Connector_EmitsCxnSpWithStraightConnector1()
    {
        var layout = LayoutWith(new ResolvedLine
        {
            X = 10, Y = 10, Width = 5, Height = 100,
            IsConnector = true, Orientation = LineOrientation.Vertical
        });
        using var document = EmitAndOpen(layout, out _);
        var connector = FirstShapeTree(document).Elements<P.ConnectionShape>().Single();
        var geometry = connector.ShapeProperties!.Elements<Drawing.PresetGeometry>().Single();
        Assert.Equal("straightConnector1", Attr(geometry, "prst"));

        var transform = connector.ShapeProperties.Transform2D!;
        Assert.Equal("158750", Attr(transform, "x")); // (10 + 5/2) × 12700
        Assert.Equal("0", Attr(transform, "cx"));
        Assert.Equal("1270000", Attr(transform, "cy"));
    }

    #endregion

    #region Text

    [Fact]
    public void Text_BoxAnchorAlignInsetsAndRuns()
    {
        var shape = SingleShape(LayoutWith(new ResolvedText
        {
            X = 50, Y = 40, Width = 200, Height = 60,
            Runs =
            [
                new ResolvedTextRun { Text = "Bold ", FontFamily = "Aptos Display", FontSizePt = 20, ColorHex = "#0B3D91", Bold = true, Italic = false },
                new ResolvedTextRun { Text = "italic", FontFamily = null, FontSizePt = 14, ColorHex = null, Bold = false, Italic = true }
            ],
            TextAlign = TextAlign.Center,
            Anchor = TextAnchor.Middle,
            Insets = new EdgeInsets(5, 10, 5, 10),
            Overflow = OverflowPolicy.Shrink,
            FontScale = 0.5
        }));

        var body = shape.TextBody!;
        var bodyProperties = body.Elements<Drawing.BodyProperties>().Single();
        Assert.Equal("127000", Attr(bodyProperties, "lIns"));
        Assert.Equal("63500", Attr(bodyProperties, "tIns"));
        Assert.Equal("127000", Attr(bodyProperties, "rIns"));
        Assert.Equal("63500", Attr(bodyProperties, "bIns"));
        Assert.Equal("ctr", Attr(bodyProperties, "anchor"));

        var paragraph = body.Elements<Drawing.Paragraph>().Single();
        Assert.Equal("ctr", Attr(paragraph.ParagraphProperties!, "algn"));

        var runs = paragraph.Elements<Drawing.Run>().ToList();
        Assert.Equal(2, runs.Count);

        var first = runs[0].RunProperties!;
        Assert.Equal("1000", Attr(first, "sz")); // 20pt × FontScale 0.5 × 100
        Assert.Equal("1", Attr(first, "b"));
        Assert.Null(Attr(first, "typeface"));
        Assert.Equal("0B3D91", Attr(first, "val"));

        var second = runs[1].RunProperties!;
        Assert.Equal("700", Attr(second, "sz")); // 14pt × 0.5 × 100
        Assert.Equal("1", Attr(second, "i"));
        Assert.Null(Attr(second, "b"));
    }

    [Fact]
    public void Text_NewlinesBecomeBreaks()
    {
        var shape = SingleShape(LayoutWith(new ResolvedText
        {
            X = 0, Y = 0, Width = 100, Height = 60,
            Runs = [new ResolvedTextRun { Text = "one\ntwo", FontFamily = null, FontSizePt = 12, ColorHex = null, Bold = false, Italic = false }],
            Insets = new EdgeInsets(0, 0, 0, 0),
            Overflow = OverflowPolicy.Clip
        }));
        var paragraph = shape.TextBody!.Elements<Drawing.Paragraph>().Single();
        Assert.Single(paragraph.Elements<Drawing.Break>());
        Assert.Equal(2, paragraph.Elements<Drawing.Run>().Count());
        Assert.Equal("one", paragraph.Elements<Drawing.Run>().First().Text?.Text);
        Assert.Equal("two", paragraph.Elements<Drawing.Run>().Last().Text?.Text);
    }

    [Fact]
    public void Text_AnchorBottomMapsToB()
    {
        var shape = SingleShape(LayoutWith(new ResolvedText
        {
            X = 0, Y = 0, Width = 100, Height = 60,
            Runs = [new ResolvedTextRun { Text = "x", FontFamily = null, FontSizePt = 12, ColorHex = null, Bold = false, Italic = false }],
            Anchor = TextAnchor.Bottom,
            Insets = new EdgeInsets(0, 0, 0, 0),
            Overflow = OverflowPolicy.Clip
        }));
        Assert.Equal("b", Attr(shape.TextBody!.Elements<Drawing.BodyProperties>().Single(), "anchor"));
    }

    [Fact]
    public void Text_NoAutofitAnywhere()
    {
        // Never apply global autofit. The emitter applies FontScale
        // literally, so no normAutofit element may appear at all.
        var layout = LayoutWith(new ResolvedText
        {
            X = 0, Y = 0, Width = 100, Height = 60,
            Runs = [new ResolvedTextRun { Text = "shrunk", FontFamily = null, FontSizePt = 20, ColorHex = null, Bold = false, Italic = false }],
            Insets = new EdgeInsets(0, 0, 0, 0),
            Overflow = OverflowPolicy.Shrink,
            FontScale = 0.6
        });
        using var document = EmitAndOpen(layout, out _);
        Assert.DoesNotContain("normAutofit", FirstShapeTree(document).OuterXml, StringComparison.Ordinal);
        Assert.DoesNotContain("spAutoFit", FirstShapeTree(document).OuterXml, StringComparison.Ordinal);
    }

    #endregion

    #region Containers, groups, paint order

    [Fact]
    public void Container_SurfacePaintsBelowChildren()
    {
        var layout = LayoutWith();
        layout = new LayoutResult
        {
            Slides = [new ResolvedSlide
            {
                WidthPt = 960, HeightPt = 540,
                Root = new ResolvedContainer
                {
                    X = 0, Y = 0, Width = 960, Height = 540,
                    Fill = new SolidFill("#FFFFFF"),
                    Overflow = OverflowPolicy.Error,
                    Children = [new ResolvedRect { X = 1, Y = 2, Width = 10, Height = 10, Fill = new SolidFill("#0B3D91") }]
                }
            }],
            Warnings = []
        };
        using var document = EmitAndOpen(layout, out _);
        // A solid-filled root becomes a real p:bg (no full-slide rect); the child is the
        // only shape, and the background still paints below it by definition.
        var commonSlideData = document.PresentationPart!.SlideParts.First().Slide!.CommonSlideData!;
        var background = Assert.IsType<P.BackgroundProperties>(commonSlideData.Background?.BackgroundProperties);
        Assert.Equal("FFFFFF", Attr(background.Elements<Drawing.SolidFill>().Single(), "val"));
        var shapes = FirstShapeTree(document).Elements<P.Shape>().ToList();
        var shape = Assert.Single(shapes);
        Assert.Equal("0B3D91", Attr(shape.ShapeProperties!.Elements<Drawing.SolidFill>().Single(), "val"));
    }

    [Fact]
    public void Container_WithoutVisualsEmitsNoSurfaceShape()
    {
        using var document = EmitAndOpen(LayoutWith(new ResolvedRect { X = 0, Y = 0, Width = 10, Height = 10 }), out _);
        Assert.Single(FirstShapeTree(document).Elements<P.Shape>());
    }

    [Fact]
    public void Root_GradientFill_KeepsFullSlideRectSurface()
    {
        // Known limitation: only solid root fills become p:bg; gradients keep the rect.
        var layout = new LayoutResult
        {
            Slides = [new ResolvedSlide
            {
                WidthPt = 960, HeightPt = 540,
                Root = new ResolvedContainer
                {
                    X = 0, Y = 0, Width = 960, Height = 540,
                    Fill = new LinearGradientFill
                    {
                        Angle = 90,
                        Stops = [new GradientStop { Color = "#FFFFFF", Offset = 0 }, new GradientStop { Color = "#000000", Offset = 1 }]
                    },
                    Overflow = OverflowPolicy.Error,
                    Children = []
                }
            }],
            Warnings = []
        };
        using var document = EmitAndOpen(layout, out _);
        var commonSlideData = document.PresentationPart!.SlideParts.First().Slide!.CommonSlideData!;
        Assert.Null(commonSlideData.Background);
        var shape = Assert.Single(FirstShapeTree(document).Elements<P.Shape>());
        Assert.Single(shape.ShapeProperties!.Elements<Drawing.GradientFill>());
    }

    [Fact]
    public void Group_FlattensChildrenInDocumentOrder()
    {
        var layout = LayoutWith(
            new ResolvedRect { X = 0, Y = 0, Width = 10, Height = 10, Fill = new SolidFill("#111111") },
            new ResolvedGroup
            {
                X = 0, Y = 0, Width = 100, Height = 100,
                Children =
                [
                    new ResolvedEllipse { X = 0, Y = 0, Width = 10, Height = 10, Fill = new SolidFill("#222222") },
                    new ResolvedRect { X = 0, Y = 0, Width = 10, Height = 10, Fill = new SolidFill("#333333") }
                ]
            },
            new ResolvedRect { X = 0, Y = 0, Width = 10, Height = 10, Fill = new SolidFill("#444444") });

        using var document = EmitAndOpen(layout, out _);
        var tree = FirstShapeTree(document);
        Assert.Empty(tree.Elements<P.GroupShape>());
        var fills = tree.ChildElements
            .OfType<P.Shape>()
            .Select(s => Attr(s.ShapeProperties!.Elements<Drawing.SolidFill>().Single(), "val"))
            .ToList();
        Assert.Equal(["111111", "222222", "333333", "444444"], fills);
    }

    #endregion

    #region Image fits (F7 reuse)

    private P.Picture EmitSinglePicture(LayoutResult layout, out OoxmlEmissionResult result)
    {
        var emission = new OoxmlEmitter().Emit(layout);
        result = emission;
        var document = PresentationDocument.Open(new MemoryStream(emission.Bytes), false);
        var picture = FirstShapeTree(document).Elements<P.Picture>().Single();
        document.Dispose();
        return picture;
    }

    [Fact]
    public void Image_FillEmitsCenterCropSourceRect()
    {
        var path = WriteImage("wide.png", PngWithSize(4000, 2000));
        var picture = EmitSinglePicture(LayoutWith(new ResolvedImage
        {
            X = 10, Y = 20, Width = 100, Height = 100, Source = path, Fit = ImageFitMode.Fill
        }), out var result);

        var srcRect = picture.BlipFill!.Elements<Drawing.SourceRectangle>().Single();
        Assert.Equal("25000", Attr(srcRect, "l"));
        Assert.Equal("25000", Attr(srcRect, "r"));
        Assert.Empty(result.Warnings);
    }

    [Fact]
    public void Image_CropEmitsVerbatimSourceRect()
    {
        var path = WriteImage("any.png", PngWithSize(100, 100));
        var picture = EmitSinglePicture(LayoutWith(new ResolvedImage
        {
            X = 0, Y = 0, Width = 50, Height = 50, Source = path,
            Fit = ImageFitMode.Crop, Crop = new SourceRect(10000, 20000, 30000, 40000)
        }), out _);
        var srcRect = picture.BlipFill!.Elements<Drawing.SourceRectangle>().Single();
        Assert.Equal("10000", Attr(srcRect, "l"));
        Assert.Equal("20000", Attr(srcRect, "t"));
        Assert.Equal("30000", Attr(srcRect, "r"));
        Assert.Equal("40000", Attr(srcRect, "b"));
    }

    [Fact]
    public void Image_CropWithoutRect_BehavesAsFill()
    {
        var path = WriteImage("wide2.png", PngWithSize(4000, 2000));
        var picture = EmitSinglePicture(LayoutWith(new ResolvedImage
        {
            X = 0, Y = 0, Width = 100, Height = 100, Source = path, Fit = ImageFitMode.Crop, Crop = null
        }), out _);
        Assert.Equal("25000", Attr(picture.BlipFill!.Elements<Drawing.SourceRectangle>().Single(), "l"));
    }

    [Fact]
    public void Image_ContainShrinksFrameAroundCenterAndNeverCrops()
    {
        var path = WriteImage("tall.png", PngWithSize(2000, 4000));
        var picture = EmitSinglePicture(LayoutWith(new ResolvedImage
        {
            X = 10, Y = 20, Width = 100, Height = 100, Source = path, Fit = ImageFitMode.Contain
        }), out _);

        Assert.Empty(picture.BlipFill!.Elements<Drawing.SourceRectangle>());
        var transform = picture.ShapeProperties!.Transform2D!;
        Assert.Equal("444500", Attr(transform, "x"));  // 10pt + (100 - 50)/2 pt = 35pt
        Assert.Equal("254000", Attr(transform, "y"));
        Assert.Equal("635000", Attr(transform, "cx")); // 50pt
        Assert.Equal("1270000", Attr(transform, "cy"));
    }

    [Fact]
    public void Image_UnknownDimensions_FallsBackToStretchWithWarning()
    {
        var path = WriteImage("broken.png", [0x00, 0x01, 0x02]);
        var picture = EmitSinglePicture(LayoutWith(new ResolvedImage
        {
            X = 0, Y = 0, Width = 100, Height = 100, Source = path, Fit = ImageFitMode.Fill
        }), out var result);

        Assert.Empty(picture.BlipFill!.Elements<Drawing.SourceRectangle>());
        var warning = Assert.Single(result.Warnings);
        Assert.Contains("fell back to 'stretch'", warning);
    }

    [Fact]
    public void Image_DataUriSourceAndAltText()
    {
        var dataUri = "data:image/png;base64," + Convert.ToBase64String(PngWithSize(4000, 2000));
        var picture = EmitSinglePicture(LayoutWith(new ResolvedImage
        {
            X = 0, Y = 0, Width = 100, Height = 100, Source = dataUri, Fit = ImageFitMode.Fill, Alt = "hero image"
        }), out _);

        Assert.Equal("25000", Attr(picture.BlipFill!.Elements<Drawing.SourceRectangle>().Single(), "l"));
        Assert.Equal("hero image", Attr(picture.NonVisualPictureProperties!.Elements<P.NonVisualDrawingProperties>().Single(), "descr"));
        Assert.NotNull(picture.BlipFill.Blip?.Embed?.Value);
    }

    [Fact]
    public void Image_RemoteUrl_Throws()
    {
        var layout = LayoutWith(new ResolvedImage { X = 0, Y = 0, Width = 10, Height = 10, Source = "https://example.com/a.png", Fit = ImageFitMode.Fill });
        Assert.Throws<ArgumentException>(() => new OoxmlEmitter().Emit(layout));
    }

    [Fact]
    public void Image_MissingFile_Throws()
    {
        var layout = LayoutWith(new ResolvedImage { X = 0, Y = 0, Width = 10, Height = 10, Source = "/no/such/file.png", Fit = ImageFitMode.Fill });
        Assert.Throws<FileNotFoundException>(() => new OoxmlEmitter().Emit(layout));
    }

    [Fact]
    public void Image_EmptySource_Throws()
    {
        var layout = LayoutWith(new ResolvedImage { X = 0, Y = 0, Width = 10, Height = 10, Source = " ", Fit = ImageFitMode.Fill });
        Assert.Throws<ArgumentException>(() => new OoxmlEmitter().Emit(layout));
    }

    [Fact]
    public void Image_TraversalSource_Throws()
    {
        // "../../…" must be rejected whether or not the target exists: the containment
        // check runs before any file probe (repo root first, then the CWD fallback,
        // which is containment-checked too), so neither root can be escaped.
        var layout = LayoutWith(new ResolvedImage { X = 0, Y = 0, Width = 10, Height = 10, Source = "../../../../../../etc/passwd", Fit = ImageFitMode.Fill });
        Assert.Throws<ArgumentException>(() => new OoxmlEmitter().Emit(layout));
    }

    [Fact]
    public void Image_UnknownExtension_Throws()
    {
        // The default arm used to silently label any unknown extension as JPEG.
        var path = WriteImage("mystery.xyz", PngWithSize(100, 100));
        var layout = LayoutWith(new ResolvedImage { X = 0, Y = 0, Width = 10, Height = 10, Source = path, Fit = ImageFitMode.Stretch });
        var ex = Assert.Throws<ArgumentException>(() => new OoxmlEmitter().Emit(layout));
        Assert.Contains(".xyz", ex.Message);
    }

    [Fact]
    public void Image_RepoRelativeSource_ResolvesAgainstRepositoryRoot()
    {
        // demo/assets/dashboard.png ships with the repo (also exercised end-to-end by
        // the API's DeckGenerationServiceTests) — legitimate repo-root-relative sources
        // must keep resolving under the new repo-root-first precedence.
        var layout = LayoutWith(new ResolvedImage { X = 0, Y = 0, Width = 100, Height = 100, Source = "demo/assets/dashboard.png", Fit = ImageFitMode.Fill });
        using var document = EmitAndOpen(layout, out _);
        Assert.Single(FirstShapeTree(document).Elements<P.Picture>());
    }

    #endregion

    #region End-to-end: JSON → resolve → emit → validate

    [Fact]
    public void EndToEnd_GenerationDocumentProducesValidDeck()
    {
        const string json = """
        {
          "version": "2.0",
          "design": {
            "palette": { "primary": "#0B3D91", "accent": "#FF6B00", "ink": "#1A1A1A", "paper": "#FFFFFF" },
            "fonts": { "display": "Aptos Display", "body": "Aptos" },
            "metrics": { "marginPt": 40, "gutterPt": 16, "titleSizePt": 30, "bodySizePt": 14 }
          },
          "slides": [
            {
              "type": "container",
              "fill": "paper",
              "padding": 40,
              "layout": { "mode": "column", "gap": 16 },
              "children": [
                { "type": "text", "text": "Quarterly Results", "font": "display", "fontSize": 30, "color": "primary", "bold": true,
                  "size": { "h": 44 } },
                { "type": "rect", "fill": { "angle": 90, "stops": [ { "color": "accent", "offset": 0 }, { "color": "primary", "offset": 1 } ] },
                  "radius": 8, "size": { "h": 6 } },
                { "type": "text", "text": "Revenue grew 34%", "fontSize": 14, "color": "ink", "size": { "h": 24 } }
              ]
            }
          ]
        }
        """;

        var document = new GenerationDocumentParser().Parse(json);
        var layout = new LayoutResolver().Resolve(document);
        var result = new OoxmlEmitter().Emit(layout);

        using var package = PresentationDocument.Open(new MemoryStream(result.Bytes), false);
        AssertValidates(package);

        var commonSlideData = package.PresentationPart!.SlideParts.First().Slide!.CommonSlideData!;
        // The solid "paper" root fill becomes a real slide background, not a rect shape.
        var background = Assert.IsType<P.BackgroundProperties>(commonSlideData.Background?.BackgroundProperties);
        Assert.Equal("FFFFFF", Attr(background.Elements<Drawing.SolidFill>().Single(), "val"));
        var tree = commonSlideData.ShapeTree!;
        Assert.Equal(3, tree.ChildElements.OfType<P.Shape>().Count()); // text + rect + text (root surface is now p:bg)
        Assert.Contains(tree.Elements<P.Shape>(), s => s.ShapeProperties!.Elements<Drawing.GradientFill>().Any());
    }

    #endregion
}
