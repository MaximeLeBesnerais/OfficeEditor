using PptxEditor.Core.Generation.Emit.Typst;
using PptxEditor.Core.Generation.Layout;
using PptxEditor.Core.Generation.Model;
using PptxEditor.Core.Models;

namespace DocxEditor.Tests.Generation.Emit.Typst;

/// <summary>
/// Unit tests for emitter behaviors the JSON fixtures cannot reach (font scale, stretch fit,
/// text shadows, error paths) plus formatting invariants (escaping, negative zero, clip origins).
/// </summary>
public sealed class TypstEmitterUnitTests
{
    private static ResolvedSlide SlideWith(ResolvedElement child)
        => new()
        {
            WidthPt = 960,
            HeightPt = 540,
            Root = new ResolvedContainer
            {
                X = 0, Y = 0, Width = 960, Height = 540,
                Overflow = OverflowPolicy.Error,
                Children = [child]
            }
        };

    private static string EmitChild(ResolvedElement child) => new TypstEmitter().EmitSlide(SlideWith(child));

    private static ResolvedText Text(IReadOnlyList<ResolvedTextRun> runs)
        => new()
        {
            X = 10, Y = 20, Width = 200, Height = 40,
            Runs = runs,
            Insets = new EdgeInsets(0, 0, 0, 0),
            Overflow = OverflowPolicy.Shrink
        };

    [Fact]
    public void Emit_NullLayoutResult_Throws()
        => Assert.Throws<ArgumentNullException>(() => new TypstEmitter().Emit(null!));

    [Fact]
    public void EmitSlide_NullSlide_Throws()
        => Assert.Throws<ArgumentNullException>(() => new TypstEmitter().EmitSlide(null!));

    [Fact]
    public void Text_FontScale_ScalesRunSizes()
    {
        var text = Text([new ResolvedTextRun
        {
            Text = "Scaled", FontFamily = "Aptos", FontSizePt = 20, ColorHex = "#1A1A1A", Bold = false, Italic = false
        }]) with { FontScale = 0.75 };

        var source = EmitChild(text);

        Assert.Contains("size: 15pt", source);
    }

    [Fact]
    public void Text_Shadow_EmitsOffsetCopyInShadowColorBeforeRealText()
    {
        var text = Text([new ResolvedTextRun
        {
            Text = "Shadowed", FontFamily = "Aptos", FontSizePt = 14, ColorHex = "#1A1A1A", Bold = false, Italic = false
        }]) with
        {
            Shadow = new ShadowSpec { Color = "#000000", Dx = 2, Dy = 3, Blur = 4, Alpha = 0.5 }
        };

        var lines = EmitChild(text).Split('\n', StringSplitOptions.RemoveEmptyEntries);

        // Page setup lines, then the shadow copy, then the real text.
        var shadowLine = Assert.Single(lines, l => l.Contains("fill: rgb(\"#00000080\")"));
        var realLine = Assert.Single(lines, l => l.Contains("fill: rgb(\"#1A1A1A\")"));
        Assert.Contains("dx: 12pt", shadowLine); // 10 + 2
        Assert.Contains("dy: 23pt", shadowLine); // 20 + 3
        Assert.Contains("dx: 10pt", realLine);
        Assert.True(
            Array.IndexOf(lines, shadowLine) < Array.IndexOf(lines, realLine),
            "shadow copy must paint before the real text");
    }

    [Fact]
    public void Text_EmptyRuns_Throws()
    {
        var text = Text([]);
        var ex = Assert.Throws<TypstEmitException>(() => EmitChild(text));
        Assert.Contains("no runs", ex.Message);
    }

    [Fact]
    public void Text_MarkupCharacters_AreEscaped()
    {
        var text = Text([new ResolvedTextRun
        {
            Text = "a_b[c]#d$e%f&g", FontFamily = "Aptos", FontSizePt = 14, ColorHex = null, Bold = false, Italic = false
        }]);

        var source = EmitChild(text);

        Assert.Contains("a\\_b\\[c\\]\\#d\\$e\\%f\\&g", source);
    }

    [Fact]
    public void Text_FontFamilyWithQuote_IsEscapedInStringLiteral()
    {
        var text = Text([new ResolvedTextRun
        {
            Text = "x", FontFamily = "Weird \"Font\"", FontSizePt = 14, ColorHex = null, Bold = false, Italic = false
        }]);

        var source = EmitChild(text);

        Assert.Contains("font: \"Weird \\\"Font\\\"\"", source);
    }

    [Fact]
    public void Text_Defaults_OmitOptionalParameters()
    {
        var text = Text([new ResolvedTextRun
        {
            Text = "Plain", FontFamily = null, FontSizePt = 14, ColorHex = null, Bold = false, Italic = false
        }]);

        var source = EmitChild(text);

        Assert.Contains("#text(size: 14pt)[Plain]", source);
        Assert.Contains("#align(top + left)", source);
    }

    [Fact]
    public void Line_NoStroke_UsesDocumentedDefault()
    {
        var line = new ResolvedLine { X = 10, Y = 20, Width = 300, Height = 2, Orientation = LineOrientation.Horizontal };

        var source = EmitChild(line);

        Assert.Contains($"stroke: {TypstEmitter.DefaultLineWidthPt}pt + rgb(\"{TypstEmitter.DefaultLineColorHex}\")", source);
    }

    [Fact]
    public void Line_Vertical_SpansBoxOnVerticalAxis()
    {
        var line = new ResolvedLine
        {
            X = 10, Y = 20, Width = 4, Height = 100,
            Orientation = LineOrientation.Vertical,
            Stroke = new StrokeSpec { Color = "#123456", WidthPt = 1 }
        };

        var source = EmitChild(line);

        Assert.Contains("#line(start: (2pt, 0pt), end: (2pt, 100pt)", source);
    }

    [Fact]
    public void Image_StretchFit_EmitsStretch()
    {
        var image = new ResolvedImage
        {
            X = 10, Y = 20, Width = 200, Height = 100,
            Source = "assets/pic.png", Fit = ImageFitMode.Stretch
        };

        var source = EmitChild(image);

        Assert.Contains("#image(\"assets/pic.png\", width: 200pt, height: 100pt, fit: \"stretch\")", source);
    }

    [Fact]
    public void Image_Base64Source_ThrowsLoud()
    {
        var image = new ResolvedImage
        {
            X = 0, Y = 0, Width = 100, Height = 100,
            Source = "data:image/png;base64,iVBORw0KGgo=", Fit = ImageFitMode.Fill
        };

        var ex = Assert.Throws<TypstEmitException>(() => EmitChild(image));
        Assert.Contains("base64", ex.Message);
    }

    [Theory]
    [InlineData("http://example.com/pic.png")]
    [InlineData("https://example.com/pic.png")]
    public void Image_RemoteSource_ThrowsLoud(string source)
    {
        var image = new ResolvedImage
        {
            X = 0, Y = 0, Width = 100, Height = 100,
            Source = source, Fit = ImageFitMode.Fill
        };

        var ex = Assert.Throws<TypstEmitException>(() => EmitChild(image));
        Assert.Contains("remote image URLs", ex.Message);
    }

    [Fact]
    public void Image_CropLeavingNoVisibleRegion_ThrowsLoud()
    {
        var image = new ResolvedImage
        {
            X = 0, Y = 0, Width = 100, Height = 100,
            Source = "assets/pic.png", Fit = ImageFitMode.Crop,
            Crop = new SourceRect(60000, 0, 50000, 0)
        };

        var ex = Assert.Throws<TypstEmitException>(() => EmitChild(image));
        Assert.Contains("no visible source region", ex.Message);
    }

    [Fact]
    public void Image_CropRect_ScalesAndOffsetsImageInsideClippedBlock()
    {
        // Crop 10% left and right: visible width 80% → scaled width 250 for a 200pt box,
        // offset by -10% of the scaled width.
        var image = new ResolvedImage
        {
            X = 10, Y = 20, Width = 200, Height = 100,
            Source = "assets/pic.png", Fit = ImageFitMode.Crop,
            Crop = new SourceRect(10000, 0, 10000, 0)
        };

        var source = EmitChild(image);

        Assert.Contains(
            "#place(top + left, dx: 10pt, dy: 20pt)[#block(width: 200pt, height: 100pt, clip: true)" +
            "[#place(top + left, dx: -25pt, dy: 0pt)[#image(\"assets/pic.png\", width: 250pt, height: 100pt, fit: \"stretch\")]]]",
            source);
    }

    [Fact]
    public void Container_ClipOverflow_PlacesChildrenRelativeToContainer()
    {
        var container = new ResolvedContainer
        {
            X = 100, Y = 100, Width = 200, Height = 200,
            Overflow = OverflowPolicy.Clip,
            Children =
            [
                new ResolvedRect { X = 150, Y = 120, Width = 40, Height = 40, Fill = new SolidFill("#FF0000") }
            ]
        };

        var source = EmitChild(container);
        var lines = source.Split('\n', StringSplitOptions.RemoveEmptyEntries);

        Assert.Contains("#place(top + left, dx: 100pt, dy: 100pt)[#block(width: 200pt, height: 200pt, clip: true)[", lines);
        // Child at slide (150,120) is placed relative to the clip origin (100,100).
        Assert.Contains(
            "#place(top + left, dx: 50pt, dy: 20pt)[#rect(width: 40pt, height: 40pt, fill: rgb(\"#FF0000\"))]",
            lines);
        Assert.Equal("]]", lines[^1]);
    }

    [Fact]
    public void Container_RadiusWithoutFillOrStroke_PaintsNoBackground()
    {
        var container = new ResolvedContainer
        {
            X = 10, Y = 10, Width = 100, Height = 100,
            Radius = CornerRadii.All(8),
            Overflow = OverflowPolicy.Error,
            Children = []
        };

        var source = EmitChild(container);

        Assert.DoesNotContain("#rect(", source);
    }

    [Fact]
    public void Formatting_NegativeZero_IsNormalized()
    {
        var image = new ResolvedImage
        {
            X = 0, Y = 0, Width = 100, Height = 100,
            Source = "assets/pic.png", Fit = ImageFitMode.Crop,
            Crop = new SourceRect(0, 0, 10000, 0)
        };

        var source = EmitChild(image);

        Assert.DoesNotContain("-0pt", source);
    }

    [Fact]
    public void Slide_EmitsPageSetupPerSlideAndPagebreakBetweenSlides()
    {
        var result = new LayoutResult
        {
            Slides =
            [
                SlideWith(new ResolvedRect { X = 0, Y = 0, Width = 10, Height = 10, Fill = new SolidFill("#000000") }),
                SlideWith(new ResolvedRect { X = 0, Y = 0, Width = 10, Height = 10, Fill = new SolidFill("#000000") })
            ],
            Warnings = []
        };

        var source = new TypstEmitter().Emit(result);

        Assert.Equal(2, CountOccurrences(source, "#set page(width:"));
        Assert.Equal(1, CountOccurrences(source, "#pagebreak()"));
    }

    private static int CountOccurrences(string haystack, string needle)
        => haystack.Split(needle, StringSplitOptions.None).Length - 1;
}
