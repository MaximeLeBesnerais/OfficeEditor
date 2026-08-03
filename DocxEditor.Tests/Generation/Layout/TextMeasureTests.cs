using DocxEditor.Tests.Unit;
using PptxEditor.Core.Generation.Layout;
using PptxEditor.Core.Generation.Model;
using PptxEditor.Core.Models;
using PptxEditor.Core.Services;

namespace DocxEditor.Tests.Generation.Layout;

/// <summary>
/// P3 golden tests: real font-metric text measurement wired into the
/// LayoutResolver overflow seam. Metrics come from the synthetic "TestSans" TTF parsed by
/// the real OpenTypeFontMetricsReader chain (unitsPerEm 1000, uniform advance 500 → char
/// width = fontSize/2; TypoAsc − TypoDesc + TypoLineGap → default line factor 1.2), with a
/// hermetic catalog (IncludeSystemFonts = false) so results are host-independent.
///
/// With FontSize=10: char width = 5pt ("aaaa" = 20pt, space = 5pt), line pitch = 12pt.
/// </summary>
public sealed class TextMeasureTests : IDisposable
{
    private readonly string _tempDir = SlideOpsTestHelpers.CreateTempDirectory(nameof(TextMeasureTests));
    private readonly TextMeasure _measurer;

    public TextMeasureTests()
    {
        var testSansPath = TextFitTestFont.WriteFont(_tempDir, "testsans.ttf", "TestSans");
        _measurer = new TextMeasure(new FontMetricsCatalog(options: new FontMetricsCatalogOptions
        {
            AdditionalFontPaths = [("TestSans", testSansPath)],
            IncludeSystemFonts = false
        }));
    }

    public void Dispose() => SlideOpsTestHelpers.BestEffortDelete(_tempDir);

    // ------------------------------------------------------------- helpers

    private static ResolvedTextRun Run(string text, double size = 10, string family = "TestSans")
        => new() { Text = text, FontFamily = family, FontSizePt = size, ColorHex = null, Bold = false, Italic = false };

    private static TextMeasureRequest Request(string text, double width, double height, double size = 10, string family = "TestSans")
        => new()
        {
            Runs = [Run(text, size, family)],
            BoxWidthPt = width,
            BoxHeightPt = height,
            MinScale = 0.5
        };

    /// <summary>Space-separated 4-char words — wrap requires break opportunities.</summary>
    private static string Words(int count) => string.Join(' ', Enumerable.Repeat("aaaa", count));

    private static readonly DesignTokens Design = new()
    {
        Palette = new Dictionary<string, string> { ["primary"] = "#0B3D91" }
    };

    private LayoutResult Resolve(ContainerElement slide)
        => new LayoutResolver(_measurer).Resolve(new GenerationDocument
        {
            Version = "2.0",
            Design = Design,
            Slides = [slide]
        });

    private static ContainerElement RootRow(params GenElement[] children)
        => new() { Layout = new LayoutSpec { Mode = LayoutMode.Row }, Children = children };

    private static TextElement FixedText(string value, double w, double h, OverflowPolicy? overflow = null)
    {
        var text = new TextElement
        {
            Value = value,
            Font = "TestSans",
            FontSize = 10,
            Size = new SizeSpec { Width = w, Height = h }
        };
        return overflow is { } policy ? text with { Overflow = policy } : text;
    }

    // ------------------------------------------------------------- FitScale (direct, real metrics)

    [Fact]
    public void FitScale_TextFitsUnscaled_ReturnsOne()
    {
        var scale = _measurer.FitScale(Request("hello world", 100, 40));

        Assert.Equal(1.0, scale);
        Assert.Empty(_measurer.Warnings);
    }

    [Fact]
    public void FitScale_OversizedText_ShrinksWithinMinScaleBounds()
    {
        // 8 words wrap to 2 lines (24pt) in a 20pt box; first fit at 0.83 (2 × 9.96 = 19.92pt).
        var scale = _measurer.FitScale(Request(Words(8), 100, 20));

        Assert.Equal(0.83, scale);
    }

    [Fact]
    public void FitScale_NotFittingAtMinScale_ReturnsTrueScaleBelowMinScale()
    {
        // 20 words in a 15pt box: nothing fits down to MinScale 0.5; first fit at 0.41
        // (3 lines × 4.92 = 14.76pt). The honest sub-MinScale value lets the resolver warn
        // with real numbers instead of silently clipping at a bottomed-out 0.5 (§7.5).
        var scale = _measurer.FitScale(Request(Words(20), 100, 15));

        Assert.Equal(0.41, scale);
    }

    [Fact]
    public void FitScale_UnbreakableWordTooWide_ShrinksToFitWidth()
    {
        // 30 unbreakable chars (150pt) in a 100pt box: width binds at 150·s ≤ 100 → 0.66.
        var scale = _measurer.FitScale(Request(new string('a', 30), 100, 40));

        Assert.Equal(0.66, scale);
    }

    [Fact]
    public void FitScale_HardBreakForcesLines_ShrinksByHeight()
    {
        // Two forced lines (24pt) in a 20pt box → first fit at 0.83.
        var scale = _measurer.FitScale(Request("aaaa\naaaa", 100, 20));

        Assert.Equal(0.83, scale);
    }

    [Fact]
    public void FitScale_MixedRunSizes_DominantRunDrivesPitch()
    {
        // Pitch comes from the 20pt run (24pt), not the 10pt run (12pt): a 10pt pitch
        // would fit the 20pt box at scale 1, the dominant 20pt pitch shrinks to 0.83.
        var request = new TextMeasureRequest
        {
            Runs = [Run("aaaa "), Run("bbbb", size: 20)],
            BoxWidthPt = 200,
            BoxHeightPt = 20,
            MinScale = 0.5
        };

        var scale = _measurer.FitScale(request);

        Assert.Equal(0.83, scale);
    }

    [Fact]
    public void FitScale_EmptyRuns_FitAdequateBox()
    {
        var request = new TextMeasureRequest { Runs = [], BoxWidthPt = 100, BoxHeightPt = 40, MinScale = 0.5 };

        Assert.Equal(1.0, _measurer.FitScale(request));
    }

    [Fact]
    public void FitScale_ZeroHeightBox_NeverReturnsZero()
    {
        var scale = _measurer.FitScale(Request(Words(8), 100, 0));

        Assert.Equal(0.01, scale);
    }

    [Fact]
    public void FitScale_UnknownFamily_FallsBackAndWarns()
    {
        // Unmeasurable faces degrade to the 0.5em estimate (equal to TestSans advances
        // here) and surface a resolution warning — never silent.
        var scale = _measurer.FitScale(Request(Words(8), 100, 20, family: "NoSuchFamily"));

        Assert.Equal(0.83, scale);
        Assert.Contains(_measurer.Warnings, w => w.Contains("NoSuchFamily"));
    }

    [Fact]
    public void FitScale_ComplexScript_Warns()
    {
        _measurer.FitScale(Request("こんにちは世界", 100, 40));

        Assert.Contains(_measurer.Warnings, w => w.Contains("complex-script"));
    }

    [Fact]
    public void FitScale_WarningsReflectMostRecentCall()
    {
        _measurer.FitScale(Request("hello", 100, 40, family: "NoSuchFamily"));
        Assert.NotEmpty(_measurer.Warnings);

        _measurer.FitScale(Request("hello", 100, 40));
        Assert.Empty(_measurer.Warnings);
    }

    [Fact]
    public void FitScale_NullRequest_Throws()
        => Assert.Throws<ArgumentNullException>(() => _measurer.FitScale(null!));

    [Fact]
    public void FitScale_NegativeBoxDimension_Throws()
        => Assert.Throws<ArgumentOutOfRangeException>(() => _measurer.FitScale(Request("x", -1, 40)));

    [Fact]
    public void Ctor_NullCatalog_Throws()
        => Assert.Throws<ArgumentNullException>(() => new TextMeasure(null!));

    // ------------------------------------------------------------- resolver integration (§7.5)

    [Fact]
    public void Resolve_OversizedTextInFixedBox_ShrinksToFit_NeverSilentlyClips()
    {
        // A deliberately oversized text in a fixed box shrinks (≥ MinScale)
        // and the overflow is fully resolved — nothing left to flag, nothing clipped.
        var result = Resolve(RootRow(FixedText(Words(8), 100, 20)));

        var text = Assert.IsType<ResolvedText>(result.Slides[0].Root.Children[0]);
        Assert.Equal(0.83, text.FontScale);
        Assert.Empty(result.Warnings);
    }

    [Fact]
    public void Resolve_TextBeyondMinScale_AppliesTrueScaleAndWarns()
    {
        var result = Resolve(RootRow(FixedText(Words(20), 100, 15)));

        var text = Assert.IsType<ResolvedText>(result.Slides[0].Root.Children[0]);
        Assert.Equal(0.41, text.FontScale);
        var warning = Assert.Single(result.Warnings);
        Assert.Contains("slides[0].children[0]", warning);
        Assert.Contains("MinScale", warning);
    }

    [Fact]
    public void Resolve_ErrorPolicyOversizedText_ThrowsWithElementPath()
    {
        var ex = Assert.Throws<LayoutException>(
            () => Resolve(RootRow(FixedText(Words(8), 100, 20, OverflowPolicy.Error))));

        Assert.Equal("slides[0].children[0]", ex.Path);
        Assert.Contains("overflow policy", ex.Message);
        Assert.Contains("100", ex.Message);
    }

    [Fact]
    public void Resolve_FittingText_KeepsScaleOne()
    {
        var result = Resolve(RootRow(FixedText("hello", 100, 40)));

        var text = Assert.IsType<ResolvedText>(result.Slides[0].Root.Children[0]);
        Assert.Equal(1, text.FontScale);
        Assert.Empty(result.Warnings);
    }
}
