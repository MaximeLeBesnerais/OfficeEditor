using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PptxEditor.Core.Builders;
using PptxEditor.Core.Models;
using PptxEditor.Core.Services;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Wrap-simulation and normAutofit emulation tests. Fixture strategy: synthetic
/// decks (PresentationBuilder shell + raw OpenXML shape injection) and a synthetic
/// "TestSans" TTF whose metrics are exactly known (unitsPerEm 1000, uniform
/// advance 500 → char width = fontSize/2; default line factor 1.2). All catalog
/// lookups are hermetic (IncludeSystemFonts = false) so results are host-independent.
///
/// With sz=1000 (10pt): char width = 5pt, line pitch = 12pt.
/// </summary>
public sealed class TextFitServiceTests : IDisposable
{
    private const string Ns = "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" " +
                              "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                              "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";

    private readonly string _tempDir = SlideOpsTestHelpers.CreateTempDirectory(nameof(TextFitServiceTests));
    private readonly string _testSansPath;
    private readonly string _carlitoPath;
    private readonly string _ttcPath;

    public TextFitServiceTests()
    {
        _testSansPath = TextFitTestFont.WriteFont(_tempDir, "testsans.ttf", "TestSans");
        _carlitoPath = TextFitTestFont.WriteFont(_tempDir, "carlito.ttf", "Carlito");
        _ttcPath = TextFitTestFont.WriteTrueTypeCollection(_tempDir);
    }

    public void Dispose() => SlideOpsTestHelpers.BestEffortDelete(_tempDir);

    // ------------------------------------------------------------- helpers

    private static long Pt(double pt) => (long)Math.Round(pt * 12700);

    private static string ShapeXml(uint id, double widthPt, double heightPt, string bodyPrContent, string paragraphsXml,
        bool zeroInsets = true)
    {
        var insets = zeroInsets ? "lIns=\"0\" rIns=\"0\" tIns=\"0\" bIns=\"0\" " : "";
        return $"<p:sp {Ns}>" +
               $"<p:nvSpPr><p:cNvPr id=\"{id}\" name=\"Shape{id}\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>" +
               "<p:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/>" +
               $"<a:ext cx=\"{Pt(widthPt)}\" cy=\"{Pt(heightPt)}\"/></a:xfrm></p:spPr>" +
               $"<p:txBody><a:bodyPr {insets}>{bodyPrContent}</a:bodyPr><a:lstStyle/>{paragraphsXml}</p:txBody>" +
               "</p:sp>";
    }

    private static string ParaXml(string text, string pPrAttrs = "", string pPrInner = "",
        string rPrAttrs = "sz=\"1000\"", string typeface = "TestSans")
    {
        var pPr = string.IsNullOrEmpty(pPrAttrs) && string.IsNullOrEmpty(pPrInner)
            ? ""
            : $"<a:pPr {pPrAttrs}>{pPrInner}</a:pPr>";
        return $"<a:p>{pPr}<a:r><a:rPr lang=\"en-US\" {rPrAttrs}><a:latin typeface=\"{typeface}\"/></a:rPr>" +
               $"<a:t>{text}</a:t></a:r></a:p>";
    }

    private static string Chars(int count) => new('a', count);

    /// <summary>Space-separated words — wrap requires break opportunities (unbreakable
    /// words stay on one line and overflow horizontally, PowerPoint-like).</summary>
    private static string Words(int wordLength, int count)
        => string.Join(' ', Enumerable.Repeat(Chars(wordLength), count));

    private string CreateDeck(params string[] shapesXml)
    {
        var path = Path.Combine(_tempDir, $"{Guid.NewGuid():N}.pptx");
        using (var builder = PresentationBuilder.Create(path))
        {
            builder.AddSlide();
            var slidePart = ((SlideBuilder)builder.GetSlide(0)).SlidePart;
            var tree = slidePart.Slide!.CommonSlideData!.ShapeTree!;
            foreach (var xml in shapesXml)
                tree.AppendChild(new P.Shape(xml));
            slidePart.Slide.Save();
            builder.Save();
        }
        return path;
    }

    private TextFitService CreateService(PresentationDocument doc, params (string Family, string Path)[] fonts)
        => new(new FontMetricsCatalog(doc, new FontMetricsCatalogOptions
        {
            AdditionalFontPaths = fonts,
            IncludeSystemFonts = false
        }));

    private TextFitService CreateTestSansService(PresentationDocument doc)
        => CreateService(doc, ("TestSans", _testSansPath));

    private static SlidePart FirstSlidePart(PresentationDocument doc)
        => doc.PresentationPart!.SlideParts.First();

    private static string SlideShapeBodyPrXml(string path, uint id)
    {
        using var doc = PresentationDocument.Open(path, false);
        var tree = FirstSlidePart(doc).Slide!.CommonSlideData!.ShapeTree!;
        foreach (var shape in tree.ChildElements.OfType<P.Shape>())
        {
            var cNvPr = shape.NonVisualShapeProperties?.ChildElements.FirstOrDefault(c => c.LocalName == "cNvPr");
            if (cNvPr != null && Regex.IsMatch(cNvPr.OuterXml, $@"\bid\s*=\s*""{id}"""))
            {
                var bodyPr = shape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                return bodyPr?.OuterXml ?? "";
            }
        }
        throw new InvalidOperationException($"Shape {id} not found");
    }

    private static string AllMasterAndLayoutXml(string path)
    {
        using var doc = PresentationDocument.Open(path, false);
        var parts = doc.PresentationPart!.SlideMasterParts
            .SelectMany(m => new OpenXmlPart[] { m }.Concat(m.SlideLayoutParts));
        return string.Join("\n", parts.Select(p =>
        {
            using var reader = new System.IO.StreamReader(p.GetStream());
            return reader.ReadToEnd();
        }));
    }

    // ------------------------------------------------------------- wrap simulation

    [Fact]
    public void ExactlyFits_NoOverflow()
    {
        var path = CreateDeck(ShapeXml(2, 100, 12, "", ParaXml(Chars(20)))); // 20 × 5pt = 100pt, pitch 12pt

        using var doc = PresentationDocument.Open(path, true);
        var result = CreateTestSansService(doc).CheckShapeFit(FirstSlidePart(doc), 2);

        Assert.False(result.Overflow);
        Assert.Equal(100, result.ContentWidthPt);
        Assert.Equal(12, result.ContentHeightPt);
        Assert.Equal(100, result.BoxContentWidthPt);
        Assert.Equal(12, result.BoxContentHeightPt);
        Assert.Null(result.AppliedFontScale);
        Assert.Empty(result.Warnings);
    }

    [Fact]
    public void OnePointOverflow_WrapsToSecondLine()
    {
        var path = CreateDeck(ShapeXml(2, 100, 12, "", ParaXml(Words(10, 2)))); // 105pt → 2 lines × 12pt

        using var doc = PresentationDocument.Open(path, true);
        var result = CreateTestSansService(doc).CheckShapeFit(FirstSlidePart(doc), 2);

        Assert.True(result.Overflow);
        Assert.Equal(24, result.ContentHeightPt);
    }

    [Fact]
    public void EmptyText_OccupiesOneLineAtRunSize()
    {
        var path = CreateDeck(ShapeXml(2, 100, 12, "", ParaXml("")));

        using var doc = PresentationDocument.Open(path, true);
        var result = CreateTestSansService(doc).CheckShapeFit(FirstSlidePart(doc), 2);

        Assert.False(result.Overflow);
        Assert.Equal(12, result.ContentHeightPt); // empty run still sizes its line (endParaRPr behavior)
    }

    [Fact]
    public void MultiParagraph_WithSpacingBefore_Accumulates()
    {
        var paragraphs = ParaXml(Chars(5)) + ParaXml(Chars(5), pPrInner: "<a:spcBef><a:spcPts val=\"600\"/></a:spcBef>");
        var fitsPath = CreateDeck(ShapeXml(2, 100, 30, "", paragraphs));
        var tightPath = CreateDeck(ShapeXml(2, 100, 29.9, "", paragraphs));

        using (var doc = PresentationDocument.Open(fitsPath, true))
        {
            var result = CreateTestSansService(doc).CheckShapeFit(FirstSlidePart(doc), 2);
            Assert.False(result.Overflow);
            Assert.Equal(30, result.ContentHeightPt); // 12 + 6 + 12
        }
        using (var doc = PresentationDocument.Open(tightPath, true))
        {
            Assert.True(CreateTestSansService(doc).CheckShapeFit(FirstSlidePart(doc), 2).Overflow);
        }
    }

    [Fact]
    public void Bullet_GutterAbsorbsBulletChar()
    {
        const string pPr = "marL=\"457200\" indent=\"-457200\""; // 36pt margin + 36pt hanging gutter → 64pt line width
        var fits = CreateDeck(ShapeXml(2, 100, 12, "", ParaXml("aaaaaa aaaaa", pPr, "<a:buChar char=\"•\"/>")));   // 60pt ≤ 64
        var wraps = CreateDeck(ShapeXml(2, 100, 12, "", ParaXml("aaaaaa aaaaaa", pPr, "<a:buChar char=\"•\"/>"))); // 65pt > 64 → 2 lines

        using (var doc = PresentationDocument.Open(fits, true))
        {
            var result = CreateTestSansService(doc).CheckShapeFit(FirstSlidePart(doc), 2);
            Assert.False(result.Overflow);
            Assert.Equal(60, result.ContentWidthPt);
        }
        using (var doc = PresentationDocument.Open(wraps, true))
        {
            Assert.True(CreateTestSansService(doc).CheckShapeFit(FirstSlidePart(doc), 2).Overflow);
        }
    }

    [Fact]
    public void Bullet_WiderThanGutter_PushesFirstLineRight()
    {
        // 2pt gutter, bullet 5pt → first line = 100 − 2 − (5−2) = 95pt. The same
        // 100pt text without a bullet gets the full hanging indent back (100pt) and fits.
        const string pPr = "marL=\"25400\" indent=\"-25400\"";
        var text = Words(6, 3); // 100pt
        var withBullet = CreateDeck(ShapeXml(2, 100, 12, "", ParaXml(text, pPr, "<a:buChar char=\"•\"/>")));
        var noBullet = CreateDeck(ShapeXml(2, 100, 12, "", ParaXml(text, pPr, "<a:buNone/>")));

        using (var doc = PresentationDocument.Open(withBullet, true))
            Assert.True(CreateTestSansService(doc).CheckShapeFit(FirstSlidePart(doc), 2).Overflow);
        using (var doc = PresentationDocument.Open(noBullet, true))
            Assert.False(CreateTestSansService(doc).CheckShapeFit(FirstSlidePart(doc), 2).Overflow);
    }

    [Fact]
    public void DefaultInsets_AppliedWhenAbsent()
    {
        // No inset attributes → OOXML defaults 91440/91440/45720/45720 EMU
        // → content 85.6 × 4.8pt; one 12pt line no longer fits vertically.
        var path = CreateDeck(ShapeXml(2, 100, 12, "", ParaXml(Chars(10)), zeroInsets: false));

        using var doc = PresentationDocument.Open(path, true);
        var result = CreateTestSansService(doc).CheckShapeFit(FirstSlidePart(doc), 2);

        Assert.True(result.Overflow);
        Assert.Equal(85.6, result.BoxContentWidthPt);
        Assert.Equal(4.8, result.BoxContentHeightPt);
    }

    [Fact]
    public void ExplicitLineSpacing_Points_OverridesDefaultPitch()
    {
        var para = ParaXml(Words(10, 2), pPrInner: "<a:lnSpc><a:spcPts val=\"2000\"/></a:lnSpc>"); // 2 lines × 20pt
        var path = CreateDeck(ShapeXml(2, 100, 40, "", para));

        using var doc = PresentationDocument.Open(path, true);
        var result = CreateTestSansService(doc).CheckShapeFit(FirstSlidePart(doc), 2);

        Assert.False(result.Overflow);
        Assert.Equal(40, result.ContentHeightPt);
    }

    // ------------------------------------------------------------- auto-shrink

    [Fact]
    public void AutoShrink_FontScaleSearch_FindsExactScale()
    {
        // Four 50pt words = 4 lines at 100% (48pt > 12pt box); without spacing
        // reduction the first fitting scale is exactly 0.5 (3 words/line → 2 lines of 6pt).
        var path = CreateDeck(ShapeXml(2, 100, 12, "<a:normAutofit/>", ParaXml(Words(10, 4))));

        using var doc = PresentationDocument.Open(path, true);
        var result = CreateTestSansService(doc).ReplaceTextWithFitCheck(
            FirstSlidePart(doc), 2, Words(10, 4), new FitCheckOptions { AllowLineSpacingReduction = false });

        Assert.False(result.Overflow);
        Assert.Equal(0.5, result.AppliedFontScale);
        Assert.Null(result.AppliedLineSpacingReduction);
    }

    [Fact]
    public void AutoShrink_TriesLineSpacingReductionBeforeFontScale()
    {
        // One 12pt line in an 11pt box: 10% spacing reduction alone fits (10.8 ≤ 11).
        var path = CreateDeck(ShapeXml(2, 100, 11, "<a:normAutofit/>", ParaXml(Chars(20))));

        using var doc = PresentationDocument.Open(path, true);
        var result = CreateTestSansService(doc).CheckShapeFit(FirstSlidePart(doc), 2);

        Assert.False(result.Overflow);
        Assert.Null(result.AppliedFontScale);
        Assert.Equal(0.1, result.AppliedLineSpacingReduction);
    }

    [Fact]
    public void AutoShrink_MinScaleFloor_IsFlagged()
    {
        var path = CreateDeck(ShapeXml(2, 100, 12, "<a:normAutofit/>", ParaXml(Words(10, 40))));

        using var doc = PresentationDocument.Open(path, true);
        var result = CreateTestSansService(doc).ReplaceTextWithFitCheck(FirstSlidePart(doc), 2, Words(10, 40), new FitCheckOptions());

        Assert.True(result.Overflow);
        Assert.Equal(0.5, result.AppliedFontScale); // floor persisted, PowerPoint-like
        Assert.Contains(result.Warnings, w => w.Contains("minimum scale"));
    }

    [Fact]
    public void AutoShrink_MinScale_IsConfigurable()
    {
        var path = CreateDeck(ShapeXml(2, 100, 12, "<a:normAutofit/>", ParaXml(Words(10, 4))));

        using var doc = PresentationDocument.Open(path, true);
        var result = CreateTestSansService(doc).ReplaceTextWithFitCheck(
            FirstSlidePart(doc), 2, Words(10, 4), new FitCheckOptions { MinScale = 0.8, AllowLineSpacingReduction = false });

        Assert.True(result.Overflow);
        Assert.Equal(0.8, result.AppliedFontScale);
        Assert.Contains(result.Warnings, w => w.Contains("minimum scale"));
    }

    [Fact]
    public void AutoShrink_ScaleSearch_IsMonotonicInContentSize()
    {
        // 30pt-tall box, no spacing reduction: longer text → equal-or-smaller applied scale.
        var path = CreateDeck(ShapeXml(2, 100, 30, "<a:normAutofit/>", ParaXml(Words(10, 4))));

        using var doc = PresentationDocument.Open(path, true);
        var service = CreateTestSansService(doc);
        var options = new FitCheckOptions { AllowLineSpacingReduction = false };

        var shortResult = service.ReplaceTextWithFitCheck(FirstSlidePart(doc), 2, Words(10, 4), options);
        var midResult = service.ReplaceTextWithFitCheck(FirstSlidePart(doc), 2, Words(10, 6), options);
        var longResult = service.ReplaceTextWithFitCheck(FirstSlidePart(doc), 2, Words(10, 8), options);

        Assert.Equal(0.95, shortResult.AppliedFontScale);
        Assert.Equal(0.83, midResult.AppliedFontScale);
        Assert.Equal(0.62, longResult.AppliedFontScale);
        Assert.True(shortResult.AppliedFontScale >= midResult.AppliedFontScale);
        Assert.True(midResult.AppliedFontScale >= longResult.AppliedFontScale);
    }

    // ------------------------------------------------------------- warnings

    [Fact]
    public void Substitution_SurfacesNotMetricCompatibleWarning()
    {
        var path = CreateDeck(ShapeXml(2, 100, 12, "", ParaXml(Chars(5), typeface: "Aptos")));

        using var doc = PresentationDocument.Open(path, true);
        var result = CreateService(doc, ("Carlito", _carlitoPath)).CheckShapeFit(FirstSlidePart(doc), 2);

        Assert.Contains(result.Warnings, w => w.Contains("Aptos") && w.Contains("not metric-compatible"));
    }

    [Fact]
    public void TrueTypeCollection_SurfacesWarning_AndStillMeasuresWithFallback()
    {
        var path = CreateDeck(ShapeXml(2, 100, 12, "", ParaXml(Chars(5), typeface: "TtcFont")));

        using var doc = PresentationDocument.Open(path, true);
        var result = CreateService(doc, ("TtcFont", _ttcPath)).CheckShapeFit(FirstSlidePart(doc), 2);

        Assert.Contains(result.Warnings, w => w.Contains("TrueType Collection"));
        Assert.True(result.ContentWidthPt > 0); // 0.5em fallback, never silently zero
    }

    [Fact]
    public void ComplexScript_SurfacesShapingWarning()
    {
        var path = CreateDeck(ShapeXml(2, 100, 50, "", ParaXml("日本語テスト")));

        using var doc = PresentationDocument.Open(path, true);
        var result = CreateTestSansService(doc).CheckShapeFit(FirstSlidePart(doc), 2);

        Assert.Contains(result.Warnings, w => w.Contains("shaping/kerning"));
    }

    [Fact]
    public void UnresolvableFont_SurfacesUnmeasurableWarning()
    {
        var path = CreateDeck(ShapeXml(2, 100, 12, "", ParaXml(Chars(5), typeface: "NoSuchFont")));

        using var doc = PresentationDocument.Open(path, true);
        var result = CreateTestSansService(doc).CheckShapeFit(FirstSlidePart(doc), 2);

        Assert.Contains(result.Warnings, w => w.Contains("unmeasurable"));
        Assert.True(result.ContentWidthPt > 0);
    }

    [Fact]
    public void UnbreakableWord_WarnsAboutHorizontalOverflow()
    {
        var path = CreateDeck(ShapeXml(2, 100, 50, "", ParaXml(Chars(30)))); // 150pt word in 100pt box

        using var doc = PresentationDocument.Open(path, true);
        var result = CreateTestSansService(doc).CheckShapeFit(FirstSlidePart(doc), 2);

        Assert.Contains(result.Warnings, w => w.Contains("overflow horizontally"));
    }

    // ------------------------------------------------------------- persistence

    [Fact]
    public void WriteBack_PersistsNormAutofitOnSlideShape_MasterUntouched()
    {
        var path = CreateDeck(ShapeXml(2, 100, 12, "<a:normAutofit/>", ParaXml(Chars(5))));

        using (var doc = PresentationDocument.Open(path, true))
        {
            var result = CreateTestSansService(doc).ReplaceTextWithFitCheck(
                FirstSlidePart(doc), 2, Words(10, 4), new FitCheckOptions { AllowLineSpacingReduction = false });
            Assert.False(result.Overflow);
            Assert.Equal(0.5, result.AppliedFontScale);
        }

        var bodyPrXml = SlideShapeBodyPrXml(path, 2);
        var autofit = Regex.Match(bodyPrXml, @"<a:normAutofit\b[^>]*>");
        Assert.True(autofit.Success, "normAutofit missing on slide shape bodyPr");
        Assert.Contains("fontScale=\"50000\"", autofit.Value);
        Assert.DoesNotContain("lnSpcReduction", autofit.Value);
        Assert.DoesNotContain("normAutofit", AllMasterAndLayoutXml(path));
    }

    [Fact]
    public void WriteBack_AutoShrinkOption_AddsAutofitToNonAutofitShape()
    {
        var path = CreateDeck(ShapeXml(2, 100, 12, "", ParaXml(Chars(5))));

        using (var doc = PresentationDocument.Open(path, true))
        {
            var result = CreateTestSansService(doc).ReplaceTextWithFitCheck(
                FirstSlidePart(doc), 2, Words(10, 4),
                new FitCheckOptions { AutoShrink = true, AllowLineSpacingReduction = false });
            Assert.False(result.Overflow);
        }

        Assert.Contains("fontScale=\"50000\"", SlideShapeBodyPrXml(path, 2));
        Assert.DoesNotContain("normAutofit", AllMasterAndLayoutXml(path));
    }

    [Fact]
    public void WriteBack_StaleScale_ResetWhenNewTextFits()
    {
        var path = CreateDeck(ShapeXml(2, 100, 12, "<a:normAutofit fontScale=\"50000\"/>", ParaXml(Chars(40))));

        using (var doc = PresentationDocument.Open(path, true))
        {
            var result = CreateTestSansService(doc).ReplaceTextWithFitCheck(FirstSlidePart(doc), 2, Chars(5), new FitCheckOptions());
            Assert.False(result.Overflow);
            Assert.Null(result.AppliedFontScale);
        }

        var bodyPrXml = SlideShapeBodyPrXml(path, 2);
        Assert.Contains("normAutofit", bodyPrXml);          // element kept
        Assert.DoesNotContain("fontScale", bodyPrXml);      // stale scale cleared
    }

    [Fact]
    public void CheckShapeFit_ComputesScaleButNeverPersists()
    {
        var path = CreateDeck(ShapeXml(2, 100, 12, "<a:normAutofit/>", ParaXml(Words(10, 4))));

        using (var doc = PresentationDocument.Open(path, true))
        {
            var result = CreateTestSansService(doc).CheckShapeFit(FirstSlidePart(doc), 2);
            // default options allow spacing reduction: 20% + scale 0.62 → 2 lines × 11.9pt ≤ 12pt
            Assert.Equal(0.62, result.AppliedFontScale);
            Assert.Equal(0.2, result.AppliedLineSpacingReduction);
        }

        Assert.DoesNotContain("fontScale", SlideShapeBodyPrXml(path, 2)); // … not persisted
    }

    [Fact]
    public void NoAutoFit_NoAutoShrink_NothingPersisted()
    {
        var path = CreateDeck(ShapeXml(2, 100, 12, "", ParaXml(Chars(5))));

        using (var doc = PresentationDocument.Open(path, true))
        {
            var result = CreateTestSansService(doc).ReplaceTextWithFitCheck(FirstSlidePart(doc), 2, Words(10, 4), new FitCheckOptions());
            Assert.True(result.Overflow);
            Assert.Null(result.AppliedFontScale);
        }

        Assert.DoesNotContain("normAutofit", SlideShapeBodyPrXml(path, 2));
    }

    // ------------------------------------------------------------- error paths

    [Fact]
    public void MissingShape_Check_ReturnsWarningInsteadOfThrowing()
    {
        var path = CreateDeck(ShapeXml(2, 100, 12, "", ParaXml(Chars(5))));

        using var doc = PresentationDocument.Open(path, true);
        var result = CreateTestSansService(doc).CheckShapeFit(FirstSlidePart(doc), 999);

        Assert.False(result.Overflow);
        Assert.Contains(result.Warnings, w => w.Contains("No shape with id 999"));
    }

    [Fact]
    public void MissingShape_Replace_ReportsFailure()
    {
        var path = CreateDeck(ShapeXml(2, 100, 12, "", ParaXml(Chars(5))));

        using var doc = PresentationDocument.Open(path, true);
        var result = CreateTestSansService(doc).ReplaceTextWithFitCheck(FirstSlidePart(doc), 999, "text", new FitCheckOptions());

        Assert.False(result.Replaced);
        Assert.NotNull(result.ReplaceError);
        Assert.Contains("999", result.ReplaceError);
    }
}
