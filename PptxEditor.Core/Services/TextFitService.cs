using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PptxEditor.Core.Converters;
using PptxEditor.Core.Models;

namespace PptxEditor.Core.Services;

/// <summary>
/// Fit-checked text measurement/replacement for slide shapes (Phase-4 F6 differentiator).
///
/// Re-resolves a shape's effective text model (cascaded bodyPr slide ← layout ← master,
/// OOXML inset defaults 91440/91440/45720/45720 EMU, paragraph marL/indent/spcBef/spcAft/lnSpc,
/// per-run font size/bold), simulates greedy line wrapping with advance-width metrics, and
/// reports whether the content overflows the box. When the shape carries
/// <c>&lt;a:normAutofit&gt;</c> (or the caller passes <c>autoShrink</c>), it emulates
/// PowerPoint's shrink: line-spacing reduction 10%/20% first, then font scale 99%→MinScale
/// at 1% steps, and persists the result on the <b>slide shape's</b> bodyPr only
/// (never master/layout; never any global autofit).
///
/// Re-baseline note: any <c>fontScale</c> already stored on a shape reflects
/// <i>previous</i> text. Measurement always starts at 100%; the stored value is
/// ignored when deciding (PowerPoint behaves the same on edit), and a stale scale
/// is reset when the new text fits unscaled.
///
/// Accuracy: raw advance-width accumulation without a shaping/kerning engine —
/// Latin prose error ~1–3% of line width (fine for a boolean overflow verdict and
/// 1% scale steps); complex scripts (CJK/Arabic/Hebrew/Indic/Thai) are materially
/// wrong and always surface a warning. Missing metrics degrade to a 0.5em-per-rune
/// estimate with an "unmeasurable" warning; nothing fails silently.
/// </summary>
public sealed class TextFitService
{
    // OOXML (ECMA-376) bodyPr inset defaults: 0.1" left/right, 0.05" top/bottom.
    private const long DefaultHorizontalInsetEmu = 91440;
    private const long DefaultVerticalInsetEmu = 45720;
    private const double EmuPerPoint = 12700.0;

    // Attribute reads follow the regex-on-OuterXml pattern (SDK
    // GetAttribute crashes on unreliable attributes such as marL/indent/char/idx).
    private static readonly Regex IdAttributePattern = new(@"\bid\s*=\s*""([^""]*)""", RegexOptions.Compiled);

    private readonly FontMetricsCatalog _catalog;

    /// <summary>Creates a service with a catalog bound to the given deck (embedded fonts honored).</summary>
    public TextFitService(PresentationDocument? document, FontMetricsCatalogOptions? catalogOptions = null)
        : this(new FontMetricsCatalog(document, catalogOptions))
    {
    }

    /// <summary>Creates a service over a shared (cached) catalog — preferred for repeated checks on one deck.</summary>
    public TextFitService(FontMetricsCatalog catalog)
    {
        ArgumentNullException.ThrowIfNull(catalog);
        _catalog = catalog;
    }

    /// <summary>
    /// Measures the shape's current text against its box. Read-only: computes the
    /// scale that would be applied when the shape has normAutofit, but persists nothing.
    /// </summary>
    public FitCheckResult CheckShapeFit(SlidePart slidePart, uint elementId)
    {
        ArgumentNullException.ThrowIfNull(slidePart);
        return MeasureShape(slidePart, elementId, new FitCheckOptions(), persist: false);
    }

    /// <summary>
    /// Replaces the shape's text via the style-preserving <see cref="PptxElementReplacer"/>,
    /// then measures the new content; on overflow with normAutofit/autoShrink, emulates
    /// PowerPoint shrink and (when <see cref="FitCheckOptions.PersistAutoFit"/>) writes
    /// <c>&lt;a:normAutofit fontScale="…" lnSpcReduction="…"/&gt;</c> on the slide shape only.
    /// </summary>
    public FitCheckResult ReplaceTextWithFitCheck(SlidePart slidePart, uint elementId, string text, FitCheckOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(slidePart);
        options ??= new FitCheckOptions();

        var replaceResult = new PptxElementReplacer().ReplaceText(slidePart, elementId, text);
        if (!replaceResult.Success)
        {
            return new FitCheckResult
            {
                Replaced = false,
                ReplaceError = replaceResult.Error,
                Warnings = [replaceResult.Error ?? "Replacement failed."]
            };
        }

        return MeasureShape(slidePart, elementId, options, persist: options.PersistAutoFit);
    }

    // ---------------------------------------------------------------- pipeline

    private FitCheckResult MeasureShape(SlidePart slidePart, uint elementId, FitCheckOptions options, bool persist)
    {
        var warnings = new List<string>();

        var shapeTree = slidePart.Slide?.CommonSlideData?.ShapeTree;
        if (shapeTree == null)
            return NoMeasurement(warnings, "Slide has no shape tree.");

        var matches = FindShapesById(shapeTree, elementId);
        if (matches.Count == 0)
            return NoMeasurement(warnings, $"No shape with id {elementId} found on the slide.");
        if (matches.Count > 1)
            return NoMeasurement(warnings, $"Multiple shapes with id {elementId} found on the slide; refusing to measure.");

        var shape = matches[0];
        var textBody = shape.TextBody;
        if (textBody == null)
            return NoMeasurement(warnings, $"Shape with id {elementId} has no text body.");

        var document = slidePart.OpenXmlPackage as PresentationDocument;
        var styleResolver = document != null ? new StyleResolver(document, slidePart) : null;
        var (placeholderIdx, placeholderType) = GetPlaceholderInfo(shape);

        // --- cascaded bodyPr (slide ← layout ← master element copy, converter :1955-2024 pattern)
        var bodyPr = CascadeBodyPr(textBody, styleResolver, placeholderIdx, placeholderType);
        var lIns = EmuToPt(bodyPr?.LeftInset?.Value ?? DefaultHorizontalInsetEmu);
        var rIns = EmuToPt(bodyPr?.RightInset?.Value ?? DefaultHorizontalInsetEmu);
        var tIns = EmuToPt(bodyPr?.TopInset?.Value ?? DefaultVerticalInsetEmu);
        var bIns = EmuToPt(bodyPr?.BottomInset?.Value ?? DefaultVerticalInsetEmu);
        var hasNormAutofit = bodyPr?.Elements<Drawing.NormalAutoFit>().FirstOrDefault() != null;

        // --- box extents
        var extents = shape.ShapeProperties?.Transform2D?.Extents;
        if (extents?.Cx?.Value == null || extents?.Cy?.Value == null)
            return NoMeasurement(warnings, $"Shape with id {elementId} has no extents; box is unmeasurable.");

        var boxContentW = Math.Max(0, EmuToPt(extents.Cx.Value) - lIns - rIns);
        var boxContentH = Math.Max(0, EmuToPt(extents.Cy.Value) - tIns - bIns);

        // --- effective text model
        var bodyLstStyle = textBody.ChildElements.FirstOrDefault(e => e.LocalName == "lstStyle");
        var paragraphs = new List<MeasuredParagraph>();
        foreach (var paragraph in textBody.Elements<Drawing.Paragraph>())
        {
            paragraphs.Add(ResolveParagraph(paragraph, bodyLstStyle, styleResolver, placeholderIdx, placeholderType, warnings));
        }
        if (paragraphs.Count == 0)
        {
            paragraphs.Add(MeasuredParagraph.Empty);
        }

        // --- baseline measurement at 100% (re-baseline: stored fontScale reflects OLD text)
        var baseline = MeasureContent(paragraphs, boxContentW, warnings);
        var overflow = baseline.Height > boxContentH + 1e-6;
        var contentWidthPt = baseline.Width;
        var contentHeightPt = baseline.Height;

        if (baseline.Width > boxContentW + 0.5)
        {
            warnings.Add($"Content width {baseline.Width:F1}pt exceeds the box content width {boxContentW:F1}pt; PowerPoint will overflow horizontally (unbreakable word or insets too large).");
        }

        double? appliedScale = null;
        double? appliedSpacingReduction = null;
        var shrinkInPlay = hasNormAutofit || options.AutoShrink;

        if (overflow && shrinkInPlay)
        {
            (appliedScale, appliedSpacingReduction, overflow) = SearchShrink(
                paragraphs, boxContentW, boxContentH, options, warnings);
        }

        // --- persist on the SLIDE SHAPE's bodyPr only (style preservation; rule 5)
        if (persist && shrinkInPlay)
        {
            if (appliedScale != null || appliedSpacingReduction != null)
            {
                PersistNormAutofit(shape, appliedScale ?? 1.0, appliedSpacingReduction ?? 0.0);
            }
            else if (!overflow && hasNormAutofit)
            {
                // Stale-scale reset: new text fits at 100%, so drop fontScale/lnSpcReduction
                // stored for previous text (keeps the normAutofit element itself).
                PersistNormAutofit(shape, 1.0, 0.0);
            }
        }

        return new FitCheckResult
        {
            Overflow = overflow,
            ContentWidthPt = Math.Round(contentWidthPt, 2),
            ContentHeightPt = Math.Round(contentHeightPt, 2),
            BoxContentWidthPt = Math.Round(boxContentW, 2),
            BoxContentHeightPt = Math.Round(boxContentH, 2),
            AppliedFontScale = appliedScale,
            AppliedLineSpacingReduction = appliedSpacingReduction,
            Warnings = Deduplicate(warnings)
        };
    }

    private static FitCheckResult NoMeasurement(List<string> warnings, string message)
    {
        warnings.Add(message);
        return new FitCheckResult { Warnings = warnings };
    }

    // -------------------------------------------------------- shrink search

    /// <summary>
    /// PowerPoint-order shrink: line-spacing reduction (10%, 20%) at full scale first,
    /// then fontScale 99%→MinScale at 1% steps (keeping the max spacing reduction).
    /// Returns the applied (scale, spacingReduction, stillOverflowing). Even when nothing
    /// fits, the floor values are returned so the caller can persist/flag them.
    /// </summary>
    private (double? Scale, double? SpacingReduction, bool Overflow) SearchShrink(
        List<MeasuredParagraph> paragraphs, double boxContentW, double boxContentH,
        FitCheckOptions options, List<string> warnings)
    {
        var maxReduction = options.AllowLineSpacingReduction ? 0.20 : 0.0;

        if (options.AllowLineSpacingReduction)
        {
            foreach (var r in new[] { 0.10, 0.20 })
            {
                if (MeasureContent(paragraphs, boxContentW, warnings, fontScale: 1.0, spacingReduction: r).Height <= boxContentH + 1e-6)
                    return (null, r, false);
            }
        }

        var minScale = Math.Clamp(options.MinScale, 0.01, 1.0);
        for (var s = 0.99; s >= minScale - 1e-9; s -= 0.01)
        {
            if (MeasureContent(paragraphs, boxContentW, warnings, fontScale: s, spacingReduction: maxReduction).Height <= boxContentH + 1e-6)
                return (Math.Round(s, 2), maxReduction > 0 ? maxReduction : null, false);
        }

        warnings.Add($"Auto-shrink reached the minimum scale ({minScale:P0}) without fitting; text still overflows.");
        return (minScale, maxReduction > 0 ? maxReduction : null, true);
    }

    // -------------------------------------------------------- persistence

    /// <summary>
    /// Sets/replaces &lt;a:normAutofit&gt; on the slide shape's own bodyPr.
    /// fontScale=1.0 / reduction=0 are written as absent attributes (PowerPoint's
    /// plain &lt;a:normAutofit/&gt;). Typed SDK writes; regex reads only (rule 1).
    /// </summary>
    private static void PersistNormAutofit(P.Shape shape, double fontScale, double spacingReduction)
    {
        var textBody = shape.TextBody!;
        var bodyPr = textBody.Elements<Drawing.BodyProperties>().FirstOrDefault();
        if (bodyPr == null)
        {
            bodyPr = new Drawing.BodyProperties();
            textBody.InsertAt(bodyPr, 0);
        }

        foreach (var child in bodyPr.ChildElements.ToList())
        {
            if (child is Drawing.NoAutoFit or Drawing.NormalAutoFit or Drawing.ShapeAutoFit)
                child.Remove();
        }

        var normAutofit = new Drawing.NormalAutoFit();
        if (fontScale < 0.9999)
            normAutofit.FontScale = (int)Math.Round(fontScale * 100000);
        if (spacingReduction > 0.0001)
            normAutofit.LineSpaceReduction = (int)Math.Round(spacingReduction * 100000);

        // CT_TextBodyProperties sequence: prstTxWarp?, autofit?, scene3d?, sp3d|flatTx?, extLst?
        var prstTxWarp = bodyPr.Elements<Drawing.PresetTextWarp>().FirstOrDefault();
        if (prstTxWarp != null)
            bodyPr.InsertAfter(normAutofit, prstTxWarp);
        else
            bodyPr.InsertAt(normAutofit, 0);
    }

    // -------------------------------------------------------- text model

    private sealed record MeasuredRun(
        string Text, string Family, bool Bold, bool Italic, double FontSize, TypstFontMetrics? Metrics);

    private sealed record MeasuredParagraph(
        List<MeasuredRun> Runs,
        double MarLPt,
        double IndentPt,
        bool HasBullet,
        string? BulletChar,
        bool HasAutoNumber,
        double DominantFontSize,
        TypstFontMetrics? DominantMetrics,
        string DominantFamily,
        bool DominantBold,
        double? LineSpacing,
        bool LineSpacingIsPct,
        double? SpaceBefore,
        bool SpaceBeforeIsPct,
        double? SpaceAfter,
        bool SpaceAfterIsPct,
        bool HasComplexRunes)
    {
        public static readonly MeasuredParagraph Empty = new(
            [], 0, 0, false, null, false, 18.0, null, "Arial", false,
            null, false, null, false, null, false, false);
    }

    private MeasuredParagraph ResolveParagraph(
        Drawing.Paragraph paragraph,
        OpenXmlElement? bodyLstStyle,
        StyleResolver? styleResolver,
        int? placeholderIdx,
        PlaceholderValues? placeholderType,
        List<string> warnings)
    {
        var pPr = paragraph.Elements<Drawing.ParagraphProperties>().FirstOrDefault();
        var level = pPr?.Level?.Value ?? 0;

        var runs = paragraph.ChildElements
            .Where(e => e is Drawing.Run or Drawing.Field)
            .Select(e => e switch
            {
                Drawing.Run r => (Text: r.Text?.Text, Fmt: ExtractRunFormatting(r.RunProperties)),
                Drawing.Field f => (Text: f.Text?.Text, Fmt: ExtractRunFormatting(f.RunProperties)),
                _ => (Text: (string?)null, Fmt: RawFormatting.Default)
            })
            // Empty-text runs are kept: they carry the rPr that sizes an empty
            // paragraph's line (PowerPoint endParaRPr behavior), contributing no tokens.
            .Select(r => (Text: r.Text ?? string.Empty, r.Fmt))
            .ToList();

        // Paragraph defaults from the first run only when formatting is uniform
        // (converter pattern: mixed formatting must not let the first run dictate).
        var mixed = runs.Count > 1 && runs.Any(r => r.Fmt != runs[0].Fmt);
        var paragraphFmt = ResolveParagraphFormatting(
            mixed || runs.Count == 0 ? null : runs[0].Fmt,
            pPr, bodyLstStyle, level, styleResolver, placeholderIdx, placeholderType);

        // Master txStyles defaults (converter ExtractText applies them last, by sentinel).
        var txDefault = styleResolver?.GetDefaultTextStyle(placeholderType);
        if (txDefault != null)
            paragraphFmt = MergeDefaultTextStyle(paragraphFmt, txDefault);

        var measuredRuns = new List<MeasuredRun>();
        var hasComplex = false;
        foreach (var (text, rawFmt) in runs)
        {
            var fmt = MergeRunFormatting(paragraphFmt, rawFmt);
            var resolution = _catalog.Resolve(fmt.Family, fmt.Bold, fmt.Italic);
            foreach (var w in resolution.Warnings)
                warnings.Add(w);
            if (ContainsComplexScript(text))
                hasComplex = true;
            measuredRuns.Add(new MeasuredRun(text, resolution.ResolvedFamily, fmt.Bold, fmt.Italic, fmt.FontSize, resolution.Metrics));
        }

        if (hasComplex)
        {
            warnings.Add("Text contains complex-script runes (CJK/Arabic/Hebrew/Indic/Thai); measurement uses raw advance widths without a shaping/kerning engine and may be materially wrong.");
        }

        // Dominant (tallest) run drives the line pitch; paragraph default when empty.
        var dominant = measuredRuns.Count > 0
            ? measuredRuns.OrderByDescending(r => r.FontSize).First()
            : null;
        var dominantSize = dominant?.FontSize ?? paragraphFmt.FontSize;
        TypstFontMetrics? dominantMetrics = dominant?.Metrics;
        var dominantFamily = dominant?.Family ?? paragraphFmt.Family;
        var dominantBold = dominant?.Bold ?? paragraphFmt.Bold;
        if (dominant == null)
        {
            var res = _catalog.Resolve(dominantFamily, dominantBold, paragraphFmt.Italic);
            foreach (var w in res.Warnings)
                warnings.Add(w);
            dominantMetrics = res.Metrics;
            dominantFamily = res.ResolvedFamily;
        }

        var (marL, indent) = ExtractMarginAndIndent(pPr);
        var (bulletChar, autoNumber, hasBullet) = ResolveBullet(pPr, bodyLstStyle, level, styleResolver, placeholderIdx, placeholderType);
        var (lineSpacing, lineSpacingIsPct) = ResolveLineSpacing(pPr, bodyLstStyle, level, styleResolver, placeholderIdx, placeholderType);
        var (bef, befIsPct) = ResolveSpacing("spcBef", pPr, bodyLstStyle, level, styleResolver, placeholderIdx, placeholderType);
        var (aft, aftIsPct) = ResolveSpacing("spcAft", pPr, bodyLstStyle, level, styleResolver, placeholderIdx, placeholderType);

        return new MeasuredParagraph(
            measuredRuns, marL, indent, hasBullet, bulletChar, autoNumber != null,
            dominantSize, dominantMetrics, dominantFamily, dominantBold,
            lineSpacing, lineSpacingIsPct, bef, befIsPct, aft, aftIsPct, hasComplex);
    }

    // -------------------------------------------------------- measurement

    private sealed record Token(string Text, bool IsBreak, bool IsHardBreak, double Width);

    /// <summary>Greedy wrap + height at (fontScale, spacingReduction). Widths scale linearly with fontScale.</summary>
    private (double Width, double Height) MeasureContent(
        List<MeasuredParagraph> paragraphs, double boxContentW, List<string> warnings,
        double fontScale = 1.0, double spacingReduction = 0.0)
    {
        var totalHeight = 0.0;
        var maxLineWidth = 0.0;
        var missingGlyphFaces = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var para in paragraphs)
        {
            var contentWidth = Math.Max(1.0, boxContentW - para.MarLPt);
            var firstLineWidth = contentWidth - Math.Max(0, para.IndentPt);
            if (para.IndentPt < 0 && !para.HasBullet)
            {
                firstLineWidth = contentWidth - para.IndentPt; // hanging without bullet: first line starts left of marL
            }

            var dominantSize = para.DominantFontSize * fontScale;

            if (para.HasBullet && para.IndentPt < 0)
            {
                // Bullet sits in the hanging gutter; only overflow beyond the gutter
                // pushes the first text line right (PowerPoint gutter-jump behavior).
                var bulletText = para.BulletChar ?? (para.HasAutoNumber ? "99." : null);
                var bulletWidth = bulletText != null
                    ? MeasureString(bulletText, para.DominantMetrics, dominantSize, para.DominantFamily, missingGlyphFaces)
                    : 0;
                firstLineWidth = contentWidth - Math.Max(0, bulletWidth - Math.Abs(para.IndentPt));
            }

            var (lineCount, widest) = WrapParagraph(para, firstLineWidth, contentWidth, fontScale, missingGlyphFaces);
            maxLineWidth = Math.Max(maxLineWidth, widest);

            // Pitch: lnSpc pct of dominant size, absolute spcPts, or metrics default
            // (TypoAsc − TypoDesc + TypoLineGap) — ≈1.2 for typical text faces.
            double pitch;
            if (para.LineSpacing is { } ls)
            {
                pitch = para.LineSpacingIsPct ? dominantSize * ls : ls * fontScale;
            }
            else
            {
                var factor = DefaultLineFactor(para.DominantMetrics);
                pitch = dominantSize * factor;
            }
            pitch *= 1.0 - spacingReduction;

            var before = ParagraphSpacing(para.SpaceBefore, para.SpaceBeforeIsPct, dominantSize, fontScale);
            var after = ParagraphSpacing(para.SpaceAfter, para.SpaceAfterIsPct, dominantSize, fontScale);

            totalHeight += lineCount * pitch + before + after;
        }

        foreach (var face in missingGlyphFaces)
        {
            warnings.Add($"Font '{face}' lacks advance widths for some measured runes; 0.5em fallback used (unmeasurable glyphs).");
        }

        return (maxLineWidth, totalHeight);
    }

    private static double ParagraphSpacing(double? value, bool isPct, double dominantSize, double scale)
    {
        if (value is not { } v) return 0;
        // Percent spacing is of the paragraph's line height (≈ size × 1.2 with the
        // default factor); absolute points scale with the font scale (PowerPoint
        // shrinks the whole block proportionally). dominantSize is already scaled.
        return isPct ? dominantSize * 1.2 * v : v * scale;
    }

    private (int Lines, double Widest) WrapParagraph(
        MeasuredParagraph para, double firstLineWidth, double continuationWidth, double scale,
        HashSet<string> missingGlyphFaces)
    {
        var tokens = Tokenize(para, scale, missingGlyphFaces);
        if (tokens.Count == 0)
            return (1, 0); // empty paragraph still occupies one line in PowerPoint

        var lines = 1;
        var lineLimit = firstLineWidth;
        var lineUsed = 0.0;
        var pendingSpace = 0.0;
        var widest = 0.0;

        foreach (var token in tokens)
        {
            if (token.IsHardBreak)
            {
                widest = Math.Max(widest, lineUsed);
                lines++;
                lineUsed = 0;
                pendingSpace = 0;
                lineLimit = continuationWidth;
                continue;
            }
            if (token.IsBreak)
            {
                pendingSpace += token.Width;
                continue;
            }

            var needed = (lineUsed > 0 ? pendingSpace : 0) + token.Width;
            if (lineUsed > 0 && lineUsed + needed > lineLimit + 1e-6)
            {
                widest = Math.Max(widest, lineUsed);
                lines++;
                lineUsed = token.Width; // overlong words stay on their own line (horizontal overflow, PowerPoint-like)
                lineLimit = continuationWidth;
            }
            else
            {
                lineUsed += needed;
            }
            pendingSpace = 0;
        }

        widest = Math.Max(widest, lineUsed);
        return (lines, widest);
    }

    private List<Token> Tokenize(MeasuredParagraph para, double scale, HashSet<string> missingGlyphFaces)
    {
        var tokens = new List<Token>();
        foreach (var run in para.Runs)
        {
            var size = run.FontSize * scale;
            var word = new System.Text.StringBuilder();
            double wordWidth = 0;

            void FlushWord()
            {
                if (word.Length > 0)
                {
                    tokens.Add(new Token(word.ToString(), false, false, wordWidth));
                    word.Clear();
                    wordWidth = 0;
                }
            }

            foreach (var rune in run.Text.EnumerateRunes())
            {
                if (rune.Value == '\n')
                {
                    FlushWord();
                    tokens.Add(new Token("\n", false, true, 0));
                    continue;
                }
                if (rune.Value == ' ' || rune.Value == '\t')
                {
                    FlushWord();
                    tokens.Add(new Token(rune.ToString(), true, false, MeasureRune(rune, run, size, missingGlyphFaces)));
                    continue;
                }
                // CJK runes are individually breakable in PowerPoint line layout.
                if (IsCjk(rune.Value))
                {
                    FlushWord();
                    tokens.Add(new Token(rune.ToString(), false, false, MeasureRune(rune, run, size, missingGlyphFaces)));
                    continue;
                }

                word.Append(rune.ToString());
                wordWidth += MeasureRune(rune, run, size, missingGlyphFaces);
            }
            FlushWord();
        }
        return tokens;
    }

    private double MeasureString(string text, TypstFontMetrics? metrics, double sizePt, string family, HashSet<string> missingGlyphFaces)
    {
        var total = 0.0;
        foreach (var rune in text.EnumerateRunes())
        {
            total += AdvancePt(rune, metrics, sizePt, family, missingGlyphFaces);
        }
        return total;
    }

    private double MeasureRune(System.Text.Rune rune, MeasuredRun run, double sizePt, HashSet<string> missingGlyphFaces)
        => AdvancePt(rune, run.Metrics, sizePt, run.Family, missingGlyphFaces);

    private static double AdvancePt(System.Text.Rune rune, TypstFontMetrics? metrics, double sizePt, string family, HashSet<string> missingGlyphFaces)
    {
        if (metrics is { UnitsPerEm: > 0 } && metrics.AdvanceWidths.Count > 0)
        {
            if (metrics.AdvanceWidths.TryGetValue(rune.Value, out var advance))
                return advance * sizePt / metrics.UnitsPerEm;
            missingGlyphFaces.Add(family);
        }
        return 0.5 * sizePt; // unmeasurable fallback — never silently zero
    }

    private static double DefaultLineFactor(TypstFontMetrics? metrics)
    {
        if (metrics is { UnitsPerEm: > 0 })
        {
            var factor = (metrics.TypoAscender - metrics.TypoDescender + metrics.TypoLineGap) / (double)metrics.UnitsPerEm;
            if (factor > 0.5 && factor < 3.0)
                return factor;
        }
        return 1.2;
    }

    // -------------------------------------------------------- cascades

    private sealed record RawFormatting(double FontSize, bool Bold, bool Italic, string Family)
    {
        public static readonly RawFormatting Default = new(18.0, false, false, "Arial");
    }

    private static RawFormatting ExtractRunFormatting(Drawing.RunProperties? rPr)
    {
        if (rPr == null)
            return RawFormatting.Default;

        var fmt = RawFormatting.Default;
        if (rPr.FontSize?.Value != null)
            fmt = fmt with { FontSize = rPr.FontSize.Value / 100.0 };
        if (rPr.Bold?.Value != null)
            fmt = fmt with { Bold = rPr.Bold.Value };
        if (rPr.Italic?.Value != null)
            fmt = fmt with { Italic = rPr.Italic.Value };

        var latin = rPr.Elements<Drawing.LatinFont>().FirstOrDefault()?.Typeface?.Value;
        if (!string.IsNullOrEmpty(latin))
        {
            if (latin.EndsWith(" Bold", StringComparison.OrdinalIgnoreCase))
            {
                latin = latin[..^5];
                fmt = fmt with { Bold = true };
            }
            fmt = fmt with { Family = latin };
        }
        return fmt;
    }

    private RawFormatting ResolveParagraphFormatting(
        RawFormatting? firstRunFmt,
        Drawing.ParagraphProperties? pPr,
        OpenXmlElement? bodyLstStyle,
        int level,
        StyleResolver? styleResolver,
        int? placeholderIdx,
        PlaceholderValues? placeholderType)
    {
        var fmt = firstRunFmt ?? RawFormatting.Default;

        // 2. Paragraph default: pPr/defRPr (fill sentinel slots only, converter pattern)
        fmt = MergeDefRPr(fmt, pPr?.Elements<Drawing.DefaultRunProperties>().FirstOrDefault());

        // 3. Text body list style (shape-level lstStyle)
        fmt = MergeDefRPr(fmt, FindLevelDefRPr(bodyLstStyle, level));

        // 4./5. Layout then master placeholder list styles (StyleResolver public surface)
        if (styleResolver != null)
        {
            fmt = MergeDefaultTextStyle(fmt, styleResolver.GetLayoutPlaceholderLstStyle(placeholderIdx, placeholderType, level));
            fmt = MergeDefaultTextStyle(fmt, styleResolver.GetMasterPlaceholderLstStyle(placeholderIdx, placeholderType, level));
        }
        return fmt;
    }

    private static RawFormatting MergeRunFormatting(RawFormatting paragraphDefault, RawFormatting runFmt)
    {
        // Run properties override paragraph defaults when present (converter's
        // MergeRunWithParagraphDefaults): sentinel values mean "inherit".
        // (Bold/Italic sentinels conflate "absent" with explicit off — a shared
        // converter approximation, acceptable for width measurement.)
        var fmt = paragraphDefault;
        if (runFmt.FontSize != 18.0) fmt = fmt with { FontSize = runFmt.FontSize };
        if (runFmt.Bold) fmt = fmt with { Bold = true };
        if (runFmt.Italic) fmt = fmt with { Italic = true };
        if (runFmt.Family != "Arial") fmt = fmt with { Family = runFmt.Family };
        return fmt;
    }

    private static RawFormatting MergeDefRPr(RawFormatting fmt, OpenXmlElement? defRPr)
    {
        if (defRPr == null)
            return fmt;

        if (fmt.FontSize == 18.0)
        {
            var sz = GetAttributeValue(defRPr, "sz");
            if (sz != null && int.TryParse(sz, out var szHundredths))
                fmt = fmt with { FontSize = szHundredths / 100.0 };
        }
        if (!fmt.Bold)
        {
            var b = GetAttributeValue(defRPr, "b");
            if (b is "1" or "true")
                fmt = fmt with { Bold = true };
        }
        if (!fmt.Italic)
        {
            var i = GetAttributeValue(defRPr, "i");
            if (i is "1" or "true")
                fmt = fmt with { Italic = true };
        }
        if (fmt.Family == "Arial")
        {
            var latin = defRPr.Elements<Drawing.LatinFont>().FirstOrDefault()?.Typeface?.Value;
            if (!string.IsNullOrEmpty(latin))
            {
                if (latin.EndsWith(" Bold", StringComparison.OrdinalIgnoreCase))
                {
                    latin = latin[..^5];
                    fmt = fmt with { Bold = true };
                }
                fmt = fmt with { Family = latin };
            }
        }
        return fmt;
    }

    private static RawFormatting MergeDefaultTextStyle(RawFormatting fmt, StyleResolver.DefaultTextStyle? style)
    {
        if (style == null)
            return fmt;
        if (fmt.FontSize == 18.0 && style.FontSize.HasValue)
            fmt = fmt with { FontSize = style.FontSize.Value };
        if (!fmt.Bold && style.Bold.HasValue)
            fmt = fmt with { Bold = style.Bold.Value };
        if (!fmt.Italic && style.Italic.HasValue)
            fmt = fmt with { Italic = style.Italic.Value };
        if (fmt.Family == "Arial" && !string.IsNullOrEmpty(style.FontFamily))
        {
            var family = style.FontFamily;
            if (family.EndsWith(" Bold", StringComparison.OrdinalIgnoreCase))
            {
                family = family[..^5];
                fmt = fmt with { Bold = true };
            }
            fmt = fmt with { Family = family };
        }
        return fmt;
    }

    private static OpenXmlElement? FindLevelDefRPr(OpenXmlElement? lstStyle, int level)
    {
        var lvlPpr = lstStyle?.ChildElements.FirstOrDefault(e => e.LocalName == $"lvl{level + 1}pPr");
        return lvlPpr?.ChildElements.FirstOrDefault(e => e.LocalName == "defRPr");
    }

    private static (double MarLPt, double IndentPt) ExtractMarginAndIndent(Drawing.ParagraphProperties? pPr)
    {
        double marL = 0, indent = 0;
        if (pPr != null)
        {
            var marLAttr = GetAttributeValue(pPr, "marL");
            if (marLAttr != null && int.TryParse(marLAttr, out var m))
                marL = EmuToPt(m);
            var indentAttr = GetAttributeValue(pPr, "indent");
            if (indentAttr != null && int.TryParse(indentAttr, out var i))
                indent = EmuToPt(i);
        }
        return (marL, indent);
    }

    private static (string? Char, string? AutoNumber, bool HasBullet) ResolveBullet(
        Drawing.ParagraphProperties? pPr,
        OpenXmlElement? bodyLstStyle,
        int level,
        StyleResolver? styleResolver,
        int? placeholderIdx,
        PlaceholderValues? placeholderType)
    {
        var info = ExtractBulletInfo(pPr);
        if (info.Decided) return (info.Char, info.AutoNumber, info.HasBullet);

        info = ExtractBulletInfo(FindLevelPPr(bodyLstStyle, level));
        if (info.Decided) return (info.Char, info.AutoNumber, info.HasBullet);

        if (styleResolver != null)
        {
            var layout = styleResolver.GetLayoutPlaceholderBulletInfo(placeholderIdx, placeholderType, level);
            if (layout.HasBullet || layout.HasBulletNone) return (layout.BulletChar, layout.AutoNumberType, layout.HasBullet);
            var master = styleResolver.GetMasterPlaceholderBulletInfo(placeholderIdx, placeholderType, level);
            if (master.HasBullet || master.HasBulletNone) return (master.BulletChar, master.AutoNumberType, master.HasBullet);
            var tx = styleResolver.GetMasterTxStyleBulletInfo(placeholderType, level);
            if (tx.HasBullet || tx.HasBulletNone) return (tx.BulletChar, tx.AutoNumberType, tx.HasBullet);
        }
        return (null, null, false);
    }

    private static (string? Char, string? AutoNumber, bool HasBullet, bool Decided) ExtractBulletInfo(OpenXmlElement? pPr)
    {
        if (pPr == null) return (null, null, false, false);

        var buChar = pPr.ChildElements.FirstOrDefault(e => e.LocalName == "buChar");
        if (buChar != null)
            return (GetAttributeValue(buChar, "char"), null, true, true);

        var buAutoNum = pPr.ChildElements.FirstOrDefault(e => e.LocalName == "buAutoNum");
        if (buAutoNum != null)
            return (null, GetAttributeValue(buAutoNum, "type"), true, true);

        var buNone = pPr.ChildElements.FirstOrDefault(e => e.LocalName == "buNone");
        if (buNone != null)
            return (null, null, false, true);

        return (null, null, false, false);
    }

    private static (double? Value, bool IsPct) ResolveLineSpacing(
        Drawing.ParagraphProperties? pPr,
        OpenXmlElement? bodyLstStyle,
        int level,
        StyleResolver? styleResolver,
        int? placeholderIdx,
        PlaceholderValues? placeholderType)
    {
        var fromPPr = ExtractLineSpacing(pPr);
        if (fromPPr.Value != null) return fromPPr;

        var fromLst = ExtractLineSpacing(FindLevelPPr(bodyLstStyle, level));
        if (fromLst.Value != null) return fromLst;

        if (styleResolver != null)
        {
            // StyleResolver returns merged doubles; disambiguate pct vs points with
            // the converter's heuristic (values < 10 are percentage multipliers).
            var v = styleResolver.GetLayoutPlaceholderLineSpacing(placeholderIdx, placeholderType, level)
                ?? styleResolver.GetMasterPlaceholderLineSpacing(placeholderIdx, placeholderType, level)
                ?? styleResolver.GetMasterTxStyleLineSpacing(placeholderType, level);
            if (v.HasValue)
                return (v.Value, v.Value < 10);
        }
        return (null, false);
    }

    private static (double? Value, bool IsPct) ExtractLineSpacing(OpenXmlElement? pPr)
    {
        var lnSpc = pPr?.ChildElements.FirstOrDefault(e => e.LocalName == "lnSpc");
        if (lnSpc == null) return (null, false);

        var spcPts = lnSpc.ChildElements.FirstOrDefault(e => e.LocalName == "spcPts");
        var ptsVal = GetAttributeValue(spcPts, "val");
        if (ptsVal != null && int.TryParse(ptsVal, out var pts))
            return (pts / 100.0, false);

        var spcPct = lnSpc.ChildElements.FirstOrDefault(e => e.LocalName == "spcPct");
        var pctVal = GetAttributeValue(spcPct, "val");
        if (pctVal != null && int.TryParse(pctVal, out var pct))
            return (pct / 100000.0, true); // spcPct: 1/1000ths of a percent (rule 2)

        return (null, false);
    }

    private static (double? Value, bool IsPct) ResolveSpacing(
        string which,
        Drawing.ParagraphProperties? pPr,
        OpenXmlElement? bodyLstStyle,
        int level,
        StyleResolver? styleResolver,
        int? placeholderIdx,
        PlaceholderValues? placeholderType)
    {
        var fromPPr = ExtractSpacing(pPr, which);
        if (fromPPr.Value != null) return fromPPr;

        var fromLst = ExtractSpacing(FindLevelPPr(bodyLstStyle, level), which);
        if (fromLst.Value != null) return fromLst;

        if (styleResolver != null)
        {
            var (bef, aft) = styleResolver.GetLayoutPlaceholderSpacing(placeholderIdx, placeholderType, level);
            if (bef == null && aft == null)
                (bef, aft) = styleResolver.GetMasterPlaceholderSpacing(placeholderIdx, placeholderType, level);
            if (bef == null && aft == null)
                (bef, aft) = styleResolver.GetMasterTxStyleSpacing(placeholderType, level);

            var v = which == "spcBef" ? bef : aft;
            if (v.HasValue)
                return (v.Value, v.Value < 10); // converter pct/points heuristic
        }
        return (null, false);
    }

    private static (double? Value, bool IsPct) ExtractSpacing(OpenXmlElement? pPr, string which)
    {
        var el = pPr?.ChildElements.FirstOrDefault(e => e.LocalName == which);
        if (el == null) return (null, false);

        var spcPts = el.ChildElements.FirstOrDefault(e => e.LocalName == "spcPts");
        var ptsVal = GetAttributeValue(spcPts, "val");
        if (ptsVal != null && int.TryParse(ptsVal, out var pts))
            return (pts / 100.0, false); // val in hundredths of a point

        var spcPct = el.ChildElements.FirstOrDefault(e => e.LocalName == "spcPct");
        var pctVal = GetAttributeValue(spcPct, "val");
        if (pctVal != null && int.TryParse(pctVal, out var pct))
            return (pct / 100000.0, true);

        return (null, false);
    }

    private static OpenXmlElement? FindLevelPPr(OpenXmlElement? lstStyle, int level)
        => lstStyle?.ChildElements.FirstOrDefault(e => e.LocalName == $"lvl{level + 1}pPr");

    private static Drawing.BodyProperties? CascadeBodyPr(
        OpenXmlElement textBody,
        StyleResolver? styleResolver,
        int? placeholderIdx,
        PlaceholderValues? placeholderType)
    {
        var slideBodyPr = textBody.Elements<Drawing.BodyProperties>().FirstOrDefault();
        var layoutBodyPr = styleResolver?.GetLayoutPlaceholderBodyPr(placeholderIdx, placeholderType);
        var masterBodyPr = styleResolver?.GetMasterPlaceholderBodyPr(placeholderIdx, placeholderType);

        if (slideBodyPr == null && layoutBodyPr == null && masterBodyPr == null)
            return slideBodyPr;

        var result = new Drawing.BodyProperties();
        foreach (var source in new[] { masterBodyPr, layoutBodyPr, slideBodyPr })
        {
            if (source == null) continue;
            if (source.LeftInset?.Value != null) result.LeftInset = new Int32Value(source.LeftInset.Value);
            if (source.TopInset?.Value != null) result.TopInset = new Int32Value(source.TopInset.Value);
            if (source.RightInset?.Value != null) result.RightInset = new Int32Value(source.RightInset.Value);
            if (source.BottomInset?.Value != null) result.BottomInset = new Int32Value(source.BottomInset.Value);
        }

        // Autofit child: slide wins, then layout, then master (converter order).
        var autofit = slideBodyPr?.ChildElements.FirstOrDefault(e => e is Drawing.NoAutoFit or Drawing.NormalAutoFit or Drawing.ShapeAutoFit)
            ?? layoutBodyPr?.ChildElements.FirstOrDefault(e => e is Drawing.NoAutoFit or Drawing.NormalAutoFit or Drawing.ShapeAutoFit)
            ?? masterBodyPr?.ChildElements.FirstOrDefault(e => e is Drawing.NoAutoFit or Drawing.NormalAutoFit or Drawing.ShapeAutoFit);
        if (autofit != null)
            result.Append(autofit.CloneNode(true));

        return result;
    }

    // -------------------------------------------------------- helpers

    private static (int? Idx, PlaceholderValues? Type) GetPlaceholderInfo(P.Shape shape)
    {
        var nvSpPr = shape.NonVisualShapeProperties;
        var ph = nvSpPr?.Elements<PlaceholderShape>().FirstOrDefault()
            ?? nvSpPr?.ApplicationNonVisualDrawingProperties?.Elements<PlaceholderShape>().FirstOrDefault();
        if (ph == null)
            return (null, null);

        int? idx = null;
        var idxAttr = GetAttributeValue(ph, "idx");
        if (idxAttr != null && int.TryParse(idxAttr, out var parsed))
            idx = parsed;

        var typeAttr = GetAttributeValue(ph, "type");
        PlaceholderValues? type = typeAttr switch
        {
            "title" => PlaceholderValues.Title,
            "ctrTitle" => PlaceholderValues.CenteredTitle,
            "subTitle" => PlaceholderValues.SubTitle,
            "body" => PlaceholderValues.Body,
            "pic" => PlaceholderValues.Picture,
            "chart" => PlaceholderValues.Chart,
            "tbl" => PlaceholderValues.Table,
            "sldNum" => PlaceholderValues.SlideNumber,
            "ftr" => PlaceholderValues.Footer,
            "hdr" => PlaceholderValues.Header,
            "obj" => PlaceholderValues.Object,
            "dt" => PlaceholderValues.DateAndTime,
            _ => null
        };
        return (idx, type);
    }

    private static List<P.Shape> FindShapesById(ShapeTree shapeTree, uint elementId)
    {
        var matches = new List<P.Shape>();
        foreach (var element in shapeTree.ChildElements)
        {
            if (element is not P.Shape candidate)
                continue;
            var nv = (OpenXmlElement?)candidate.NonVisualShapeProperties;
            var cnvPr = nv?.ChildElements.FirstOrDefault(e => e.LocalName == "cNvPr");
            if (cnvPr == null) continue;
            var match = IdAttributePattern.Match(cnvPr.OuterXml);
            if (match.Success && uint.TryParse(match.Groups[1].Value, out var id) && id == elementId)
                matches.Add(candidate);
        }
        return matches;
    }

    private static bool ContainsComplexScript(string text)
    {
        foreach (var rune in text.EnumerateRunes())
        {
            var v = rune.Value;
            if (IsCjk(v) ||
                v is >= 0x0590 and <= 0x05FF ||   // Hebrew
                v is >= 0x0600 and <= 0x06FF ||   // Arabic
                v is >= 0x0750 and <= 0x077F ||   // Arabic supplement
                v is >= 0x08A0 and <= 0x08FF ||   // Arabic extended
                v is >= 0x0900 and <= 0x097F ||   // Devanagari
                v is >= 0x0E00 and <= 0x0E7F)     // Thai
                return true;
        }
        return false;
    }

    private static bool IsCjk(int v)
        => v is >= 0x2E80 and <= 0x9FFF || v is >= 0xAC00 and <= 0xD7AF || v is >= 0xF900 and <= 0xFAFF || v >= 0x20000;

    private static string? GetAttributeValue(OpenXmlElement? element, string attributeName)
    {
        if (element == null) return null;
        var match = Regex.Match(element.OuterXml, $@"\b{Regex.Escape(attributeName)}\s*=\s*""([^""]*)""");
        return match.Success ? match.Groups[1].Value : null;
    }

    private static double EmuToPt(long emu) => emu / EmuPerPoint;

    private static IReadOnlyList<string> Deduplicate(List<string> warnings)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var result = new List<string>();
        foreach (var w in warnings)
        {
            if (seen.Add(w))
                result.Add(w);
        }
        return result;
    }
}
