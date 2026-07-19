using PptxEditor.Core.Models;
using PptxEditor.Core.Services;

namespace PptxEditor.Core.Generation.Layout;

/// <summary>
/// Real font-metric <see cref="ITextMeasurer"/> (P3, plan.md §3.2 overflow pass). Measures
/// resolved runs against their box on top of TextFitService's font metrics — the shared
/// <see cref="FontMetricsCatalog"/> — rather than a second metrics stack.
/// <para>
/// Tokenization (space/tab breaks, per-rune CJK breaks, '\n' hard breaks), greedy wrapping,
/// the missing-glyph 0.5em fallback and the default line factor
/// (TypoAscender − TypoDescender + TypoLineGap, guarded to 1.2) deliberately mirror
/// <see cref="TextFitService"/> so a generation-side fit verdict matches the F6 edit-side
/// verdict for the same content. OOXML paragraph concepts that do not exist in the
/// generation vocabulary (marL/indent, bullets, spcBef/spcAft, lnSpcReduction) are not
/// emulated — the generation shrink contract is a pure fontScale (plan.md §3.2).
/// </para>
/// <para>
/// The shrink search steps 99% → <see cref="TextMeasureRequest.MinScale"/> at 1% steps
/// (TextFitService's PowerPoint-order search). When nothing fits at MinScale the search
/// continues below it, so the resolver applies the true scale and raises its MinScale
/// warning with real numbers (plan.md §7.5 — oversized text shrinks or errors loudly,
/// never silently clips); the return value bottoms out at 0.01, never 0.
/// </para>
/// </summary>
public sealed class TextMeasure : ITextMeasurer
{
    private const double ScaleStep = 0.01;
    private const double Eps = 1e-6;
    private const string DefaultFamily = "Arial";

    private readonly FontMetricsCatalog _catalog;
    private readonly List<string> _warnings = [];

    /// <summary>Creates a measurer over the given (possibly cached/shared) metrics catalog.</summary>
    public TextMeasure(FontMetricsCatalog catalog)
    {
        ArgumentNullException.ThrowIfNull(catalog);
        _catalog = catalog;
    }

    /// <summary>Creates a measurer over a default catalog (system font scan, no embedded fonts).</summary>
    public TextMeasure()
        : this(new FontMetricsCatalog())
    {
    }

    /// <summary>
    /// Non-fatal accuracy notes from the most recent <see cref="FitScale"/> call (font
    /// substitution, missing metrics, complex-script shaping limits). Never silently empty
    /// when measurement degraded.
    /// </summary>
    public IReadOnlyList<string> Warnings => _warnings;

    /// <inheritdoc/>
    public double FitScale(TextMeasureRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        ArgumentNullException.ThrowIfNull(request.Runs);
        if (request.BoxWidthPt < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(request), "BoxWidthPt must be ≥ 0.");
        }
        if (request.BoxHeightPt < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(request), "BoxHeightPt must be ≥ 0.");
        }

        _warnings.Clear();
        var text = Tokenize(request.Runs);
        DeduplicateWarnings();

        if (Fits(text, request.BoxWidthPt, request.BoxHeightPt, 1.0))
        {
            return 1.0;
        }

        var minPct = Math.Max((int)Math.Round(Math.Clamp(request.MinScale, ScaleStep, 1.0) * 100), 1);
        for (var pct = 99; pct >= minPct; pct--)
        {
            if (Fits(text, request.BoxWidthPt, request.BoxHeightPt, pct / 100.0))
            {
                return pct / 100.0;
            }
        }

        // MinScale violated: keep searching below it so the resolver applies the true scale
        // and warns loudly (plan.md §7.5) instead of silently clipping at a bottomed value.
        for (var pct = minPct - 1; pct >= 1; pct--)
        {
            if (Fits(text, request.BoxWidthPt, request.BoxHeightPt, pct / 100.0))
            {
                return pct / 100.0;
            }
        }

        return ScaleStep; // floor: content cannot fit at any legible scale — never 0
    }

    // -------------------------------------------------------- measurement
    // Mirrors TextFitService's measurement model: same token rules, same advance-width
    // accumulation, same line-pitch factor, same fallbacks.

    private sealed record MeasuredRun(string Text, string Family, double FontSizePt, TypstFontMetrics? Metrics);

    private readonly record struct Token(bool IsBreak, bool IsHardBreak, double WidthPt);

    private sealed record MeasuredText(IReadOnlyList<Token> Tokens, double PitchPt);

    private MeasuredText Tokenize(IReadOnlyList<ResolvedTextRun> runs)
    {
        var measured = new List<MeasuredRun>(runs.Count);
        var hasComplex = false;
        foreach (var run in runs)
        {
            // Null family = emitter default; measure it as the F6 default face.
            var resolution = _catalog.Resolve(run.FontFamily ?? DefaultFamily, run.Bold, run.Italic);
            foreach (var warning in resolution.Warnings)
            {
                _warnings.Add(warning);
            }
            if (ContainsComplexScript(run.Text))
            {
                hasComplex = true;
            }
            // Empty-text runs are kept: they carry the size that sets an empty line's pitch
            // (PowerPoint endParaRPr behavior, mirrored from TextFitService).
            measured.Add(new MeasuredRun(run.Text, resolution.ResolvedFamily, run.FontSizePt, resolution.Metrics));
        }
        if (hasComplex)
        {
            _warnings.Add("Text contains complex-script runes (CJK/Arabic/Hebrew/Indic/Thai); measurement uses raw advance widths without a shaping/kerning engine and may be materially wrong.");
        }

        double pitch;
        if (measured.Count == 0)
        {
            pitch = 18.0 * 1.2; // TextFitService's empty-paragraph default
        }
        else
        {
            var dominant = measured.OrderByDescending(r => r.FontSizePt).First();
            pitch = dominant.FontSizePt * DefaultLineFactor(dominant.Metrics);
        }

        var missingGlyphFaces = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var tokens = new List<Token>();
        foreach (var run in measured)
        {
            var wordWidth = 0.0;
            var hasWord = false;

            void FlushWord()
            {
                if (hasWord)
                {
                    tokens.Add(new Token(false, false, wordWidth));
                    wordWidth = 0;
                    hasWord = false;
                }
            }

            foreach (var rune in run.Text.EnumerateRunes())
            {
                if (rune.Value == '\n')
                {
                    FlushWord();
                    tokens.Add(new Token(false, true, 0));
                    continue;
                }
                if (rune.Value is ' ' or '\t')
                {
                    FlushWord();
                    tokens.Add(new Token(true, false, AdvancePt(rune, run, missingGlyphFaces)));
                    continue;
                }
                // CJK runes are individually breakable in PowerPoint line layout.
                if (IsCjk(rune.Value))
                {
                    FlushWord();
                    tokens.Add(new Token(false, false, AdvancePt(rune, run, missingGlyphFaces)));
                    continue;
                }

                hasWord = true;
                wordWidth += AdvancePt(rune, run, missingGlyphFaces);
            }
            FlushWord();
        }

        foreach (var face in missingGlyphFaces)
        {
            _warnings.Add($"Font '{face}' lacks advance widths for some measured runes; 0.5em fallback used (unmeasurable glyphs).");
        }

        return new MeasuredText(tokens, pitch);
    }

    private static bool Fits(MeasuredText text, double widthPt, double heightPt, double scale)
    {
        var (lines, widest) = Wrap(text.Tokens, Math.Max(widthPt, 1.0), scale);
        var height = lines * text.PitchPt * scale;
        return height <= heightPt + Eps && widest <= widthPt + Eps;
    }

    /// <summary>Greedy wrap at <paramref name="scale"/> (widths scale linearly). Returns line count and widest line in points.</summary>
    private static (int Lines, double WidestPt) Wrap(IReadOnlyList<Token> tokens, double widthPt, double scale)
    {
        if (tokens.Count == 0)
        {
            return (1, 0); // empty content still occupies one line in PowerPoint
        }

        var lines = 1;
        double used = 0, pendingSpace = 0, widest = 0;
        foreach (var token in tokens)
        {
            if (token.IsHardBreak)
            {
                widest = Math.Max(widest, used);
                lines++;
                used = 0;
                pendingSpace = 0;
                continue;
            }
            var width = token.WidthPt * scale;
            if (token.IsBreak)
            {
                pendingSpace += width;
                continue;
            }

            var needed = (used > 0 ? pendingSpace : 0) + width;
            if (used > 0 && used + needed > widthPt + Eps)
            {
                widest = Math.Max(widest, used);
                lines++;
                used = width; // overlong words stay on their own line (horizontal overflow, PowerPoint-like)
            }
            else
            {
                used += needed;
            }
            pendingSpace = 0;
        }

        widest = Math.Max(widest, used);
        return (lines, widest);
    }

    private static double AdvancePt(System.Text.Rune rune, MeasuredRun run, HashSet<string> missingGlyphFaces)
    {
        var metrics = run.Metrics;
        if (metrics is { UnitsPerEm: > 0 } && metrics.AdvanceWidths.Count > 0)
        {
            if (metrics.AdvanceWidths.TryGetValue(rune.Value, out var advance))
            {
                return advance * run.FontSizePt / metrics.UnitsPerEm;
            }
            missingGlyphFaces.Add(run.Family);
        }
        return 0.5 * run.FontSizePt; // unmeasurable fallback — never silently zero
    }

    private static double DefaultLineFactor(TypstFontMetrics? metrics)
    {
        if (metrics is { UnitsPerEm: > 0 })
        {
            var factor = (metrics.TypoAscender - metrics.TypoDescender + metrics.TypoLineGap) / (double)metrics.UnitsPerEm;
            if (factor > 0.5 && factor < 3.0)
            {
                return factor;
            }
        }
        return 1.2;
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
            {
                return true;
            }
        }
        return false;
    }

    private static bool IsCjk(int v)
        => v is >= 0x2E80 and <= 0x9FFF || v is >= 0xAC00 and <= 0xD7AF || v is >= 0xF900 and <= 0xFAFF || v >= 0x20000;

    private void DeduplicateWarnings()
    {
        if (_warnings.Count < 2)
        {
            return;
        }
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var deduped = _warnings.Where(seen.Add).ToList();
        _warnings.Clear();
        _warnings.AddRange(deduped);
    }
}
