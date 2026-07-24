using System.Globalization;
using System.Text;

namespace PptxBenchmark;

/// <summary>Renders the markdown benchmark report.</summary>
public static class ReportWriter
{
    public static string Build(
        EnvironmentInfo environment,
        IReadOnlyList<DeckBenchmark> decks,
        LibreOfficeReport libreOffice,
        int warmRuns)
    {
        var sb = new StringBuilder();
        sb.AppendLine("# PPTX Render Benchmark — Baseline (Pre-Optimization)");
        sb.AppendLine();
        sb.AppendLine(CultureInfo.InvariantCulture, $"Generated: {DateTimeOffset.UtcNow:yyyy-MM-dd HH:mm:ss} UTC by `tools/pptx-benchmark`");
        sb.AppendLine();
        sb.AppendLine("Ground-truth latency capture **before** the converter optimization work (W3). ");
        sb.AppendLine("Product target for the instant-preview pipeline: **<500 ms per slide** (whole-deck PNG @150ppi path).");
        sb.AppendLine();

        AppendEnvironment(sb, environment, libreOffice);
        foreach (var deck in decks)
        {
            AppendDeck(sb, deck, libreOffice, warmRuns);
        }
        AppendSummary(sb, decks, libreOffice, warmRuns);
        AppendMethodology(sb, warmRuns);
        AppendLimitations(sb);

        return sb.ToString();
    }

    private static void AppendEnvironment(StringBuilder sb, EnvironmentInfo env, LibreOfficeReport libreOffice)
    {
        sb.AppendLine("## Environment");
        sb.AppendLine();
        sb.AppendLine("| Property | Value |");
        sb.AppendLine("|---|---|");
        sb.AppendLine($"| OS | {env.Os} |");
        sb.AppendLine($"| Architecture | {env.Architecture} |");
        sb.AppendLine($"| .NET runtime | {env.Runtime} |");
        sb.AppendLine($"| .NET SDK | {env.Sdk} |");
        sb.AppendLine($"| TypstBridge | {env.TypstBridgeVersion} |");
        sb.AppendLine($"| LibreOffice | {(libreOffice.Available ? libreOffice.Version ?? libreOffice.SofficePath ?? "present" : "not available on this machine")} |");
        sb.AppendLine($"| pdftoppm (poppler) | {(libreOffice.PdfToPpmAvailable ? "present" : "not available on this machine")} |");
        sb.AppendLine();
    }

    private static void AppendDeck(StringBuilder sb, DeckBenchmark deck, LibreOfficeReport libreOffice, int warmRuns)
    {
        sb.AppendLine($"## {deck.DeckName} ({deck.SlideCount} slides)");
        sb.AppendLine();
        sb.AppendLine("### This pipeline (PptxEditor → Typst)");
        sb.AppendLine();
        sb.AppendLine($"| Stage | Cold (ms) | Warm median (ms, N={warmRuns}) | Warm per-slide (ms) |");
        sb.AppendLine("|---|---|---|---|");
        AppendStageRow(sb, deck, "open", "PresentationBuilder.Open");
        AppendStageRow(sb, deck, "png", "ExportThumbnails (whole-deck PNG @150ppi)");
        AppendCompileRow(sb, deck, "png", warmRuns);
        AppendStageRow(sb, deck, "pdf", "ExportToPdf");
        AppendCompileRow(sb, deck, "pdf", warmRuns);

        var coldTotal = deck.ColdMs("open") + deck.ColdMs("png");
        var warmTotal = deck.WarmMedianMs("open") + deck.WarmMedianMs("png");
        AppendTotalRow(sb, "Total preview path (Open + PNG)", coldTotal, warmTotal, deck.SlideCount);
        sb.AppendLine();
        sb.AppendLine("`*.compile` rows are the Typst compile (TypstBridge backend) inside the stage above,");
        sb.AppendLine("captured via the `OFFICEEDITOR_TIMING` hooks in `TypstCompilerService`.");
        sb.AppendLine();

        sb.AppendLine("### LibreOffice (soffice → PDF → pdftoppm PNG @150dpi)");
        sb.AppendLine();
        var loDeck = libreOffice.Decks.FirstOrDefault(d => d.DeckName == deck.DeckName);
        if (!libreOffice.Available)
        {
            sb.AppendLine($"> {libreOffice.SkipReason}");
        }
        else if (loDeck is not null)
        {
            sb.AppendLine($"| Stage | Cold (ms) | Warm median (ms, N={warmRuns}) | Warm per-slide (ms, derived) |");
            sb.AppendLine("|---|---|---|---|");
            sb.AppendLine(CultureInfo.InvariantCulture,
                $"| soffice convert-to pdf | {loDeck.PdfColdMs:F1} | {loDeck.PdfWarmMedianMs:F1} | {loDeck.PdfWarmMedianMs / deck.SlideCount:F1} |");
            if (loDeck.RasterWarmMedianMs is { } rasterWarm)
            {
                sb.AppendLine(CultureInfo.InvariantCulture,
                    $"| pdftoppm rasterize (PNG @150dpi) | {F(loDeck.RasterColdMs!.Value)} | {rasterWarm:F1} | {rasterWarm / deck.SlideCount:F1} |");
                AppendTotalRow(sb, "Total (PDF + rasterize)", loDeck.TotalColdMs!.Value, loDeck.TotalWarmMedianMs!.Value, deck.SlideCount);
                if (loDeck.RasterPageCount is { } pages && pages != deck.SlideCount)
                {
                    sb.AppendLine();
                    sb.AppendLine(CultureInfo.InvariantCulture,
                        $"> WARNING: pdftoppm produced {pages} pages for a {deck.SlideCount}-slide deck — per-slide LO numbers are off.");
                }
            }
            else
            {
                sb.AppendLine($"> {libreOffice.RasterSkipReason}");
            }
        }
        sb.AppendLine();
    }

    private static void AppendStageRow(StringBuilder sb, DeckBenchmark deck, string stage, string label)
    {
        sb.AppendLine(CultureInfo.InvariantCulture,
            $"| {label} | {deck.ColdMs(stage):F1} | {deck.WarmMedianMs(stage):F1} | {deck.WarmMedianMs(stage) / deck.SlideCount:F1} |");
    }

    private static void AppendCompileRow(StringBuilder sb, DeckBenchmark deck, string stage, int warmRuns)
    {
        var cold = deck.ColdCompile(stage);
        var warm = deck.WarmMedianCompileMs(stage);
        var coldText = cold is null ? "n/a" : F(cold.Milliseconds);
        var warmText = warm is null ? "n/a" : F(warm.Value);
        var backend = cold?.Backend ?? "n/a";
        sb.AppendLine(CultureInfo.InvariantCulture,
            $"| ↳ {stage}.compile (backend: {backend}) | {coldText} | {warmText} | — |");
    }

    private static void AppendTotalRow(StringBuilder sb, string label, double coldMs, double warmMs, int slideCount)
    {
        sb.AppendLine(CultureInfo.InvariantCulture,
            $"| **{label}** | **{coldMs:F1}** | **{warmMs:F1}** | **{warmMs / slideCount:F1}** |");
    }

    private static void AppendSummary(StringBuilder sb, IReadOnlyList<DeckBenchmark> decks, LibreOfficeReport libreOffice, int warmRuns)
    {
        sb.AppendLine("## Cold vs warm summary (preview path: Open + whole-deck PNG @150ppi)");
        sb.AppendLine();
        sb.AppendLine($"| Deck | Slides | Cold total (ms) | Warm median total (ms, N={warmRuns}) | Derived per-slide (ms) | LibreOffice warm median (ms) |");
        sb.AppendLine("|---|---|---|---|---|---|");
        foreach (var deck in decks)
        {
            var coldTotal = deck.ColdMs("open") + deck.ColdMs("png");
            var warmTotal = deck.WarmMedianMs("open") + deck.WarmMedianMs("png");
            var loDeck = libreOffice.Decks.FirstOrDefault(d => d.DeckName == deck.DeckName);
            var loText = !libreOffice.Available || loDeck is null
                ? "n/a (not installed)"
                : loDeck.TotalWarmMedianMs is { } loTotal
                    ? F(loTotal)
                    : F(loDeck.PdfWarmMedianMs) + " (PDF only)";
            sb.AppendLine(CultureInfo.InvariantCulture,
                $"| {deck.DeckName} | {deck.SlideCount} | {coldTotal:F1} | {warmTotal:F1} | {warmTotal / deck.SlideCount:F1} | {loText} |");
        }
        sb.AppendLine();
    }

    private static void AppendMethodology(StringBuilder sb, int warmRuns)
    {
        sb.AppendLine("## Methodology");
        sb.AppendLine();
        sb.AppendLine($"- Each deck is measured in a **fresh child process**: iteration 0 is the cold run,");
        sb.AppendLine($"  iterations 1..{warmRuns} are the warm runs (reported as the median of N={warmRuns}).");
        sb.AppendLine("- Deck bytes are read from disk once, untimed; `PresentationBuilder.Open(byte[])` is the measured open stage.");
        sb.AppendLine("- Stages per iteration: `PresentationBuilder.Open` → `ExportThumbnails` (whole-deck PNG @150ppi)");
        sb.AppendLine("  → `ExportToPdf`. Wall-clock `Stopwatch` per stage.");
        sb.AppendLine("- Typst compile time inside each export is captured with the opt-in `OFFICEEDITOR_TIMING=1`");
        sb.AppendLine("  hooks in `TypstCompilerService` (backend, total compile ms per call).");
        sb.AppendLine("- Per-slide times are derived as stage total ÷ slide count (the current pipeline renders the whole deck).");
        sb.AppendLine("- LibreOffice leg: `soffice --headless --norestore --convert-to pdf --outdir <tmp> <deck>`,");
        sb.AppendLine("  timed from process start to exit, followed by `pdftoppm -png -r 150 <pdf> <tmp>/slide`, timed");
        sb.AppendLine("  separately in the same run. Cold run uses a fresh `-env:UserInstallation` profile");
        sb.AppendLine($"  directory; warm runs reuse one profile (median of N={warmRuns}). When `soffice` is not on PATH");
        sb.AppendLine("  or at `/Applications/LibreOffice.app/Contents/MacOS/soffice`, the leg is skipped without failing;");
        sb.AppendLine("  when `pdftoppm` is missing, the rasterization half is reported as unavailable without failing.");
        sb.AppendLine("- LibreOffice cannot rasterize PPTX natively; total = PDF conversion + pdftoppm rasterization at 150dpi.");
        sb.AppendLine("  OfficeEditor renders PNGs natively in a single compile.");
        sb.AppendLine("- No third-party benchmark dependencies: `Stopwatch` + `Process` only.");
        sb.AppendLine();
    }

    private static void AppendLimitations(StringBuilder sb)
    {
        sb.AppendLine("## Limitations");
        sb.AppendLine();
        sb.AppendLine("- Single machine, single pass — absolute numbers carry normal OS/JIT noise; treat as orders of magnitude.");
        sb.AppendLine("- Cold runs include one-time costs (assembly load/JIT, TypstBridge probe,");
        sb.AppendLine("  system-font discovery) exactly as a fresh API process would experience them.");
        sb.AppendLine("- `soffice --convert-to png` historically renders only slide 1 for Impress decks, so the LibreOffice");
        sb.AppendLine("  leg converts to PDF, rasterizes with pdftoppm, and per-slide LO numbers are derived from the");
        sb.AppendLine("  total (PDF + rasterize) ÷ slide count.");
        sb.AppendLine("- LibreOffice numbers (when present) include full process start + profile cost, not just rendering.");
        sb.AppendLine("- Converter-internal stages (font extraction, system-font discovery, per-slide conversion) are not");
        sb.AppendLine("  split out yet — `ExportThumbnails`/`ExportToPdf` are measured end to end, with the Typst compile");
        sb.AppendLine("  portion attributed via the timing hooks.");
        sb.AppendLine();
    }

    private static string F(double milliseconds) =>
        milliseconds.ToString("F1", CultureInfo.InvariantCulture);
}
