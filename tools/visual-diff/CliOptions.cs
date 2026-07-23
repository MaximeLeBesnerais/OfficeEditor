using System.Globalization;

namespace VisualDiff;

internal enum CliMode
{
    Help,
    Probe,
    Run,
    Check
}

/// <summary>
/// Parsed command line. Mode selection:
/// <list type="bullet">
/// <item><see cref="CliMode.Help"/> / <see cref="CliMode.Probe"/>: informational, always exit 0.</item>
/// <item><see cref="CliMode.Run"/>: compare a suite or an arbitrary pair; exits 0 unless
/// <c>--baseline</c> is also given, in which case the post-run threshold check decides.</item>
/// <item><see cref="CliMode.Check"/>: standalone threshold check of an existing metrics.json.</item>
/// </list>
/// </summary>
internal sealed record CliOptions(
    CliMode Mode,
    string? Suite,
    string? ReferencePath,
    string? GeneratedPath,
    string OutputDirectory,
    string? Name,
    int Dpi,
    bool GenerateMissing,
    string? FontPath,
    string? CheckMetricsPath,
    string? BaselinePath,
    double Margin,
    string? ThresholdsPath,
    bool RenderTypst)
{
    public const double DefaultMargin = 0.05;

    public static CliOptions Parse(string[] args)
    {
        string? suite = null;
        string? reference = null;
        string? generated = null;
        string? output = null;
        string? name = null;
        int dpi = 150;
        bool generate = false;
        string? fontPath = null;
        string? check = null;
        string? baseline = null;
        double margin = DefaultMargin;
        string? thresholds = null;
        bool render = false;
        bool probe = false;
        bool help = args.Length == 0;

        for (int i = 0; i < args.Length; i++)
        {
            string arg = args[i];
            switch (arg)
            {
                case "--help" or "-h":
                    help = true;
                    break;
                case "--probe":
                    probe = true;
                    break;
                case "--suite":
                    suite = RequireValue(args, ref i, arg);
                    break;
                case "--ref":
                    reference = RequireValue(args, ref i, arg);
                    break;
                case "--gen":
                    generated = RequireValue(args, ref i, arg);
                    break;
                case "--out":
                    output = RequireValue(args, ref i, arg);
                    break;
                case "--name":
                    name = RequireValue(args, ref i, arg);
                    break;
                case "--dpi":
                    string dpiText = RequireValue(args, ref i, arg);
                    if (!int.TryParse(dpiText, NumberStyles.None, CultureInfo.InvariantCulture, out dpi) || dpi <= 0)
                    {
                        throw new InvalidOperationException("--dpi must be a positive integer.");
                    }
                    break;
                case "--generate":
                    generate = true;
                    break;
                case "--font-path":
                    fontPath = RequireValue(args, ref i, arg);
                    break;
                case "--check":
                    check = RequireValue(args, ref i, arg);
                    break;
                case "--baseline":
                    baseline = RequireValue(args, ref i, arg);
                    break;
                case "--margin":
                    string marginText = RequireValue(args, ref i, arg);
                    if (!double.TryParse(marginText, NumberStyles.Float, CultureInfo.InvariantCulture, out margin) || margin < 0)
                    {
                        throw new InvalidOperationException("--margin must be a non-negative number (normalized RMSE units; 0.05 = 5 percentage points).");
                    }
                    break;
                case "--thresholds":
                    thresholds = RequireValue(args, ref i, arg);
                    break;
                case "--render":
                    render = true;
                    break;
                default:
                    throw new InvalidOperationException($"Unknown argument '{arg}'.");
            }
        }

        if (help)
        {
            return new CliOptions(CliMode.Help, suite, reference, generated,
                output ?? "examples/output/visual-diff/docx", name, dpi, generate, fontPath, check, baseline, margin, thresholds, render);
        }

        if (probe)
        {
            if (suite is not null || reference is not null || generated is not null || check is not null)
            {
                throw new InvalidOperationException("--probe cannot be combined with comparison options.");
            }

            return new CliOptions(CliMode.Probe, null, null, null,
                output ?? "examples/output/visual-diff/docx", null, dpi, false, null, null, null, margin, null, false);
        }

        if (baseline is not null && thresholds is not null)
        {
            throw new InvalidOperationException("--baseline and --thresholds are mutually exclusive: baseline is the machine-dependent REF-deck gate, thresholds the per-primitive parity gate.");
        }

        if (check is not null)
        {
            if (suite is not null || reference is not null || generated is not null)
            {
                throw new InvalidOperationException("--check cannot be combined with a comparison run; pass only --check plus --baseline or --thresholds.");
            }

            if (baseline is null && thresholds is null)
            {
                throw new InvalidOperationException("--check requires --baseline <metrics.json> or --thresholds <thresholds.json>.");
            }

            return new CliOptions(CliMode.Check, null, null, null,
                output ?? "examples/output/visual-diff/docx", null, dpi, false, null, check, baseline, margin, thresholds, false);
        }

        // Run mode validation.
        if (suite is not null && (reference is not null || generated is not null))
        {
            throw new InvalidOperationException("Provide either --suite or --ref/--gen, not both.");
        }

        if (suite is null && (reference is null || generated is null))
        {
            throw new InvalidOperationException("Provide either '--suite <docx|pptx|gen>' or both '--ref <path>' and '--gen <path>'.");
        }

        if (generate && !string.Equals(suite, "pptx", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(suite, "gen", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("--generate only applies to '--suite pptx' and '--suite gen'.");
        }

        if (fontPath is not null && !generate)
        {
            throw new InvalidOperationException("--font-path only makes sense together with --generate.");
        }

        if (render && !string.Equals(suite, "gen", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("--render only applies to '--suite gen' (it renders fixture .typ sources via the typst CLI).");
        }

        if (thresholds is not null && suite is not null && !string.Equals(suite, "gen", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("--thresholds is the per-primitive parity gate and only applies to '--suite gen'.");
        }

        output ??= suite switch
        {
            not null when string.Equals(suite, "docx", StringComparison.OrdinalIgnoreCase) => "examples/output/visual-diff/docx",
            not null when string.Equals(suite, "pptx", StringComparison.OrdinalIgnoreCase) => "examples/output/visual-diff/pptx",
            not null when string.Equals(suite, "gen", StringComparison.OrdinalIgnoreCase) => "examples/output/visual-diff/gen",
            not null => throw new InvalidOperationException($"Unknown suite '{suite}'. Supported suites: docx, pptx, gen."),
            _ => throw new InvalidOperationException("--out is required when comparing an arbitrary pair.")
        };

        return new CliOptions(CliMode.Run, suite, reference, generated, output, name, dpi, generate, fontPath, null, baseline, margin, thresholds, render);
    }

    private static string RequireValue(string[] args, ref int index, string option)
    {
        if (index + 1 >= args.Length || args[index + 1].StartsWith("--", StringComparison.Ordinal))
        {
            throw new InvalidOperationException($"{option} requires a value.");
        }

        index++;
        return args[index];
    }

    public static void PrintUsage()
    {
        Console.WriteLine("""
        visual-diff - render and compare PDF pages or PNG image pairs.

        Usage:
          visual-diff --suite <docx|pptx|gen> [--out <dir>] [--dpi 150] [--generate] [--font-path <dir>]
                      [--render] [--baseline <metrics.json>] [--margin 0.05] [--thresholds <thresholds.json>]
          visual-diff --ref <ref.pdf|ref.png|ref-dir> --gen <gen.pdf|gen.png|gen-dir>
                      --out <dir> [--name <name>] [--dpi 150] [--baseline <metrics.json>] [--margin 0.05]
          visual-diff --check <metrics.json> (--baseline <metrics.json> | --thresholds <thresholds.json>) [--margin 0.05]
          visual-diff --probe
          visual-diff --help

        Suites (run from the repository root):
          docx   examples/REF/DOCX/*.pdf  vs  examples/output/ref/docx/*.pdf
          pptx   examples/REF/PPTX/northwind-demo.pdf  vs  examples/output/ref/pptx/*.pdf
                 Missing generated PDFs can be built on demand with --generate (invokes
                 tools/convert-pptx) or manually:
                   dotnet run --project tools/convert-pptx -- examples/REF/PPTX/<deck>.pptx \
                     examples/output/ref/pptx/<deck>.pdf --format pdf
          gen    Phase 5 parity fixtures (PptxEditor.Core/Generation/Fixtures catalog):
                 PowerPoint ground-truth PNGs vs Typst preview PNGs per primitive deck under
                 examples/output/gen/parity/<fixture>/. --generate writes fixture.pptx /
                 fixture.typ / manifest.json / assets in-process; --render additionally renders
                 the Typst side via the typst CLI (probed, opt-in). Ground truth is produced
                 with PowerPoint manually - the suite prints the exact steps per fixture.

        Options:
          --suite <name>     Run a built-in comparison suite (docx, pptx, gen).
          --ref <path>       Reference PDF, single PNG, or directory of PNGs.
          --gen <path>       Generated PDF, single PNG, or directory of PNGs (same kind as --ref).
                             PNG inputs skip the poppler rendering step entirely.
          --out <path>       Output directory. Defaults to examples/output/visual-diff/<suite> for suites.
          --name <name>      Display/report name for an arbitrary pair.
          --dpi <number>     Rasterization DPI for PDF inputs and --render. Default: 150.
          --generate         (suites pptx, gen) Build missing inputs (pptx: convert-pptx PDFs;
                             gen: fixture decks + Typst sources, in-process).
          --font-path <dir>  Extra font directory passed to tools/convert-pptx with --generate.
          --render           (suite gen) Render fixture .typ sources to PNG pages via the typst CLI.
          --check <file>     Threshold-check an existing metrics.json against --baseline or --thresholds.
          --baseline <file>  Baseline metrics.json. After a run: check the fresh metrics against it.
          --thresholds <f>   Per-primitive thresholds.json (gen suite). After a run: check the fresh
                             metrics against the per-primitive RMSE ceilings. Mutually exclusive
                             with --baseline. Defaults for --suite gen to the committed
                             tools/visual-diff/baselines/gen/thresholds.json when present.
          --margin <number>  Absolute margin on normalized RMSE (0.05 = 5 percentage points).
                             Default: 0.05. (--baseline checks only.)
          --probe            Print external-tool availability (poppler, ImageMagick, typst) and exit.
          --help             Show this help text.

        Exit codes:
          0  success (report written; thresholds, if checked, satisfied)
          1  operational or usage error (missing tools, missing files, bad arguments)
          2  threshold check ran and at least one page/document exceeded baseline + margin
             or its per-primitive threshold

        Note: without --baseline/--thresholds/--check the tool is report-only and never fails on RMSE values.
        """);
    }
}
