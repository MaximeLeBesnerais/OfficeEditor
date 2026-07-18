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
    double Margin)
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
                default:
                    throw new InvalidOperationException($"Unknown argument '{arg}'.");
            }
        }

        if (help)
        {
            return new CliOptions(CliMode.Help, suite, reference, generated,
                output ?? "examples/output/visual-diff/docx", name, dpi, generate, fontPath, check, baseline, margin);
        }

        if (probe)
        {
            if (suite is not null || reference is not null || generated is not null || check is not null)
            {
                throw new InvalidOperationException("--probe cannot be combined with comparison options.");
            }

            return new CliOptions(CliMode.Probe, null, null, null,
                output ?? "examples/output/visual-diff/docx", null, dpi, false, null, null, null, margin);
        }

        if (check is not null)
        {
            if (suite is not null || reference is not null || generated is not null)
            {
                throw new InvalidOperationException("--check cannot be combined with a comparison run; pass only --check, --baseline and optionally --margin.");
            }

            if (baseline is null)
            {
                throw new InvalidOperationException("--check requires --baseline <metrics.json>.");
            }

            return new CliOptions(CliMode.Check, null, null, null,
                output ?? "examples/output/visual-diff/docx", null, dpi, false, null, check, baseline, margin);
        }

        // Run mode validation.
        if (suite is not null && (reference is not null || generated is not null))
        {
            throw new InvalidOperationException("Provide either --suite or --ref/--gen, not both.");
        }

        if (suite is null && (reference is null || generated is null))
        {
            throw new InvalidOperationException("Provide either '--suite <docx|pptx>' or both '--ref <path>' and '--gen <path>'.");
        }

        if (generate && !string.Equals(suite, "pptx", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("--generate only applies to '--suite pptx'.");
        }

        if (fontPath is not null && !generate)
        {
            throw new InvalidOperationException("--font-path only makes sense together with --generate.");
        }

        output ??= suite switch
        {
            not null when string.Equals(suite, "docx", StringComparison.OrdinalIgnoreCase) => "examples/output/visual-diff/docx",
            not null when string.Equals(suite, "pptx", StringComparison.OrdinalIgnoreCase) => "examples/output/visual-diff/pptx",
            not null => throw new InvalidOperationException($"Unknown suite '{suite}'. Supported suites: docx, pptx."),
            _ => throw new InvalidOperationException("--out is required when comparing an arbitrary pair.")
        };

        return new CliOptions(CliMode.Run, suite, reference, generated, output, name, dpi, generate, fontPath, null, baseline, margin);
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
          visual-diff --suite <docx|pptx> [--out <dir>] [--dpi 150] [--generate] [--font-path <dir>]
                      [--baseline <metrics.json>] [--margin 0.05]
          visual-diff --ref <ref.pdf|ref.png|ref-dir> --gen <gen.pdf|gen.png|gen-dir>
                      --out <dir> [--name <name>] [--dpi 150] [--baseline <metrics.json>] [--margin 0.05]
          visual-diff --check <metrics.json> --baseline <metrics.json> [--margin 0.05]
          visual-diff --probe
          visual-diff --help

        Suites (run from the repository root):
          docx   examples/REF/DOCX/*.pdf  vs  examples/output/ref/docx/*.pdf
          pptx   examples/REF/PPTX/{REMOVED,pres-pro}.pdf  vs  examples/output/ref/pptx/*.pdf
                 Missing generated PDFs can be built on demand with --generate (invokes
                 tools/convert-pptx) or manually:
                   dotnet run --project tools/convert-pptx -- examples/REF/PPTX/<deck>.pptx \
                     examples/output/ref/pptx/<deck>.pdf --format pdf

        Options:
          --suite <name>     Run a built-in comparison suite (docx, pptx).
          --ref <path>       Reference PDF, single PNG, or directory of PNGs.
          --gen <path>       Generated PDF, single PNG, or directory of PNGs (same kind as --ref).
                             PNG inputs skip the poppler rendering step entirely.
          --out <path>       Output directory. Defaults to examples/output/visual-diff/<suite> for suites.
          --name <name>      Display/report name for an arbitrary pair.
          --dpi <number>     Rasterization DPI for PDF inputs. Default: 150.
          --generate         (suite pptx) Build missing generated PDFs via tools/convert-pptx.
          --font-path <dir>  Extra font directory passed to tools/convert-pptx with --generate.
          --check <file>     Threshold-check an existing metrics.json against --baseline.
          --baseline <file>  Baseline metrics.json. After a run: check the fresh metrics against it.
          --margin <number>  Absolute margin on normalized RMSE (0.05 = 5 percentage points).
                             Default: 0.05.
          --probe            Print external-tool availability (poppler, ImageMagick) and exit.
          --help             Show this help text.

        Exit codes:
          0  success (report written; thresholds, if checked, satisfied)
          1  operational or usage error (missing tools, missing files, bad arguments)
          2  threshold check ran and at least one page/document exceeded baseline + margin

        Note: without --baseline/--check the tool is report-only and never fails on RMSE values.
        """);
    }
}
