using PptxEditor.Core.Generation.Fixtures;

namespace VisualDiff;

/// <summary>
/// Per-primitive threshold gate for the gen parity suite (plan.md §5: "thresholds per
/// primitive"). Unlike <see cref="BaselineCheck"/> (machine-dependent baseline + margin),
/// thresholds are absolute authored RMSE ceilings per fixture. Exits 2 on any breach so
/// CI fails loudly; a missing metrics or thresholds file is an operational error (exit
/// 1), never a silent pass. All decision logic lives in the unit-tested
/// <see cref="ParityEvaluator"/>; this class is console glue.
/// </summary>
internal static class ThresholdCheck
{
    public const int ExitPass = 0;
    public const int ExitBreach = 2;

    public static int Run(string metricsPath, string thresholdsPath)
    {
        if (!File.Exists(metricsPath))
        {
            throw new FileNotFoundException($"Metrics file not found: {metricsPath}. Run a comparison first.", metricsPath);
        }

        IReadOnlyDictionary<string, double> thresholds = ParityThresholds.Load(thresholdsPath);
        MetricsReport current = ReportWriter.LoadReport(metricsPath);

        var samples = current.Documents.Select(document => new ParityDocumentSample(
                document.Name,
                document.PageCountMismatch,
                document.Pages.Select(page => new ParityPageSample(page.PageNumber, page.NormalizedRmse)).ToList()))
            .ToList();

        IReadOnlyList<string> failures = ParityEvaluator.Evaluate(samples, thresholds);

        foreach (string absent in thresholds.Keys
                     .Except(current.Documents.Select(d => d.Name), StringComparer.Ordinal))
        {
            Console.WriteLine($"visual-diff threshold warning: fixture '{absent}' has a threshold but is absent from the current run.");
        }

        Console.WriteLine($"visual-diff threshold check: {Path.GetFullPath(metricsPath)}");
        Console.WriteLine($"  thresholds: {Path.GetFullPath(thresholdsPath)} (per-primitive normalized-RMSE ceilings, per page and per document average)");

        if (failures.Count > 0)
        {
            foreach (string failure in failures)
            {
                Console.Error.WriteLine($"  FAIL {failure}");
            }

            Console.Error.WriteLine($"visual-diff threshold check FAILED: {failures.Count} threshold breach(es).");
            return ExitBreach;
        }

        Console.WriteLine("visual-diff threshold check passed: all fixtures within their per-primitive thresholds.");
        return ExitPass;
    }
}
