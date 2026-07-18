namespace VisualDiff;

/// <summary>
/// Threshold-check wrapper around a metrics.json report. Exits 2 when any
/// per-page or per-document average normalized RMSE exceeds baseline + margin,
/// so CI gates can fail on visual regressions while the core comparison stays
/// report-only (exit 0).
/// </summary>
internal static class BaselineCheck
{
    public const int ExitPass = 0;
    public const int ExitBreach = 2;

    public static int Run(string metricsPath, string baselinePath, double margin)
    {
        if (!File.Exists(metricsPath))
        {
            throw new FileNotFoundException($"Metrics file not found: {metricsPath}. Run a comparison first.", metricsPath);
        }

        if (!File.Exists(baselinePath))
        {
            throw new FileNotFoundException(
                $"Baseline metrics not found: {baselinePath}\n"
                + "Generate one on a machine with poppler + ImageMagick, e.g.:\n"
                + "  dotnet run --project tools/visual-diff -- --suite pptx\n"
                + $"  cp examples/output/visual-diff/pptx/metrics.json {baselinePath}\n"
                + "Baselines are machine-dependent (fonts, rendering backend) - see tools/visual-diff/baselines/README.md.",
                baselinePath);
        }

        MetricsReport current = ReportWriter.LoadReport(metricsPath);
        MetricsReport baseline = ReportWriter.LoadReport(baselinePath);

        List<string> failures = [];
        Dictionary<string, DocumentMetrics> baselineDocs = baseline.Documents
            .GroupBy(d => d.Name)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.Ordinal);

        foreach (DocumentMetrics document in current.Documents)
        {
            if (!baselineDocs.TryGetValue(document.Name, out DocumentMetrics? baselineDoc))
            {
                failures.Add($"[{document.Name}] no matching document in baseline.");
                continue;
            }

            if (document.PageCountMismatch)
            {
                failures.Add($"[{document.Name}] page count mismatch: reference={document.ReferencePageCount}, generated={document.GeneratedPageCount}.");
            }

            Dictionary<int, PageMetrics> baselinePages = baselineDoc.Pages
                .GroupBy(p => p.PageNumber)
                .ToDictionary(g => g.Key, g => g.First());

            List<double> currentValues = [];
            List<double> baselineValues = [];

            foreach (PageMetrics page in document.Pages)
            {
                string label = $"[{document.Name}] page {page.PageNumber}";
                if (!baselinePages.TryGetValue(page.PageNumber, out PageMetrics? baselinePage))
                {
                    failures.Add($"{label}: page not present in baseline.");
                    continue;
                }

                if (page.NormalizedRmse is not double currentValue)
                {
                    failures.Add($"{label}: current metric unavailable (raw: '{page.RawMetric}').");
                    continue;
                }

                if (baselinePage.NormalizedRmse is not double baselineValue)
                {
                    failures.Add($"{label}: baseline metric unavailable (raw: '{baselinePage.RawMetric}').");
                    continue;
                }

                currentValues.Add(currentValue);
                baselineValues.Add(baselineValue);
                double limit = baselineValue + margin;
                if (currentValue > limit)
                {
                    failures.Add($"{label}: normalized RMSE {Pct(currentValue)} exceeds baseline {Pct(baselineValue)} + margin {Pct(margin)} (limit {Pct(limit)}).");
                }
            }

            if (currentValues.Count > 0)
            {
                double currentAvg = currentValues.Average();
                double baselineAvg = baselineValues.Average();
                double avgLimit = baselineAvg + margin;
                if (currentAvg > avgLimit)
                {
                    failures.Add($"[{document.Name}] average normalized RMSE {Pct(currentAvg)} exceeds baseline {Pct(baselineAvg)} + margin {Pct(margin)} (limit {Pct(avgLimit)}).");
                }
            }
        }

        foreach (string missing in baselineDocs.Keys.Except(current.Documents.Select(d => d.Name), StringComparer.Ordinal))
        {
            Console.WriteLine($"visual-diff check warning: baseline document '{missing}' is absent from the current run.");
        }

        Console.WriteLine($"visual-diff check: {Path.GetFullPath(metricsPath)}");
        Console.WriteLine($"  baseline: {Path.GetFullPath(baselinePath)} (generated {baseline.GeneratedAt:u}, DPI {baseline.Dpi})");
        Console.WriteLine($"  margin:   {Pct(margin)} on normalized RMSE (per page and per document average)");

        if (failures.Count > 0)
        {
            foreach (string failure in failures)
            {
                Console.Error.WriteLine($"  FAIL {failure}");
            }

            Console.Error.WriteLine($"visual-diff check FAILED: {failures.Count} threshold breach(es).");
            return ExitBreach;
        }

        Console.WriteLine("visual-diff check passed: all pages and document averages within baseline + margin.");
        return ExitPass;
    }

    private static string Pct(double normalized) => $"{normalized * 100:0.##}%";
}
