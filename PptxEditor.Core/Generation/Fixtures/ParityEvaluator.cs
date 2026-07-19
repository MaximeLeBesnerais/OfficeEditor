namespace PptxEditor.Core.Generation.Fixtures;

/// <summary>One page of one compared fixture render.</summary>
public sealed record ParityPageSample(int PageNumber, double? NormalizedRmse);

/// <summary>One compared fixture: ground-truth render vs Typst preview render.</summary>
public sealed record ParityDocumentSample(string Name, bool PageCountMismatch, IReadOnlyList<ParityPageSample> Pages);

/// <summary>
/// Pure parity gate (plan.md §5: "visual-diff RMSE per fixture … thresholds per
/// primitive"): evaluates per-page and per-document-average normalized RMSE against the
/// per-primitive thresholds and returns one loud failure line per breach. An empty
/// result means the suite is green. A value exactly at the threshold passes; anything
/// above fails. No renderers, no file IO — the visual-diff tool adapts its metrics.json
/// reports into <see cref="ParityDocumentSample"/>s and this module decides.
/// </summary>
public static class ParityEvaluator
{
    /// <summary>Evaluates all documents; returns the failure lines (empty = pass).</summary>
    /// <exception cref="ArgumentNullException"><paramref name="documents"/> or <paramref name="thresholds"/> is null.</exception>
    /// <exception cref="ArgumentException">A threshold value is outside (0, 1].</exception>
    public static IReadOnlyList<string> Evaluate(
        IEnumerable<ParityDocumentSample> documents,
        IReadOnlyDictionary<string, double> thresholds)
    {
        ArgumentNullException.ThrowIfNull(documents);
        ArgumentNullException.ThrowIfNull(thresholds);
        foreach ((string name, double threshold) in thresholds)
        {
            if (threshold is <= 0 or > 1)
            {
                throw new ArgumentException(
                    $"Threshold for '{name}' is {threshold}; normalized RMSE thresholds must be in (0, 1].",
                    nameof(thresholds));
            }
        }

        var failures = new List<string>();
        foreach (var document in documents)
        {
            if (!thresholds.TryGetValue(document.Name, out double threshold))
            {
                failures.Add($"[{document.Name}] no parity threshold configured for this fixture.");
                continue;
            }

            if (document.PageCountMismatch)
            {
                failures.Add($"[{document.Name}] page count mismatch between ground truth and preview renders.");
            }

            var values = new List<double>();
            foreach (var page in document.Pages)
            {
                string label = $"[{document.Name}] page {page.PageNumber}";
                if (page.NormalizedRmse is not double value)
                {
                    failures.Add($"{label}: metric unavailable (comparison failed or produced no RMSE).");
                    continue;
                }

                values.Add(value);
                if (value > threshold)
                {
                    failures.Add($"{label}: normalized RMSE {Pct(value)} exceeds the {document.Name} threshold {Pct(threshold)}.");
                }
            }

            if (values.Count > 0)
            {
                double average = values.Average();
                if (average > threshold)
                {
                    failures.Add($"[{document.Name}] average normalized RMSE {Pct(average)} exceeds the {document.Name} threshold {Pct(threshold)}.");
                }
            }
        }
        return failures;
    }

    private static string Pct(double normalized) => $"{normalized * 100:0.##}%";
}
