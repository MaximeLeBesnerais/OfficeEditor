using PptxEditor.Core.Generation.Fixtures;

namespace DocxEditor.Tests.Generation.Fixtures;

/// <summary>
/// Parity gate logic (thresholds per primitive): per-page and
/// per-document-average breaches, unavailable metrics, unconfigured fixtures, page-count
/// mismatches and the exact boundary — deterministic, renderer-free.
/// </summary>
public sealed class ParityEvaluatorTests
{
    private static readonly IReadOnlyDictionary<string, double> Thresholds =
        new Dictionary<string, double>(StringComparer.Ordinal) { ["rect-radii"] = 0.04, ["shadow"] = 0.12 };

    private static ParityDocumentSample Doc(string name, params double?[] pageRmse) =>
        new(name, false, pageRmse.Select((v, i) => new ParityPageSample(i + 1, v)).ToList());

    [Fact]
    public void Evaluate_WithinThreshold_Passes()
    {
        var failures = ParityEvaluator.Evaluate(
            [Doc("rect-radii", 0.01, 0.02, 0.039)], Thresholds);

        Assert.Empty(failures);
    }

    [Fact]
    public void Evaluate_ValueExactlyAtThreshold_Passes()
    {
        var failures = ParityEvaluator.Evaluate(
            [Doc("rect-radii", 0.04)], Thresholds);

        Assert.Empty(failures);
    }

    [Fact]
    public void Evaluate_PageAboveThreshold_FailsWithFixturePageAndValues()
    {
        var failures = ParityEvaluator.Evaluate(
            [Doc("rect-radii", 0.01, 0.041)], Thresholds);

        var failure = Assert.Single(failures);
        Assert.Contains("[rect-radii] page 2", failure);
        Assert.Contains("4.1%", failure);
        Assert.Contains("threshold 4%", failure);
    }

    [Fact]
    public void Evaluate_AverageAboveThreshold_Fails()
    {
        // Both pages breach; the average line summarizes the deck (avg > t requires at
        // least one page > t, so an average breach always accompanies page breaches).
        var failures = ParityEvaluator.Evaluate(
            [Doc("rect-radii", 0.05, 0.05)], Thresholds);

        Assert.Equal(3, failures.Count);
        Assert.Contains("[rect-radii] page 1", failures[0]);
        Assert.Contains("[rect-radii] page 2", failures[1]);
        Assert.Contains("average", failures[2]);
        Assert.Contains("5%", failures[2]);
    }

    [Fact]
    public void Evaluate_MissingMetric_Fails()
    {
        var failures = ParityEvaluator.Evaluate(
            [Doc("rect-radii", 0.01, null)], Thresholds);

        var failure = Assert.Single(failures);
        Assert.Contains("page 2", failure);
        Assert.Contains("metric unavailable", failure);
    }

    [Fact]
    public void Evaluate_FixtureWithoutThreshold_Fails()
    {
        var failures = ParityEvaluator.Evaluate(
            [Doc("unknown-fixture", 0.0)], Thresholds);

        var failure = Assert.Single(failures);
        Assert.Contains("unknown-fixture", failure);
        Assert.Contains("no parity threshold configured", failure);
    }

    [Fact]
    public void Evaluate_PageCountMismatch_Fails()
    {
        var failures = ParityEvaluator.Evaluate(
            [new ParityDocumentSample("shadow", true, [new ParityPageSample(1, 0.01)])], Thresholds);

        var failure = Assert.Single(failures);
        Assert.Contains("page count mismatch", failure);
    }

    [Fact]
    public void Evaluate_PerFixtureThresholds_AreIndependent()
    {
        // 5% breaches rect-radii (4%) but not shadow (12%).
        var failures = ParityEvaluator.Evaluate(
            [Doc("rect-radii", 0.05), Doc("shadow", 0.05)], Thresholds);

        Assert.NotEmpty(failures);
        Assert.All(failures, f => Assert.Contains("rect-radii", f));
        Assert.DoesNotContain(failures, f => f.Contains("shadow", StringComparison.Ordinal));
    }

    [Fact]
    public void Evaluate_InvalidThreshold_Throws()
    {
        var bad = new Dictionary<string, double> { ["x"] = 0 };
        Assert.Throws<ArgumentException>(() => ParityEvaluator.Evaluate([Doc("x", 0.0)], bad));
    }

    [Fact]
    public void Evaluate_NullArguments_Throw()
    {
        Assert.Throws<ArgumentNullException>(() => ParityEvaluator.Evaluate(null!, Thresholds));
        Assert.Throws<ArgumentNullException>(() => ParityEvaluator.Evaluate([], null!));
    }
}
