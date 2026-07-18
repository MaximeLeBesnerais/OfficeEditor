namespace PptxBenchmark;

/// <summary>Timing attributed to the Typst compile inside an export stage.</summary>
public sealed record CompileMeasurement(string Backend, double Milliseconds);

/// <summary>Measurements for one open → PNG → PDF iteration.</summary>
public sealed class Iteration
{
    public Dictionary<string, double> StageMs { get; } = new(StringComparer.Ordinal);
    public Dictionary<string, int> StagePages { get; } = new(StringComparer.Ordinal);
    public Dictionary<string, CompileMeasurement> Compiles { get; } = new(StringComparer.Ordinal);
}

/// <summary>All iterations for one deck. Iteration 0 is the cold (fresh process) run.</summary>
public sealed class DeckBenchmark
{
    public required string DeckPath { get; init; }
    public required int SlideCount { get; set; }
    public List<Iteration> Iterations { get; } = new();

    public string DeckName => Path.GetFileName(DeckPath);

    public double ColdMs(string stage) => Iterations[0].StageMs[stage];

    public double WarmMedianMs(string stage) =>
        Median(Iterations.Skip(1).Select(i => i.StageMs[stage]));

    public CompileMeasurement? ColdCompile(string stage) =>
        Iterations[0].Compiles.GetValueOrDefault(stage);

    public double? WarmMedianCompileMs(string stage)
    {
        var samples = Iterations.Skip(1)
            .Select(i => i.Compiles.GetValueOrDefault(stage))
            .Where(c => c is not null)
            .Select(c => c!.Milliseconds)
            .ToList();
        return samples.Count > 0 ? Median(samples) : null;
    }

    public static double Median(IEnumerable<double> values)
    {
        var sorted = values.OrderBy(v => v).ToList();
        if (sorted.Count == 0)
        {
            return 0;
        }
        var mid = sorted.Count / 2;
        return sorted.Count % 2 == 1 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
    }
}
