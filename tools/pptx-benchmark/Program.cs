using System.Globalization;
using PptxBenchmark;

// pptx-benchmark — pre-optimization latency ground truth for the PPTX render pipeline.
// Parent mode: benchmarks both REF decks (each in a fresh child process) + LibreOffice leg, writes a markdown report.
// Child mode (--child): single fresh-process measurement pass, emits parseable RESULT lines.

var runs = 5;
string? outPath = null;
var updateBaseline = false;
string? childDeck = null;

for (var i = 0; i < args.Length; i++)
{
    switch (args[i])
    {
        case "--runs" when i + 1 < args.Length:
            runs = int.Parse(args[++i], CultureInfo.InvariantCulture);
            break;
        case "--out" when i + 1 < args.Length:
            outPath = args[++i];
            break;
        case "--update-baseline":
            updateBaseline = true;
            break;
        case "--child" when i + 2 < args.Length:
            childDeck = args[++i];
            runs = int.Parse(args[++i], CultureInfo.InvariantCulture);
            break;
        case "--help" or "-h":
            PrintUsage();
            return 0;
        default:
            Console.Error.WriteLine($"Unknown argument: {args[i]}");
            PrintUsage();
            return 2;
    }
}

if (childDeck is not null)
{
    return ChildMode.Run(childDeck, runs);
}

if (runs < 1)
{
    Console.Error.WriteLine("--runs must be >= 1");
    return 2;
}

var repoRoot = FindRepoRoot();
var deckPaths = new[]
{
    Path.Combine(repoRoot, "examples", "REF", "PPTX", "REMOVED.pptx"),
    Path.Combine(repoRoot, "examples", "REF", "PPTX", "pres-pro.pptx"),
};

outPath ??= Path.Combine(repoRoot, "examples", "output", "benchmark", "report.md");

Console.WriteLine($"Repo root: {repoRoot}");
Console.WriteLine($"Warm runs per deck: {runs} (plus 1 cold run each, in a fresh process)");

var environment = EnvironmentInfo.Collect();
Console.WriteLine($"TypstBridge: {environment.TypstBridgeVersion}");

var results = new List<DeckBenchmark>();
foreach (var deckPath in deckPaths)
{
    if (!File.Exists(deckPath))
    {
        Console.Error.WriteLine($"WARNING: deck not found, skipping: {deckPath}");
        continue;
    }

    Console.WriteLine($"Benchmarking {Path.GetFileName(deckPath)} in a fresh process...");
    var result = ChildMode.RunInFreshProcess(deckPath, runs, repoRoot);
    Console.WriteLine($"  {result.SlideCount} slides; cold preview path " +
        $"{result.ColdMs("open") + result.ColdMs("png"):F1} ms, warm median " +
        $"{result.WarmMedianMs("open") + result.WarmMedianMs("png"):F1} ms");
    results.Add(result);
}

Console.WriteLine("LibreOffice leg...");
var libreOffice = LibreOfficeLeg.Run(deckPaths.Where(File.Exists).ToArray(), runs);
Console.WriteLine(libreOffice.Available
    ? $"  soffice: {libreOffice.Version ?? libreOffice.SofficePath}"
    : $"  {libreOffice.SkipReason}");

var report = ReportWriter.Build(environment, results, libreOffice, runs);
WriteFile(outPath, report);
Console.WriteLine($"Report written to {outPath}");

if (updateBaseline)
{
    var baselinePath = Path.Combine(repoRoot, "tools", "pptx-benchmark", "baselines", "baseline-pre-optimization.md");
    WriteFile(baselinePath, report);
    Console.WriteLine($"Tracked baseline copy written to {baselinePath}");
}

return 0;

static void PrintUsage()
{
    Console.WriteLine("Usage: pptx-benchmark [--runs N] [--out <report.md>] [--update-baseline]");
    Console.WriteLine("  --runs N           warm iterations per deck (median reported); default 5");
    Console.WriteLine("  --out <path>       report output; default examples/output/benchmark/report.md");
    Console.WriteLine("  --update-baseline  also copy the report to tools/pptx-benchmark/baselines/baseline-pre-optimization.md");
}

static void WriteFile(string path, string contents)
{
    var fullPath = Path.GetFullPath(path);
    Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
    File.WriteAllText(fullPath, contents);
}

static string FindRepoRoot()
{
    // Resolve robustly from the tool's own location first (bin/Debug/net9.0 → repo root),
    // then fall back to the current working directory.
    foreach (var start in new[] { AppContext.BaseDirectory, Directory.GetCurrentDirectory() })
    {
        var directory = new DirectoryInfo(start);
        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "examples", "REF", "PPTX", "REMOVED.pptx")))
            {
                return directory.FullName;
            }
            directory = directory.Parent;
        }
    }

    throw new DirectoryNotFoundException(
        "Could not locate the repository root (no ancestor directory contains examples/REF/PPTX/REMOVED.pptx).");
}
