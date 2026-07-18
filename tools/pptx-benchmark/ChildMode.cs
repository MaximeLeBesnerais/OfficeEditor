using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using OfficeEditor.Core.Services;
using PptxEditor.Core.Builders;

namespace PptxBenchmark;

/// <summary>
/// Child-mode benchmark: a fresh process measures 1 cold iteration followed by
/// N warm iterations of Open → ExportThumbnails (PNG @150ppi) → ExportToPdf,
/// emitting machine-parseable RESULT lines for the parent process.
/// </summary>
public static class ChildMode
{
    private static readonly TimeSpan ChildTimeout = TimeSpan.FromMinutes(30);

    /// <summary>Runs inside the fresh child process. Prints RESULT key=value lines.</summary>
    public static int Run(string deckPath, int runs)
    {
        // Deck bytes are read once, untimed — only PresentationBuilder.Open is measured.
        var bytes = File.ReadAllBytes(deckPath);
        var slideCount = 0;

        for (var i = 0; i <= runs; i++) // iteration 0 = cold, 1..runs = warm
        {
            var openTimer = Stopwatch.StartNew();
            using var builder = PresentationBuilder.Open(bytes);
            openTimer.Stop();
            if (i == 0)
            {
                slideCount = builder.SlideCount;
            }
            Emit(i, "open", openTimer.Elapsed.TotalMilliseconds);

            var pngTimer = Stopwatch.StartNew();
            var pages = builder.ExportThumbnails(new ThumbnailOptions { Ppi = 150 });
            pngTimer.Stop();
            Emit(i, "png", pngTimer.Elapsed.TotalMilliseconds, pages.Length);
            EmitCompile(i, "png");

            var pdfTimer = Stopwatch.StartNew();
            builder.ExportToPdf();
            pdfTimer.Stop();
            Emit(i, "pdf", pdfTimer.Elapsed.TotalMilliseconds, 1);
            EmitCompile(i, "pdf");
        }

        Console.WriteLine($"RESULT slide-count={slideCount}");
        return 0;
    }

    /// <summary>Spawns a fresh benchmark process for one deck and parses its RESULT lines.</summary>
    public static DeckBenchmark RunInFreshProcess(string deckPath, int runs, string repoRoot)
    {
        var psi = new ProcessStartInfo("dotnet")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
            WorkingDirectory = repoRoot
        };
        psi.ArgumentList.Add("exec");
        psi.ArgumentList.Add(Assembly.GetExecutingAssembly().Location);
        psi.ArgumentList.Add("--child");
        psi.ArgumentList.Add(deckPath);
        psi.ArgumentList.Add(runs.ToString(CultureInfo.InvariantCulture));
        // Enable the TypstCompilerService timing hooks so compile backend time is attributed per stage.
        psi.Environment["OFFICEEDITOR_TIMING"] = "1";

        using var process = Process.Start(psi)
            ?? throw new InvalidOperationException("Failed to start child benchmark process.");

        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExitAsync().Wait(ChildTimeout))
        {
            try { process.Kill(entireProcessTree: true); } catch { /* best effort */ }
            throw new TimeoutException($"Child benchmark process timed out after {ChildTimeout.TotalMinutes:F0} minutes for {deckPath}.");
        }

        var stderr = stderrTask.GetAwaiter().GetResult();
        if (process.ExitCode != 0)
        {
            throw new InvalidOperationException(
                $"Child benchmark process exited with code {process.ExitCode} for {deckPath}.{Environment.NewLine}{stderr}");
        }

        return Parse(stdoutTask.GetAwaiter().GetResult(), deckPath, runs);
    }

    private static void Emit(int iteration, string stage, double milliseconds, int? pages = null)
    {
        var line = string.Create(CultureInfo.InvariantCulture, $"RESULT iter={iteration} stage={stage} ms={milliseconds:F1}");
        if (pages.HasValue)
        {
            line += string.Create(CultureInfo.InvariantCulture, $" pages={pages.Value}");
        }
        Console.WriteLine(line);
    }

    private static void EmitCompile(int iteration, string stage)
    {
        // Each export performs exactly one Compile; drain after the stage so the record
        // (including one-time backend probe cost on the cold iteration) is attributed to it.
        foreach (var timing in TypstCompilerService.DrainTimings())
        {
            Console.WriteLine(string.Create(CultureInfo.InvariantCulture,
                $"RESULT iter={iteration} stage={stage}.compile backend={timing.Backend} ms={timing.TotalMilliseconds:F1}"));
        }
    }

    private static DeckBenchmark Parse(string stdout, string deckPath, int runs)
    {
        var result = new DeckBenchmark { DeckPath = deckPath, SlideCount = 0 };
        for (var i = 0; i <= runs; i++)
        {
            result.Iterations.Add(new Iteration());
        }

        foreach (var line in stdout.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (!line.StartsWith("RESULT ", StringComparison.Ordinal))
            {
                continue;
            }

            var fields = line["RESULT ".Length..]
                .Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .Select(part => part.Split('=', 2))
                .Where(parts => parts.Length == 2)
                .ToDictionary(parts => parts[0], parts => parts[1], StringComparer.Ordinal);

            if (fields.TryGetValue("slide-count", out var slideCount))
            {
                result.SlideCount = int.Parse(slideCount, CultureInfo.InvariantCulture);
                continue;
            }

            var iteration = int.Parse(fields["iter"], CultureInfo.InvariantCulture);
            var stage = fields["stage"];
            var ms = double.Parse(fields["ms"], CultureInfo.InvariantCulture);
            var target = result.Iterations[iteration];

            if (stage.EndsWith(".compile", StringComparison.Ordinal))
            {
                target.Compiles[stage[..^".compile".Length]] = new CompileMeasurement(fields["backend"], ms);
            }
            else
            {
                target.StageMs[stage] = ms;
                if (fields.TryGetValue("pages", out var pages))
                {
                    target.StagePages[stage] = int.Parse(pages, CultureInfo.InvariantCulture);
                }
            }
        }

        return result;
    }
}
