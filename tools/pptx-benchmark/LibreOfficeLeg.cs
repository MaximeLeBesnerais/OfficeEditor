using System.Diagnostics;

namespace PptxBenchmark;

/// <summary>Outcome of the LibreOffice comparison leg for a single deck.</summary>
public sealed record LibreOfficeDeckResult(string DeckName, double ColdMs, double WarmMedianMs);

/// <summary>Outcome of the LibreOffice leg across all decks (or a clean skip).</summary>
public sealed class LibreOfficeReport
{
    public bool Available { get; init; }
    public string? SofficePath { get; init; }
    public string? Version { get; init; }
    public string? SkipReason { get; init; }
    public List<LibreOfficeDeckResult> Decks { get; } = new();
}

/// <summary>
/// Times `soffice --headless --convert-to pdf` per deck: one cold run with a fresh
/// user profile, then N warm runs reusing one profile (median). Skips cleanly when
/// LibreOffice is not installed.
/// </summary>
public static class LibreOfficeLeg
{
    private static readonly TimeSpan ConversionTimeout = TimeSpan.FromMinutes(5);

    private static readonly string[] WellKnownPaths =
    {
        "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    };

    public static LibreOfficeReport Run(IReadOnlyList<string> deckPaths, int runs)
    {
        var soffice = FindSoffice();
        if (soffice is null)
        {
            return new LibreOfficeReport
            {
                Available = false,
                SkipReason = "LibreOffice not available on this machine " +
                    "(soffice not found on PATH or at /Applications/LibreOffice.app/Contents/MacOS/soffice) — leg skipped."
            };
        }

        var report = new LibreOfficeReport
        {
            Available = true,
            SofficePath = soffice,
            Version = QueryVersion(soffice)
        };

        foreach (var deckPath in deckPaths)
        {
            report.Decks.Add(RunDeck(soffice, deckPath, runs));
        }

        return report;
    }

    private static string? FindSoffice()
    {
        var exeName = OperatingSystem.IsWindows() ? "soffice.exe" : "soffice";
        var pathVariable = Environment.GetEnvironmentVariable("PATH");
        if (!string.IsNullOrEmpty(pathVariable))
        {
            foreach (var directory in pathVariable.Split(Path.PathSeparator))
            {
                string candidate;
                try
                {
                    candidate = Path.Combine(directory.Trim(), exeName);
                }
                catch
                {
                    continue; // malformed PATH entry
                }

                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }
        }

        foreach (var candidate in WellKnownPaths)
        {
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        return null;
    }

    private static string? QueryVersion(string soffice)
    {
        try
        {
            var psi = new ProcessStartInfo(soffice, "--version")
            {
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            using var process = Process.Start(psi)!;
            var stdoutTask = process.StandardOutput.ReadToEndAsync();
            var stderrTask = process.StandardError.ReadToEndAsync();
            if (!process.WaitForExitAsync().Wait(TimeSpan.FromSeconds(30)))
            {
                try { process.Kill(entireProcessTree: true); } catch { /* best effort */ }
                return null;
            }
            _ = stderrTask;
            var version = stdoutTask.GetAwaiter().GetResult().Split('\n', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()?.Trim();
            return string.IsNullOrEmpty(version) ? null : version;
        }
        catch
        {
            return null;
        }
    }

    private static LibreOfficeDeckResult RunDeck(string soffice, string deckPath, int runs)
    {
        // Cold: completely fresh user profile (first-start cost included).
        var coldProfile = CreateProfileDirectory();
        double coldMs;
        try
        {
            coldMs = RunConversion(soffice, deckPath, coldProfile);
        }
        finally
        {
            DeleteDirectory(coldProfile);
        }

        // Warm: one profile reused across runs, median of N.
        var warmProfile = CreateProfileDirectory();
        try
        {
            RunConversion(soffice, deckPath, warmProfile); // prime the profile, not timed
            var samples = new List<double>();
            for (var i = 0; i < runs; i++)
            {
                samples.Add(RunConversion(soffice, deckPath, warmProfile));
            }

            return new LibreOfficeDeckResult(Path.GetFileName(deckPath), coldMs, DeckBenchmark.Median(samples));
        }
        finally
        {
            DeleteDirectory(warmProfile);
        }
    }

    private static double RunConversion(string soffice, string deckPath, string profileDirectory)
    {
        var outputDirectory = Path.Combine(Path.GetTempPath(), "pptx-benchmark-lo-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(outputDirectory);
        try
        {
            var psi = new ProcessStartInfo(soffice)
            {
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            psi.ArgumentList.Add("--headless");
            psi.ArgumentList.Add("--norestore");
            psi.ArgumentList.Add($"-env:UserInstallation={new Uri(profileDirectory).AbsoluteUri}");
            psi.ArgumentList.Add("--convert-to");
            psi.ArgumentList.Add("pdf");
            psi.ArgumentList.Add("--outdir");
            psi.ArgumentList.Add(outputDirectory);
            psi.ArgumentList.Add(deckPath);

            var timer = Stopwatch.StartNew();
            using var process = Process.Start(psi)
                ?? throw new InvalidOperationException($"Failed to start soffice at {soffice}.");
            var stdoutTask = process.StandardOutput.ReadToEndAsync();
            var stderrTask = process.StandardError.ReadToEndAsync();
            if (!process.WaitForExitAsync().Wait(ConversionTimeout))
            {
                try { process.Kill(entireProcessTree: true); } catch { /* best effort */ }
                throw new TimeoutException($"soffice conversion timed out after {ConversionTimeout.TotalMinutes:F0} minutes for {deckPath}.");
            }
            timer.Stop();

            _ = stdoutTask;
            var stderr = stderrTask.GetAwaiter().GetResult();
            var expectedPdf = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(deckPath) + ".pdf");
            if (process.ExitCode != 0 || !File.Exists(expectedPdf))
            {
                throw new InvalidOperationException(
                    $"soffice conversion failed (exit {process.ExitCode}) for {deckPath}: {stderr}");
            }

            return timer.Elapsed.TotalMilliseconds;
        }
        finally
        {
            DeleteDirectory(outputDirectory);
        }
    }

    private static string CreateProfileDirectory()
    {
        var directory = Path.Combine(Path.GetTempPath(), "pptx-benchmark-lo-profile-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        return directory;
    }

    private static void DeleteDirectory(string directory)
    {
        try
        {
            Directory.Delete(directory, true);
        }
        catch
        {
            // Best effort cleanup of temp directories.
        }
    }
}
