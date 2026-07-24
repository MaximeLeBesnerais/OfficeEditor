using System.Diagnostics;

namespace PptxBenchmark;

/// <summary>Outcome of the LibreOffice comparison leg for a single deck.</summary>
public sealed record LibreOfficeDeckResult(
    string DeckName,
    double PdfColdMs,
    double PdfWarmMedianMs,
    double? RasterColdMs,
    double? RasterWarmMedianMs,
    int? RasterPageCount)
{
    /// <summary>PDF conversion + pdftoppm rasterization (null when pdftoppm is unavailable).</summary>
    public double? TotalColdMs => RasterColdMs is null ? null : PdfColdMs + RasterColdMs.Value;

    /// <summary>Warm median of PDF conversion + pdftoppm rasterization (null when pdftoppm is unavailable).</summary>
    public double? TotalWarmMedianMs => RasterWarmMedianMs is null ? null : PdfWarmMedianMs + RasterWarmMedianMs.Value;
}

/// <summary>Outcome of the LibreOffice leg across all decks (or a clean skip).</summary>
public sealed class LibreOfficeReport
{
    public bool Available { get; init; }
    public string? SofficePath { get; init; }
    public string? Version { get; init; }
    public string? SkipReason { get; init; }
    public bool PdfToPpmAvailable { get; init; }
    public string? RasterSkipReason { get; init; }
    public List<LibreOfficeDeckResult> Decks { get; } = new();
}

/// <summary>
/// Times `soffice --headless --convert-to pdf` per deck: one cold run with a fresh
/// user profile, then N warm runs reusing one profile (median). Each produced PDF is
/// then rasterized with `pdftoppm -png -r 150` (timed separately, same runs) so the
/// comparison covers the same per-slide PNG artifacts as the OfficeEditor leg.
/// Skips cleanly when LibreOffice is not installed; the rasterization half is reported
/// as unavailable (without failing) when pdftoppm is missing.
/// </summary>
public static class LibreOfficeLeg
{
    private static readonly TimeSpan ConversionTimeout = TimeSpan.FromMinutes(5);

    private static readonly string[] WellKnownSofficePaths =
    {
        "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    };

    private static readonly string[] WellKnownPdfToPpmPaths =
    {
        "/opt/homebrew/bin/pdftoppm",
        "/usr/local/bin/pdftoppm"
    };

    public static LibreOfficeReport Run(IReadOnlyList<string> deckPaths, int runs)
    {
        var soffice = FindOnPath(OperatingSystem.IsWindows() ? "soffice.exe" : "soffice", WellKnownSofficePaths);
        if (soffice is null)
        {
            return new LibreOfficeReport
            {
                Available = false,
                SkipReason = "LibreOffice not available on this machine " +
                    "(soffice not found on PATH or at /Applications/LibreOffice.app/Contents/MacOS/soffice) — leg skipped."
            };
        }

        var pdftoppm = FindOnPath(OperatingSystem.IsWindows() ? "pdftoppm.exe" : "pdftoppm", WellKnownPdfToPpmPaths);
        var report = new LibreOfficeReport
        {
            Available = true,
            SofficePath = soffice,
            Version = QueryVersion(soffice),
            PdfToPpmAvailable = pdftoppm is not null,
            RasterSkipReason = pdftoppm is null
                ? "pdftoppm (poppler) not found on PATH or at /opt/homebrew/bin or /usr/local/bin — rasterization half skipped."
                : null
        };

        foreach (var deckPath in deckPaths)
        {
            report.Decks.Add(RunDeck(soffice, pdftoppm, deckPath, runs));
        }

        return report;
    }

    private static string? FindOnPath(string exeName, string[] wellKnownPaths)
    {
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

        foreach (var candidate in wellKnownPaths)
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

    private static LibreOfficeDeckResult RunDeck(string soffice, string? pdftoppm, string deckPath, int runs)
    {
        // Cold: completely fresh user profile (first-start cost included).
        var coldProfile = CreateProfileDirectory();
        double pdfColdMs;
        double? rasterColdMs = null;
        int? rasterPageCount = null;
        try
        {
            var run = RunOnce(soffice, pdftoppm, deckPath, coldProfile);
            pdfColdMs = run.PdfMs;
            rasterColdMs = run.RasterMs;
            rasterPageCount = run.RasterPageCount;
        }
        finally
        {
            DeleteDirectory(coldProfile);
        }

        // Warm: one profile reused across runs, median of N per stage.
        var warmProfile = CreateProfileDirectory();
        try
        {
            RunOnce(soffice, pdftoppm, deckPath, warmProfile); // prime the profile, not timed
            var pdfSamples = new List<double>();
            var rasterSamples = new List<double>();
            for (var i = 0; i < runs; i++)
            {
                var run = RunOnce(soffice, pdftoppm, deckPath, warmProfile);
                pdfSamples.Add(run.PdfMs);
                if (run.RasterMs is { } rasterMs)
                {
                    rasterSamples.Add(rasterMs);
                }
            }

            return new LibreOfficeDeckResult(
                Path.GetFileName(deckPath),
                pdfColdMs,
                DeckBenchmark.Median(pdfSamples),
                rasterColdMs,
                rasterSamples.Count > 0 ? DeckBenchmark.Median(rasterSamples) : null,
                rasterPageCount);
        }
        finally
        {
            DeleteDirectory(warmProfile);
        }
    }

    /// <summary>One timed convert-to-PDF + (when available) timed pdftoppm rasterization.</summary>
    private static (double PdfMs, double? RasterMs, int? RasterPageCount) RunOnce(
        string soffice, string? pdftoppm, string deckPath, string profileDirectory)
    {
        var outputDirectory = Path.Combine(Path.GetTempPath(), "pptx-benchmark-lo-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(outputDirectory);
        try
        {
            var (pdfMs, pdfPath) = TimeConversion(soffice, deckPath, profileDirectory, outputDirectory);
            if (pdftoppm is null)
            {
                return (pdfMs, null, null);
            }

            var (rasterMs, pageCount) = TimeRasterization(pdftoppm, pdfPath, outputDirectory);
            return (pdfMs, rasterMs, pageCount);
        }
        finally
        {
            DeleteDirectory(outputDirectory);
        }
    }

    private static (double Ms, string PdfPath) TimeConversion(
        string soffice, string deckPath, string profileDirectory, string outputDirectory)
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

        return (timer.Elapsed.TotalMilliseconds, expectedPdf);
    }

    private static (double Ms, int PageCount) TimeRasterization(string pdftoppm, string pdfPath, string outputDirectory)
    {
        var outputPrefix = Path.Combine(outputDirectory, "slide");
        var psi = new ProcessStartInfo(pdftoppm)
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        psi.ArgumentList.Add("-png");
        psi.ArgumentList.Add("-r");
        psi.ArgumentList.Add("150");
        psi.ArgumentList.Add(pdfPath);
        psi.ArgumentList.Add(outputPrefix);

        var timer = Stopwatch.StartNew();
        using var process = Process.Start(psi)
            ?? throw new InvalidOperationException($"Failed to start pdftoppm at {pdftoppm}.");
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExitAsync().Wait(ConversionTimeout))
        {
            try { process.Kill(entireProcessTree: true); } catch { /* best effort */ }
            throw new TimeoutException($"pdftoppm rasterization timed out after {ConversionTimeout.TotalMinutes:F0} minutes for {pdfPath}.");
        }
        timer.Stop();

        _ = stdoutTask;
        var stderr = stderrTask.GetAwaiter().GetResult();
        if (process.ExitCode != 0)
        {
            throw new InvalidOperationException(
                $"pdftoppm rasterization failed (exit {process.ExitCode}) for {pdfPath}: {stderr}");
        }

        var pageCount = Directory.GetFiles(outputDirectory, "slide-*.png").Length;
        if (pageCount == 0)
        {
            throw new InvalidOperationException($"pdftoppm produced no PNG pages for {pdfPath}.");
        }

        return (timer.Elapsed.TotalMilliseconds, pageCount);
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
