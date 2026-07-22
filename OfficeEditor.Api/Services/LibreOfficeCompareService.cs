using System.Diagnostics;

namespace OfficeEditor.Api.Services;

/// <summary>One-shot probe of the external tools backing the LibreOffice compare leg.</summary>
public sealed record LibreOfficeProbe(
    bool Available,
    string? SofficePath,
    string? Version,
    bool PdfToPpmAvailable,
    string? SkipReason);

/// <summary>
/// Outcome of one LibreOffice compare render: PDF conversion always reported when it
/// happened; PNG pages only when pdftoppm is installed and rasterization succeeded.
/// Never throws for missing tools or conversion failures — failures ride on
/// <see cref="Error"/>.
/// </summary>
public sealed record LibreOfficeRenderResult(
    bool Available,
    string? Version,
    bool PdfToPpmAvailable,
    double ConversionMilliseconds,
    double? RasterizationMilliseconds,
    double TotalMilliseconds,
    IReadOnlyList<byte[]> PngPages,
    byte[]? PdfBytes,
    string? Error);

public interface ILibreOfficeCompareService
{
    /// <summary>Probes for soffice + pdftoppm. Cached for process lifetime: installed
    /// desktop tools realistically don't appear mid-process, so the discovery cost
    /// (PATH scan + `--version` subprocess) is paid once.</summary>
    LibreOfficeProbe Probe();

    /// <summary>
    /// Converts <paramref name="pptxBytes"/> to PDF via headless LibreOffice (fresh temp
    /// user profile per call) and, when pdftoppm is available, rasterizes the PDF to one
    /// PNG per slide at <paramref name="ppi"/>. Never throws for missing tools or
    /// conversion failures.
    /// </summary>
    LibreOfficeRenderResult RenderDeck(byte[] pptxBytes, string fileName, int ppi);
}

/// <summary>
/// LibreOffice headless compare leg, ported from tools/pptx-benchmark/LibreOfficeLeg.cs
/// (logic copied, not referenced): soffice discovery, fresh-profile PDF conversion with
/// a 5-minute timeout and process-tree kill on timeout, and pdftoppm rasterization.
/// soffice runs are serialized on a semaphore — concurrent compares would otherwise
/// fight over CPU and profile directories.
/// </summary>
public sealed class LibreOfficeCompareService : ILibreOfficeCompareService
{
    private static readonly TimeSpan ConversionTimeout = TimeSpan.FromMinutes(5);
    private static readonly TimeSpan VersionTimeout = TimeSpan.FromSeconds(30);

    private static readonly string[] WellKnownSofficePaths =
    {
        "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    };

    private static readonly string[] WellKnownPdfToPpmPaths =
    {
        "/opt/homebrew/bin/pdftoppm",
        "/usr/local/bin/pdftoppm"
    };

    private readonly Lazy<LibreOfficeProbe> _probe = new(ComputeProbe);
    private readonly SemaphoreSlim _sofficeGate = new(1, 1);

    public LibreOfficeProbe Probe() => _probe.Value;

    public LibreOfficeRenderResult RenderDeck(byte[] pptxBytes, string fileName, int ppi)
    {
        ArgumentNullException.ThrowIfNull(pptxBytes);
        ArgumentException.ThrowIfNullOrWhiteSpace(fileName);

        var probe = Probe();
        if (!probe.Available)
        {
            return new LibreOfficeRenderResult(
                Available: false,
                Version: null,
                PdfToPpmAvailable: probe.PdfToPpmAvailable,
                ConversionMilliseconds: 0,
                RasterizationMilliseconds: null,
                TotalMilliseconds: 0,
                PngPages: [],
                PdfBytes: null,
                Error: probe.SkipReason);
        }

        _sofficeGate.Wait();
        var workDirectory = CreateTempDirectory("officeeditor-lo-");
        var profileDirectory = CreateTempDirectory("officeeditor-lo-profile-");
        try
        {
            var safeFileName = SanitizeFileName(fileName);
            var deckPath = Path.Combine(workDirectory, safeFileName + ".pptx");
            File.WriteAllBytes(deckPath, pptxBytes);

            double conversionMs;
            string pdfPath;
            try
            {
                (conversionMs, pdfPath) = RunConversion(probe.SofficePath!, deckPath, profileDirectory, workDirectory);
            }
            catch (Exception ex) when (ex is TimeoutException or InvalidOperationException)
            {
                return new LibreOfficeRenderResult(
                    Available: true,
                    Version: probe.Version,
                    PdfToPpmAvailable: probe.PdfToPpmAvailable,
                    ConversionMilliseconds: 0,
                    RasterizationMilliseconds: null,
                    TotalMilliseconds: 0,
                    PngPages: [],
                    PdfBytes: null,
                    Error: ex.Message);
            }

            var pdfBytes = File.ReadAllBytes(pdfPath);

            if (!probe.PdfToPpmAvailable)
            {
                return new LibreOfficeRenderResult(
                    Available: true,
                    Version: probe.Version,
                    PdfToPpmAvailable: false,
                    ConversionMilliseconds: conversionMs,
                    RasterizationMilliseconds: null,
                    TotalMilliseconds: conversionMs,
                    PngPages: [],
                    PdfBytes: pdfBytes,
                    Error: null);
            }

            double? rasterizationMs = null;
            IReadOnlyList<byte[]> pages = [];
            string? error = null;
            try
            {
                (rasterizationMs, pages) = Rasterize(pdfPath, workDirectory, ppi);
            }
            catch (Exception ex) when (ex is TimeoutException or InvalidOperationException)
            {
                // Rasterization failure still returns the PDF.
                error = ex.Message;
            }

            return new LibreOfficeRenderResult(
                Available: true,
                Version: probe.Version,
                PdfToPpmAvailable: true,
                ConversionMilliseconds: conversionMs,
                RasterizationMilliseconds: rasterizationMs,
                TotalMilliseconds: conversionMs + (rasterizationMs ?? 0),
                PngPages: pages,
                PdfBytes: pdfBytes,
                Error: error);
        }
        finally
        {
            _sofficeGate.Release();
            DeleteDirectory(workDirectory);
            DeleteDirectory(profileDirectory);
        }
    }

    private static LibreOfficeProbe ComputeProbe()
    {
        var soffice = FindOnPath(OperatingSystem.IsWindows() ? "soffice.exe" : "soffice", WellKnownSofficePaths);
        var pdfToPpm = FindOnPath(OperatingSystem.IsWindows() ? "pdftoppm.exe" : "pdftoppm", WellKnownPdfToPpmPaths) is not null;

        if (soffice is null)
        {
            return new LibreOfficeProbe(
                Available: false,
                SofficePath: null,
                Version: null,
                PdfToPpmAvailable: pdfToPpm,
                SkipReason: "LibreOffice not available on this machine " +
                    "(soffice not found on PATH or at /Applications/LibreOffice.app/Contents/MacOS/soffice) — compare leg skipped.");
        }

        return new LibreOfficeProbe(
            Available: true,
            SofficePath: soffice,
            Version: QueryVersion(soffice),
            PdfToPpmAvailable: pdfToPpm,
            SkipReason: null);
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
            if (!process.WaitForExitAsync().Wait(VersionTimeout))
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

    private static (double Milliseconds, string PdfPath) RunConversion(
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
            throw new TimeoutException(
                $"soffice conversion timed out after {ConversionTimeout.TotalMinutes:F0} minutes for {Path.GetFileName(deckPath)}.");
        }
        timer.Stop();

        _ = stdoutTask;
        var stderr = stderrTask.GetAwaiter().GetResult();
        var expectedPdf = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(deckPath) + ".pdf");
        if (process.ExitCode != 0 || !File.Exists(expectedPdf))
        {
            throw new InvalidOperationException(
                $"soffice conversion failed (exit {process.ExitCode}): {TrimForError(stderr)}");
        }

        return (timer.Elapsed.TotalMilliseconds, expectedPdf);
    }

    private static (double Milliseconds, IReadOnlyList<byte[]> Pages) Rasterize(
        string pdfPath, string workDirectory, int ppi)
    {
        var outputPrefix = Path.Combine(workDirectory, "slide");
        var psi = new ProcessStartInfo(FindOnPath(
            OperatingSystem.IsWindows() ? "pdftoppm.exe" : "pdftoppm", WellKnownPdfToPpmPaths)!)
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        psi.ArgumentList.Add("-png");
        psi.ArgumentList.Add("-r");
        psi.ArgumentList.Add(ppi.ToString(System.Globalization.CultureInfo.InvariantCulture));
        psi.ArgumentList.Add(pdfPath);
        psi.ArgumentList.Add(outputPrefix);

        var timer = Stopwatch.StartNew();
        using var process = Process.Start(psi)
            ?? throw new InvalidOperationException("Failed to start pdftoppm.");
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExitAsync().Wait(ConversionTimeout))
        {
            try { process.Kill(entireProcessTree: true); } catch { /* best effort */ }
            throw new TimeoutException("pdftoppm rasterization timed out.");
        }
        timer.Stop();

        _ = stdoutTask;
        var stderr = stderrTask.GetAwaiter().GetResult();
        if (process.ExitCode != 0)
        {
            throw new InvalidOperationException(
                $"pdftoppm rasterization failed (exit {process.ExitCode}): {TrimForError(stderr)}");
        }

        // Natural sort: pdftoppm zero-pads page numbers only when the document has
        // >9 pages (slide-9.png vs slide-09.png), so plain ordinal order breaks.
        var pageFiles = Directory.GetFiles(workDirectory, "slide-*.png")
            .OrderBy(ExtractPageNumber)
            .ThenBy(f => f, StringComparer.Ordinal)
            .ToList();

        if (pageFiles.Count == 0)
        {
            throw new InvalidOperationException("pdftoppm produced no PNG pages.");
        }

        return (timer.Elapsed.TotalMilliseconds, pageFiles.Select(File.ReadAllBytes).ToList());
    }

    private static long ExtractPageNumber(string path)
    {
        var name = Path.GetFileNameWithoutExtension(path);
        var dash = name.LastIndexOf('-');
        return dash >= 0 && long.TryParse(name[(dash + 1)..], out var n) ? n : long.MaxValue;
    }

    private static string SanitizeFileName(string fileName)
    {
        var baseName = Path.GetFileNameWithoutExtension(fileName);
        var invalid = Path.GetInvalidFileNameChars();
        var chars = baseName.Select(c => invalid.Contains(c) ? '_' : c).ToArray();
        var sanitized = new string(chars);
        return string.IsNullOrWhiteSpace(sanitized) ? "deck" : sanitized;
    }

    private static string TrimForError(string stderr)
    {
        var trimmed = stderr.Trim();
        return trimmed.Length <= 500 ? trimmed : trimmed[..500] + "…";
    }

    private static string CreateTempDirectory(string prefix)
    {
        var directory = Path.Combine(Path.GetTempPath(), prefix + Guid.NewGuid().ToString("N"));
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
