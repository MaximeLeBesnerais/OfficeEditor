using System.Collections.Concurrent;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;

namespace OfficeEditor.Api.Services;

/// <summary>
/// Rasterized official-render pages for one whitelisted demo deck: the sibling PDF
/// export from online PowerPoint (the ground-truth render) turned into per-slide PNGs.
/// </summary>
public sealed record OfficialSlidesResult(string Name, int SlideCount, IReadOnlyList<byte[]> Pages);

public interface IOfficialRenderService
{
    /// <summary>
    /// Whether pdftoppm was found (PATH + well-known locations). Probed once and cached
    /// for process lifetime, same convention as the LibreOffice compare leg.
    /// </summary>
    bool PdfToPpmAvailable { get; }

    /// <summary>
    /// Returns the rasterized official-render pages for a whitelisted demo deck, false
    /// when the name is unknown, the sibling PDF is missing, pdftoppm is unavailable, or
    /// rasterization failed. Results are cached per deck in memory — REF files are
    /// immutable, so no invalidation is needed.
    /// </summary>
    bool TryGetOfficialSlides(string name, [NotNullWhen(true)] out OfficialSlidesResult? result);
}

/// <summary>
/// Official-render (PowerPoint ground truth) leg of the demo: locates the sibling PDF
/// export next to a whitelisted REF .pptx and rasterizes it with pdftoppm at 110 PPI
/// (same resolution as the demo render default). pdftoppm discovery mirrors
/// <see cref="LibreOfficeCompareService"/> (PATH + /opt/homebrew/bin + /usr/local/bin).
/// Rasterized pages are cached per deck; the first (cold) rasterization per deck is
/// timed and logged.
/// </summary>
public sealed class OfficialRenderService : IOfficialRenderService
{
    private const int RasterizationPpi = 110;
    private static readonly TimeSpan RasterizationTimeout = TimeSpan.FromMinutes(2);

    private static readonly string[] WellKnownPdfToPpmPaths =
    {
        "/opt/homebrew/bin/pdftoppm",
        "/usr/local/bin/pdftoppm"
    };

    private readonly IDemoDeckService _demoDeckService;
    private readonly ILogger<OfficialRenderService> _logger;
    private readonly Lazy<string?> _pdftoppmPath;
    private readonly ConcurrentDictionary<string, Lazy<OfficialSlidesResult?>> _cache =
        new(StringComparer.OrdinalIgnoreCase);

    public OfficialRenderService(IDemoDeckService demoDeckService, ILogger<OfficialRenderService> logger)
    {
        _demoDeckService = demoDeckService;
        _logger = logger;
        _pdftoppmPath = new Lazy<string?>(() => FindOnPath(
            OperatingSystem.IsWindows() ? "pdftoppm.exe" : "pdftoppm", WellKnownPdfToPpmPaths));
    }

    public bool PdfToPpmAvailable => _pdftoppmPath.Value is not null;

    public bool TryGetOfficialSlides(string name, [NotNullWhen(true)] out OfficialSlidesResult? result)
    {
        result = null;
        if (string.IsNullOrWhiteSpace(name) || !PdfToPpmAvailable)
        {
            return false;
        }

        if (_demoDeckService is not DemoDeckService concreteDemoDeckService
            || !concreteDemoDeckService.TryGetDeckFile(name, out var deckPath))
        {
            return false;
        }

        var pdfPath = Path.ChangeExtension(deckPath, ".pdf");
        if (!File.Exists(pdfPath))
        {
            return false;
        }

        // Lazy per deck: concurrent first requests for the same deck pay one
        // rasterization; a cached null means a failed rasterization (not retried).
        var lazy = _cache.GetOrAdd(name, _ => new Lazy<OfficialSlidesResult?>(() => Rasterize(name, pdfPath)));
        result = lazy.Value;
        return result is not null;
    }

    private OfficialSlidesResult? Rasterize(string name, string pdfPath)
    {
        var workDirectory = Path.Combine(Path.GetTempPath(), "officeeditor-official-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(workDirectory);
        try
        {
            var timer = Stopwatch.StartNew();
            var pages = RunPdfToPpm(pdfPath, workDirectory);
            timer.Stop();
            _logger.LogInformation(
                "Official-render rasterization for '{Deck}' completed in {ElapsedMs:F0} ms ({PageCount} pages).",
                name,
                timer.Elapsed.TotalMilliseconds,
                pages.Count);
            return new OfficialSlidesResult(name, pages.Count, pages);
        }
        catch (Exception ex) when (ex is TimeoutException or InvalidOperationException)
        {
            // Never throw for external-tool failures — the endpoint degrades to 404.
            _logger.LogWarning(ex, "Official-render rasterization for '{Deck}' failed.", name);
            return null;
        }
        finally
        {
            try
            {
                Directory.Delete(workDirectory, true);
            }
            catch
            {
                // Best effort cleanup of temp directories.
            }
        }
    }

    private IReadOnlyList<byte[]> RunPdfToPpm(string pdfPath, string workDirectory)
    {
        var outputPrefix = Path.Combine(workDirectory, "slide");
        var psi = new ProcessStartInfo(_pdftoppmPath.Value!)
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        psi.ArgumentList.Add("-png");
        psi.ArgumentList.Add("-r");
        psi.ArgumentList.Add(RasterizationPpi.ToString(System.Globalization.CultureInfo.InvariantCulture));
        psi.ArgumentList.Add(pdfPath);
        psi.ArgumentList.Add(outputPrefix);

        using var process = Process.Start(psi)
            ?? throw new InvalidOperationException("Failed to start pdftoppm.");
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExitAsync().Wait(RasterizationTimeout))
        {
            try { process.Kill(entireProcessTree: true); } catch { /* best effort */ }
            throw new TimeoutException("pdftoppm rasterization timed out.");
        }

        _ = stdoutTask;
        var stderr = stderrTask.GetAwaiter().GetResult();
        if (process.ExitCode != 0)
        {
            var trimmed = stderr.Trim();
            throw new InvalidOperationException(
                $"pdftoppm rasterization failed (exit {process.ExitCode}): " +
                (trimmed.Length <= 500 ? trimmed : trimmed[..500] + "…"));
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

        return pageFiles.Select(File.ReadAllBytes).ToList();
    }

    private static long ExtractPageNumber(string path)
    {
        var name = Path.GetFileNameWithoutExtension(path);
        var dash = name.LastIndexOf('-');
        return dash >= 0 && long.TryParse(name[(dash + 1)..], out var n) ? n : long.MaxValue;
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
}
