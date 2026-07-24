using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Nodes;
using PptxEditor.Core.Builders;

namespace OfficeEditor.Api.Services;

/// <summary>One entry of the demo deck catalog (slide count computed lazily, 0 when the file is missing).</summary>
public sealed record DemoDeckInfo(string Name, string FileName, string Description, int SlideCount);

/// <summary>Outcome of a timed whole-deck demo render (all slides in one Typst compile).</summary>
public sealed record DemoRenderResult(
    Guid DeckId,
    int SlideCount,
    double TotalMilliseconds,
    string Format,
    IReadOnlyList<byte[]> Pages);

/// <summary>
/// Outcome of the Typst compare leg: a timed whole-deck PNG render plus a best-effort
/// whole-deck PDF export, run concurrently. The PDF side is optional — on failure
/// <see cref="PdfError"/> carries the message and <see cref="PdfBytes"/>/
/// <see cref="PdfMilliseconds"/> are null, while the PNG result is always returned
/// (PNG render failures throw instead). <see cref="PngMilliseconds"/> and
/// <see cref="PdfMilliseconds"/> are honest per-phase times;
/// <see cref="TotalMilliseconds"/> is the wall-clock time of the parallel pair.
/// </summary>
public sealed record TypstLegResult(
    int SlideCount,
    double PngMilliseconds,
    double? PdfMilliseconds,
    double TotalMilliseconds,
    IReadOnlyList<byte[]> PngPages,
    byte[]? PdfBytes,
    string? PdfError);

public interface IDemoDeckService
{
    /// <summary>Returns the whitelisted demo decks, in catalog order.</summary>
    IReadOnlyList<DemoDeckInfo> ListDecks();

    /// <summary>
    /// Renders every slide of a whitelisted demo deck to <paramref name="format"/> in a
    /// single whole-deck Typst compile, times the render, and stores the source bytes as
    /// a deck session. <paramref name="format"/> must already be normalized to "svg" or
    /// "png" (<see cref="DeckPreviewValidators.TryNormalizeFormat"/>); anything else throws
    /// <see cref="ArgumentException"/>, as do unknown names.
    /// <see cref="FileNotFoundException"/> is thrown when the deck file is missing on disk.
    /// Render failures propagate — callers must not swallow them.
    /// </summary>
    DemoRenderResult RenderDeck(string name, int ppi, string format);

    /// <summary>
    /// Renders every slide of an arbitrary uploaded deck to <paramref name="format"/> in a
    /// single whole-deck Typst compile, times the render, and stores the source bytes as
    /// a deck session. <paramref name="format"/> must already be normalized to "svg" or
    /// "png" (<see cref="DeckPreviewValidators.TryNormalizeFormat"/>); anything else throws
    /// <see cref="ArgumentException"/>, as do unparseable deck bytes, empty decks, and
    /// decks over the demo render slide cap. Render failures propagate.
    /// </summary>
    DemoRenderResult RenderUploadedDeck(byte[] sourceBytes, string fileName, string format, int ppi);

    /// <summary>
    /// Runs the Typst compare leg on arbitrary deck bytes: a timed whole-deck PNG render
    /// (<see cref="PresentationBuilder.ExportThumbnails"/>, format "png") plus a separately
    /// timed, best-effort whole-deck PDF export (<see cref="PresentationBuilder.ExportToPdf"/>),
    /// run CONCURRENTLY. Each render opens its own builder instance (own converter, own
    /// GUID-suffixed temp directory, own Typst compile session — no shared mutable state,
    /// so the parallel legs cannot collide). <see cref="TypstLegResult.TotalMilliseconds"/>
    /// is the wall-clock time of the parallel pair; the per-phase fields stay per-leg.
    /// Unparseable deck bytes, empty decks, and decks over the demo render slide cap throw
    /// <see cref="ArgumentException"/>; PNG render failures propagate. A PDF export failure
    /// never kills the leg — it is reported via <see cref="TypstLegResult.PdfError"/>.
    /// <paramref name="ppi"/> must already be clamped by the caller
    /// (<see cref="DeckPreviewValidators.ClampPpi"/>); it is passed through as-is.
    /// </summary>
    TypstLegResult RenderTypstLeg(byte[] sourceBytes, string fileName, int ppi);

    /// <summary>
    /// Reads demo/demo-deck.json from the repository root and normalizes every relative
    /// "src" string property to a repo-root-relative forward-slash path (e.g.
    /// "demo/assets/dashboard.png"). Callers resolve it against the repository root —
    /// the Typst preview compile uses the repository root as its project root.
    /// Throws <see cref="FileNotFoundException"/> when the template file is missing.
    /// </summary>
    string GetDeckTemplateJson();
}

/// <summary>
/// Demo support: a hardcoded whitelist of REF decks (no arbitrary file access), a timed
/// whole-deck preview render through <see cref="PresentationBuilder.ExportThumbnails"/> (one
/// Typst compile per render), and the generation-template JSON with repo-root-relative
/// asset paths.
/// </summary>
public sealed class DemoDeckService : IDemoDeckService
{
    private const string TemplateRelativePath = "demo/demo-deck.json";

    /// <summary>
    /// Upload cap for the demo render: the whole response is a base64 previews payload,
    /// so arbitrarily large decks are rejected up front.
    /// </summary>
    internal const int MaxUploadSlides = 60;

    private static readonly IReadOnlyList<DemoDeckEntry> Catalog =
    [
        new(
            "northwind",
            "examples/REF/PPTX/northwind-demo.pptx",
            "Self-made 15-slide Northwind Labs demo deck (generated by OfficeEditor)"),
        new(
            "sales",
            "examples/REF/PPTX/sales_acceleration_deck.pptx",
            "16-slide sales deck with diagrams — primary REF"),
        new(
            "aetherlink",
            "examples/REF/PPTX/AetherLink-Glass-Shareholder-Overview.pptx",
            "15-slide styled glass deck"),
        new(
            "launch-review",
            "examples/REF/PPTX/northwind-launch-review.pptx",
            "12-slide pitch deck")
    ];

    private readonly IDeckSessionStore _sessionStore;
    private readonly string? _fontDirectory;

    /// <summary>
    /// <paramref name="fontDirectory"/> is an optional extra font search path (typically
    /// the host's system font directories, e.g. "/System/Library/Fonts:/Library/Fonts")
    /// forwarded to every thumbnail render so decks without embedded fonts render in their
    /// declared sans/serif families instead of Typst's embedded serif fallback. It is
    /// COMBINED with any embedded fonts by <see cref="PresentationBuilder"/> (embedded
    /// first) — never an override. Null keeps the previous embedded-only behavior.
    /// </summary>
    public DemoDeckService(IDeckSessionStore sessionStore, string? fontDirectory = null)
    {
        _sessionStore = sessionStore;
        _fontDirectory = fontDirectory;
    }

    public IReadOnlyList<DemoDeckInfo> ListDecks() =>
        Catalog.Select(entry => new DemoDeckInfo(
                entry.Name,
                Path.GetFileName(entry.RelativePath),
                entry.Description,
                entry.GetSlideCount()))
            .ToList();

    public DemoRenderResult RenderDeck(string name, int ppi, string format)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);

        // Defensive: the endpoint normalizes via DeckPreviewValidators.TryNormalizeFormat,
        // but the service requires already-normalized input (same contract as
        // DeckGenerationService.Generate).
        if (format is not ("svg" or "png"))
        {
            throw new ArgumentException(
                $"Format must be normalized to \"svg\" or \"png\" (got '{format}').",
                nameof(format));
        }

        var entry = Catalog.FirstOrDefault(e => string.Equals(e.Name, name, StringComparison.OrdinalIgnoreCase))
            ?? throw new ArgumentException(
                $"Unknown demo deck '{name}'. Valid names: {string.Join(", ", Catalog.Select(e => e.Name))}.",
                nameof(name));

        var fullPath = ResolvePath(entry.RelativePath);
        if (!File.Exists(fullPath))
        {
            throw new FileNotFoundException($"Demo deck file not found on disk: {fullPath}");
        }

        var clampedPpi = DeckPreviewValidators.ClampPpi(ppi);
        var bytes = File.ReadAllBytes(fullPath);

        // Whole-deck export: one PptxToTypst conversion + one Typst compile for all slides.
        byte[][] pages;
        int slideCount;
        var timer = Stopwatch.StartNew();
        using (var builder = PresentationBuilder.Open(bytes))
        {
            slideCount = builder.SlideCount;
            pages = builder.ExportThumbnails(new ThumbnailOptions { Ppi = clampedPpi, Format = format, FontDirectory = _fontDirectory });
        }
        timer.Stop();

        var fileName = Path.GetFileName(entry.RelativePath);
        var deckId = _sessionStore.Store(bytes, fileName, slideCount);

        return new DemoRenderResult(deckId, slideCount, timer.Elapsed.TotalMilliseconds, format, pages);
    }

    public DemoRenderResult RenderUploadedDeck(byte[] sourceBytes, string fileName, string format, int ppi)
    {
        ArgumentNullException.ThrowIfNull(sourceBytes);
        ArgumentException.ThrowIfNullOrWhiteSpace(fileName);

        // Defensive: the endpoint normalizes via DeckPreviewValidators.TryNormalizeFormat,
        // but the service requires already-normalized input (same contract as RenderDeck).
        if (format is not ("svg" or "png"))
        {
            throw new ArgumentException(
                $"Format must be normalized to \"svg\" or \"png\" (got '{format}').",
                nameof(format));
        }

        var clampedPpi = DeckPreviewValidators.ClampPpi(ppi);

        // Validate before the timed render: any OpenXML/zip parse failure means the
        // client sent a bad deck — surface it as a client error, not a render failure.
        int slideCount;
        try
        {
            using var builder = PresentationBuilder.Open(sourceBytes);
            slideCount = builder.SlideCount;
        }
        catch (Exception ex)
        {
            throw new ArgumentException(
                $"The uploaded file is not a valid PPTX deck: {ex.Message}",
                nameof(sourceBytes));
        }

        if (slideCount < 1)
        {
            throw new ArgumentException("The deck contains no slides.", nameof(sourceBytes));
        }

        if (slideCount > MaxUploadSlides)
        {
            throw new ArgumentException(
                $"Decks with more than {MaxUploadSlides} slides are not supported in the demo render (got {slideCount}).",
                nameof(sourceBytes));
        }

        // Whole-deck export: one PptxToTypst conversion + one Typst compile for all slides.
        byte[][] pages;
        var timer = Stopwatch.StartNew();
        using (var builder = PresentationBuilder.Open(sourceBytes))
        {
            pages = builder.ExportThumbnails(new ThumbnailOptions { Ppi = clampedPpi, Format = format, FontDirectory = _fontDirectory });
        }
        timer.Stop();

        var deckId = _sessionStore.Store(sourceBytes, fileName, slideCount);

        return new DemoRenderResult(deckId, slideCount, timer.Elapsed.TotalMilliseconds, format, pages);
    }

    public TypstLegResult RenderTypstLeg(byte[] sourceBytes, string fileName, int ppi)
    {
        ArgumentNullException.ThrowIfNull(sourceBytes);
        ArgumentException.ThrowIfNullOrWhiteSpace(fileName);

        // Validate before the timed renders (same contract as RenderUploadedDeck): any
        // OpenXML/zip parse failure means the client sent a bad deck.
        int slideCount;
        try
        {
            using var builder = PresentationBuilder.Open(sourceBytes);
            slideCount = builder.SlideCount;
        }
        catch (Exception ex)
        {
            throw new ArgumentException(
                $"The uploaded file is not a valid PPTX deck: {ex.Message}",
                nameof(sourceBytes));
        }

        if (slideCount < 1)
        {
            throw new ArgumentException("The deck contains no slides.", nameof(sourceBytes));
        }

        if (slideCount > MaxUploadSlides)
        {
            throw new ArgumentException(
                $"Decks with more than {MaxUploadSlides} slides are not supported in the demo render (got {slideCount}).",
                nameof(sourceBytes));
        }

        // PNG and PDF legs run CONCURRENTLY: they are independent whole-deck renders of
        // the same source bytes — each opens its own builder (own PptxToTypstConverter with
        // a GUID-suffixed temp directory) and its own Typst compile (TypstBridge compiles
        // run in parallel per its docs; the font cache is process-wide by design). The
        // per-phase stopwatches stay honest per-leg numbers; the outer stopwatch measures
        // the wall-clock time of the pair.
        var totalTimer = Stopwatch.StartNew();

        // PNG leg: one PptxToTypst conversion + one Typst compile for all slides. Render
        // failures propagate — the endpoint surfaces them as a 500.
        var pngTask = Task.Run(() =>
        {
            var pngTimer = Stopwatch.StartNew();
            byte[][] pngPages;
            using (var builder = PresentationBuilder.Open(sourceBytes))
            {
                pngPages = builder.ExportThumbnails(new ThumbnailOptions { Ppi = ppi, Format = "png", FontDirectory = _fontDirectory });
            }
            pngTimer.Stop();
            return (Pages: pngPages, Milliseconds: pngTimer.Elapsed.TotalMilliseconds);
        });

        // PDF leg: best-effort on a freshly opened builder (never shared with the PNG
        // render). A failure here must not kill the leg — it rides on the result. The task
        // is wrapped so it NEVER faults: a faulted PDF task would turn Task.WhenAll into a
        // throw even when the PNG leg succeeded.
        var pdfTask = Task.Run(() =>
        {
            try
            {
                var pdfTimer = Stopwatch.StartNew();
                byte[] pdfResult;
                using (var builder = PresentationBuilder.Open(sourceBytes))
                {
                    pdfResult = builder.ExportToPdf();
                }
                pdfTimer.Stop();
                return (Bytes: (byte[]?)pdfResult, Milliseconds: (double?)pdfTimer.Elapsed.TotalMilliseconds, Error: (string?)null);
            }
            catch (Exception ex)
            {
                return (Bytes: null, Milliseconds: null, Error: ex.Message);
            }
        });

        // WhenAll only completes once BOTH tasks finished (even when one faulted), so a PNG
        // failure still lets the best-effort PDF leg run to completion before propagating.
        Task.WhenAll(pngTask, pdfTask).GetAwaiter().GetResult();
        totalTimer.Stop();

        var png = pngTask.Result; // completed successfully if we got past the WhenAll
        var pdf = pdfTask.Result; // never throws (wrapped above)

        return new TypstLegResult(
            slideCount,
            png.Milliseconds,
            pdf.Milliseconds,
            totalTimer.Elapsed.TotalMilliseconds,
            png.Pages,
            pdf.Bytes,
            pdf.Error);
    }

    /// <summary>
    /// Resolves a whitelisted deck name to its absolute on-disk path for callers that
    /// need the raw file (e.g. the LibreOffice compare leg). Returns false for unknown
    /// names or missing files; never throws.
    /// </summary>
    internal bool TryGetDeckFile(string name, out string absolutePath)
    {
        absolutePath = string.Empty;
        if (string.IsNullOrWhiteSpace(name))
        {
            return false;
        }

        var entry = Catalog.FirstOrDefault(e => string.Equals(e.Name, name, StringComparison.OrdinalIgnoreCase));
        if (entry is null)
        {
            return false;
        }

        var fullPath = ResolvePath(entry.RelativePath);
        if (!File.Exists(fullPath))
        {
            return false;
        }

        absolutePath = fullPath;
        return true;
    }

    public string GetDeckTemplateJson()
    {
        var fullPath = ResolvePath(TemplateRelativePath);
        if (!File.Exists(fullPath))
        {
            throw new FileNotFoundException($"Demo deck template not found: {fullPath}");
        }

        return RewriteRelativeSrcPaths(File.ReadAllText(fullPath));
    }

    /// <summary>
    /// Normalizes every "src" string property holding a relative path to a repo-root-relative
    /// forward-slash path: any leading "./" is stripped and backslashes become forward
    /// slashes. Absolute paths, data URIs, URLs and non-string values are left untouched.
    /// The rewrite is pure (no filesystem access). Internal static so it is unit-testable
    /// without a web host.
    /// </summary>
    internal static string RewriteRelativeSrcPaths(string json)
    {
        ArgumentNullException.ThrowIfNull(json);

        var root = JsonNode.Parse(json)
            ?? throw new JsonException("The demo deck template parsed to a null JSON node.");

        RewriteNode(root);

        return root.ToJsonString(new JsonSerializerOptions { WriteIndented = true });
    }

    private static void RewriteNode(JsonNode node)
    {
        switch (node)
        {
            case JsonObject obj:
                foreach (var property in obj.ToList())
                {
                    if (property.Value is JsonValue value
                        && string.Equals(property.Key, "src", StringComparison.Ordinal)
                        && value.TryGetValue<string>(out var src)
                        && !string.IsNullOrWhiteSpace(src)
                        && !Path.IsPathRooted(src))
                    {
                        obj[property.Key] = NormalizeRelativePath(src);
                    }
                    else if (property.Value is JsonNode child)
                    {
                        RewriteNode(child);
                    }
                }
                break;
            case JsonArray array:
                foreach (var item in array)
                {
                    if (item is not null)
                    {
                        RewriteNode(item);
                    }
                }
                break;
        }
    }

    private static string NormalizeRelativePath(string path)
    {
        var normalized = path.Replace('\\', '/');
        while (normalized.StartsWith("./", StringComparison.Ordinal))
        {
            normalized = normalized[2..];
        }

        return normalized;
    }

    private static string ResolvePath(string relativePath)
    {
        var repoRoot = RepositoryRootLocator.FindOrFallback();
        return Path.GetFullPath(Path.Combine(repoRoot, relativePath.Replace('/', Path.DirectorySeparatorChar)));
    }

    /// <summary>
    /// One whitelisted catalog entry. The slide count is computed once (lazily, thread-safe)
    /// by opening the deck; a missing file yields 0 instead of throwing on listing.
    /// </summary>
    private sealed class DemoDeckEntry
    {
        private readonly Lazy<int> _slideCount;

        public DemoDeckEntry(string name, string relativePath, string description)
        {
            Name = name;
            RelativePath = relativePath;
            Description = description;
            _slideCount = new Lazy<int>(ComputeSlideCount);
        }

        public string Name { get; }
        public string RelativePath { get; }
        public string Description { get; }

        public int GetSlideCount() => _slideCount.Value;

        private int ComputeSlideCount()
        {
            var fullPath = ResolvePath(RelativePath);
            if (!File.Exists(fullPath))
            {
                return 0;
            }

            var bytes = File.ReadAllBytes(fullPath);
            using var builder = PresentationBuilder.Open(bytes);
            return builder.SlideCount;
        }
    }
}
