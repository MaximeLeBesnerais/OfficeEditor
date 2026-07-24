using System.Diagnostics;
using OfficeEditor.Core.Services;
using PptxEditor.Core.Generation.Archetypes;
using PptxEditor.Core.Generation.Components;
using PptxEditor.Core.Generation.Emit.Ooxml;
using PptxEditor.Core.Generation.Emit.Typst;
using PptxEditor.Core.Generation.Layout;
using PptxEditor.Core.Generation.Schema;
using PptxEditor.Core.Services;

namespace OfficeEditor.Api.Services;

/// <summary>One rendered per-slide preview of a generated deck (1-based slide number).</summary>
public sealed record GeneratedSlidePreview(int Slide, string Format, string ContentType, byte[] Bytes);

/// <summary>
/// Outcome of a deck generation run (plan.md §7.1: JSON in → PPTX + per-slide previews out).
/// <see cref="Success"/> is false only when the document itself is rejected — a failed
/// preview render degrades to <see cref="PreviewError"/> with the PPTX still delivered.
/// </summary>
public sealed record DeckGenerationResult
{
    public required bool Success { get; init; }

    /// <summary>The generated .pptx package. Null when the document was rejected.</summary>
    public byte[]? PptxBytes { get; init; }

    public int SlideCount { get; init; }

    /// <summary>Per-slide previews, empty when the render backend was unavailable.</summary>
    public IReadOnlyList<GeneratedSlidePreview> Previews { get; init; } = [];

    /// <summary>Rejection findings, in document order: P1 validator errors verbatim, or a component/layout loud error.</summary>
    public IReadOnlyList<GenerationIssue> Errors { get; init; } = [];

    /// <summary>P1 validator warnings (e.g. off-palette hex), verbatim.</summary>
    public IReadOnlyList<GenerationIssue> Warnings { get; init; } = [];

    /// <summary>Layout/OOXML-emission diagnostics (shrink warnings, image fit fallbacks, …).</summary>
    public IReadOnlyList<string> PipelineWarnings { get; init; } = [];

    /// <summary>Set when the Typst preview render was unavailable or failed; the PPTX is still valid.</summary>
    public string? PreviewError { get; init; }

    /// <summary>Warm-path metric (plan.md §7.1 target &lt; 500ms): parse → expand → layout → PPTX, excluding preview render.</summary>
    public double GenerationMilliseconds { get; init; }

    /// <summary>End-to-end wall time including the preview render attempt.</summary>
    public double TotalMilliseconds { get; init; }
}

public interface IDeckGenerationService
{
    /// <summary>
    /// Runs the full generation pipeline: P1 validate → component expand → layout resolve →
    /// OOXML PPTX + per-slide Typst previews (SVG live-preview path through TypstBridge, the
    /// primary backend per AGENTS.typst.md). Preview rendering is best-effort.
    /// <paramref name="normalizedFormat"/> must already be normalized to "svg" or "png"
    /// (<see cref="DeckPreviewValidators.TryNormalizeFormat"/>).
    /// </summary>
    DeckGenerationResult Generate(string documentJson, string normalizedFormat, int ppi);
}

public sealed class DeckGenerationService : IDeckGenerationService
{
    /// <summary>
    /// Repository root (located once per process by walking up from the app base directory
    /// for a ".git" entry — directory in a normal checkout, file in a git worktree; see
    /// <see cref="RepositoryRootLocator"/>). Used as the Typst project root so
    /// repo-root-relative asset paths in generation documents resolve inside the
    /// repository sandbox — and to absolutize + confine the same paths for the OOXML
    /// pass (repo-root-first precedence, CWD fallback; see
    /// <see cref="ImageSourceResolver"/>). Null when no ".git" entry is found:
    /// the compile then keeps the compiler default (process CWD), the pre-fix behavior.
    /// </summary>
    private static readonly Lazy<string?> RepositoryRoot = new(RepositoryRootLocator.FindOrNull);

    private readonly TypstCompilerService _compiler;
    private readonly string? _fontDirectory;

    public DeckGenerationService(string? fontDirectory = null)
        : this(new TypstCompilerService(), fontDirectory)
    {
    }

    internal DeckGenerationService(TypstCompilerService compiler, string? fontDirectory = null)
    {
        _compiler = compiler;
        _fontDirectory = string.IsNullOrWhiteSpace(fontDirectory) ? null : fontDirectory;
    }

    /// <summary>
    /// Optional font directory (Path.PathSeparator-separated list) handed to the Typst
    /// preview compile. Null/empty keeps the compiler's default font resolution.
    /// </summary>
    internal string? FontDirectory => _fontDirectory;

    public DeckGenerationResult Generate(string documentJson, string normalizedFormat, int ppi)
    {
        ArgumentNullException.ThrowIfNull(documentJson);
        if (normalizedFormat is not ("svg" or "png"))
        {
            throw new ArgumentException($"Format must be normalized to \"svg\" or \"png\" (got '{normalizedFormat}').", nameof(normalizedFormat));
        }

        var totalTimer = Stopwatch.StartNew();
        var generationTimer = Stopwatch.StartNew();

        var validation = new GenerationDocumentParser().Validate(documentJson);
        if (!validation.IsValid)
        {
            return Rejected(validation.Errors, validation.Warnings, totalTimer);
        }

        LayoutResult layout;
        try
        {
            // The whole stack is rebuilt per call: resolver/measurer/catalog carry per-run
            // state and are not thread-safe; the system-font scan behind the catalog is
            // cached per process, so this stays cheap on the warm path.
            var archetyped = ArchetypeExpander.Expand(validation.Document!);
            var expanded = ComponentExpander.Expand(archetyped);
            layout = new LayoutResolver(new TextMeasure(new FontMetricsCatalog())).Resolve(expanded);
        }
        catch (ComponentException ex)
        {
            return Rejected([Issue(ex)], validation.Warnings, totalTimer);
        }
        catch (LayoutException ex)
        {
            return Rejected([Issue(ex)], validation.Warnings, totalTimer);
        }

        // The OOXML emitter resolves image file sources with repo-root-first precedence
        // (ImageSourceResolver), while the Typst preview resolves them against the Typst
        // project root (the repository root; Typst itself confines file reads to the
        // project root). File image sources are therefore absolutized against — and
        // confined to — the repository root for the OOXML pass only; the Typst pass
        // keeps the authored repo-relative paths.
        OoxmlEmissionResult emission;
        try
        {
            var ooxmlLayout = RepositoryRoot.Value is { } repoRoot
                ? AbsolutizeImageSources(layout, repoRoot)
                : layout;
            emission = new OoxmlEmitter().Emit(ooxmlLayout);
        }
        catch (ArgumentException ex)
        {
            // Path traversal ("../../etc/passwd", absolute paths outside the repository)
            // and unsupported image extensions are document errors: reject the document
            // with an actionable message instead of faulting the request.
            return Rejected(
                [new GenerationIssue("$", ex.Message, null, GenerationIssueSeverity.Error)],
                validation.Warnings, totalTimer);
        }
        var generationMs = generationTimer.Elapsed.TotalMilliseconds;

        var (previews, previewError) = RenderPreviews(layout, normalizedFormat, ppi);

        totalTimer.Stop();
        return new DeckGenerationResult
        {
            Success = true,
            PptxBytes = emission.Bytes,
            SlideCount = layout.Slides.Count,
            Previews = previews,
            Warnings = validation.Warnings,
            PipelineWarnings = [.. layout.Warnings, .. emission.Warnings],
            PreviewError = previewError,
            GenerationMilliseconds = generationMs,
            TotalMilliseconds = totalTimer.Elapsed.TotalMilliseconds
        };
    }

    private (IReadOnlyList<GeneratedSlidePreview> Previews, string? Error) RenderPreviews(
        LayoutResult layout, string normalizedFormat, int ppi)
    {
        string source;
        try
        {
            source = new TypstEmitter().Emit(layout);
        }
        catch (TypstEmitException ex)
        {
            return ([], $"Typst emission failed: {ex.Message}");
        }

        var result = _compiler.Compile(source, new CompileOptions
        {
            Format = normalizedFormat == "svg" ? OutputFormat.Svg : OutputFormat.Png,
            Ppi = ppi,
            FontDirectory = _fontDirectory,
            WorkingDirectory = RepositoryRoot.Value
        });

        if (!result.Success)
        {
            return ([], result.ErrorMessage ?? "Typst preview compilation failed.");
        }
        if (result.Pages.Length != layout.Slides.Count)
        {
            return ([], $"Typst preview produced {result.Pages.Length} page(s) for {layout.Slides.Count} slide(s).");
        }

        var contentType = DeckPreviewValidators.ContentTypeForFormat(normalizedFormat);
        var previews = new List<GeneratedSlidePreview>(result.Pages.Length);
        for (var i = 0; i < result.Pages.Length; i++)
        {
            previews.Add(new GeneratedSlidePreview(i + 1, normalizedFormat, contentType, result.Pages[i]));
        }
        return (previews, null);
    }

    /// <summary>
    /// Returns a copy of the layout with every file image source resolved against — and
    /// confined to — <paramref name="repoRoot"/> via
    /// <see cref="ImageSourceResolver.ResolveContained"/>: relative sources are
    /// absolutized under the root; any source that would escape the root ("../../…" or
    /// an absolute path outside the repository) throws <see cref="ArgumentException"/>.
    /// Data URIs and URLs pass through untouched. Records make this a pure structural map.
    /// </summary>
    private static LayoutResult AbsolutizeImageSources(LayoutResult layout, string repoRoot)
    {
        ResolvedElement Map(ResolvedElement element) => element switch
        {
            ResolvedImage image when IsFileSource(image.Source) =>
                image with { Source = ImageSourceResolver.ResolveContained(repoRoot, image.Source) },
            ResolvedContainer container =>
                container with { Children = [.. container.Children.Select(Map)] },
            ResolvedGroup group =>
                group with { Children = [.. group.Children.Select(Map)] },
            _ => element
        };

        return layout with
        {
            Slides = [.. layout.Slides.Select(slide => slide with { Root = (ResolvedContainer)Map(slide.Root) })]
        };
    }

    private static bool IsFileSource(string source) =>
        !string.IsNullOrWhiteSpace(source)
        && !source.StartsWith("data:", StringComparison.OrdinalIgnoreCase)
        && !source.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
        && !source.StartsWith("https://", StringComparison.OrdinalIgnoreCase);

    private static GenerationIssue Issue(ComponentException ex) =>
        new(ex.Path, StripPathPrefix(ex.Path, ex.Message), null, GenerationIssueSeverity.Error);

    private static GenerationIssue Issue(LayoutException ex) =>
        new(ex.Path, StripPathPrefix(ex.Path, ex.Message), null, GenerationIssueSeverity.Error);

    private static string StripPathPrefix(string path, string message) =>
        message.StartsWith(path + ": ", StringComparison.Ordinal)
            ? message[(path.Length + 2)..]
            : message;

    private static DeckGenerationResult Rejected(
        IReadOnlyList<GenerationIssue> errors,
        IReadOnlyList<GenerationIssue> warnings,
        Stopwatch totalTimer)
    {
        totalTimer.Stop();
        return new DeckGenerationResult
        {
            Success = false,
            Errors = errors,
            Warnings = warnings,
            TotalMilliseconds = totalTimer.Elapsed.TotalMilliseconds
        };
    }
}
