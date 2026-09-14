using System.Diagnostics;
using OfficeEditor.Core.Services;
using PptxEditor.Core.Generation.Archetypes;
using PptxEditor.Core.Generation.Components;
using PptxEditor.Core.Generation.Emit.Ooxml;
using PptxEditor.Core.Generation.Emit.Typst;
using PptxEditor.Core.Generation.Layout;
using PptxEditor.Core.Generation.Schema;
using PptxEditor.Core.Models;
using PptxEditor.Core.Services;

namespace PptxEditor.Core.Generation;

/// <summary>One rendered per-slide preview of a generated deck (1-based slide number).</summary>
public sealed record PptxSlidePreview(int Slide, string Format, string ContentType, byte[] Bytes);

/// <summary>
/// Outcome of a deck generation run (JSON in → PPTX + per-slide previews out).
/// <see cref="Success"/> is false only when the document itself is rejected — a failed
/// preview render degrades to <see cref="PreviewError"/> with the PPTX still delivered.
/// </summary>
public sealed record PptxGenerationResult
{
    public required bool Success { get; init; }

    /// <summary>Validated source with IDs added. Persist this to keep generated identities across edits.</summary>
    public string? NormalizedJson { get; init; }

    /// <summary>The generated .pptx package. Null when the document was rejected.</summary>
    public byte[]? PptxBytes { get; init; }

    public int SlideCount => Layout?.Slides.Count ?? 0;

    /// <summary>The single resolved layout used for delivery and previews; null on rejection.</summary>
    public LayoutResult? Layout { get; init; }

    /// <summary>Per-slide previews, empty when the render backend was unavailable.</summary>
    public IReadOnlyList<PptxSlidePreview> Previews { get; init; } = [];

    /// <summary>Rejection findings, in document order: P1 validator errors verbatim, or a component/layout loud error.</summary>
    public IReadOnlyList<GenerationIssue> Errors { get; init; } = [];

    /// <summary>P1 validator warnings (e.g. off-palette hex), verbatim.</summary>
    public IReadOnlyList<GenerationIssue> Warnings { get; init; } = [];

    /// <summary>Layout/OOXML-emission diagnostics (shrink warnings, image fit fallbacks, …).</summary>
    public IReadOnlyList<string> PipelineWarnings { get; init; } = [];

    /// <summary>Set when the Typst preview render was unavailable or failed; the PPTX is still valid.</summary>
    public string? PreviewError { get; init; }

    /// <summary>Warm-path metric (target &lt; 500ms): parse → expand → layout → PPTX, excluding preview render.</summary>
    public double GenerationMilliseconds { get; init; }

    /// <summary>End-to-end wall time including the preview render attempt.</summary>
    public double TotalMilliseconds { get; init; }
}

/// <summary>Host policy and optional previews for the shared PPTX generation pipeline.</summary>
public sealed record PptxGeneratorOptions
{
    /// <summary>Directory of the source JSON. Relative images are confined to this directory.</summary>
    public string? DocumentDirectory { get; init; }

    /// <summary>
    /// Optional host boundary for all file images, including absolute paths. When supplied,
    /// relative images resolve under this root; the same root is used for preview compilation.
    /// </summary>
    public string? AllowedImageRoot { get; init; }

    /// <summary>Null produces only PPTX; "svg" or "png" also requests best-effort previews.</summary>
    public string? PreviewFormat { get; init; }

    /// <summary>Preview density, between 36 and 600 pixels per inch.</summary>
    public float Ppi { get; init; } = 150;

    /// <summary>Additional font directories, separated by Path.PathSeparator, for layout and preview.</summary>
    public string? FontDirectory { get; init; }
}

/// <summary>
/// Shared JSON → validation → archetypes → components → measured layout → PPTX pipeline.
/// Each call owns its parser, layout resolver and emitters. Preview failure leaves the PPTX
/// available. File IO failures propagate; rejected documents carry actionable diagnostics.
/// </summary>
public sealed class PptxGenerator
{
    private readonly TypstCompilerService? _compiler;

    /// <summary>Creates a generator with a per-call compiler when previews are requested.</summary>
    public PptxGenerator() { }

    /// <summary>Uses a caller-owned compiler for previews; the generator never disposes it.</summary>
    public PptxGenerator(TypstCompilerService compiler)
    {
        ArgumentNullException.ThrowIfNull(compiler);
        _compiler = compiler;
    }

    /// <summary>Generates PPTX bytes and a resolved layout, with optional per-slide previews.</summary>
    public PptxGenerationResult Generate(string documentJson, PptxGeneratorOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(documentJson);
        options ??= new PptxGeneratorOptions();
        if (options.PreviewFormat is not (null or "svg" or "png"))
            throw new ArgumentException("Preview format must be 'svg' or 'png'.", nameof(options));
        if (!float.IsFinite(options.Ppi) || options.Ppi is < 36 or > 600)
            throw new ArgumentOutOfRangeException(nameof(options), "Preview PPI must be between 36 and 600.");
        if (options.DocumentDirectory is not null && options.AllowedImageRoot is not null)
            throw new ArgumentException("Specify DocumentDirectory or AllowedImageRoot, not both.", nameof(options));

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
            layout = new LayoutResolver(new TextMeasure(new FontMetricsCatalog(options: new FontMetricsCatalogOptions
            {
                AdditionalFontDirectories = options.FontDirectory?.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries) ?? []
            }))).Resolve(expanded);
        }
        catch (ComponentException ex)
        {
            return Rejected([Issue(ex)], validation.Warnings, totalTimer);
        }
        catch (LayoutException ex)
        {
            return Rejected([Issue(ex)], validation.Warnings, totalTimer);
        }

        OoxmlEmissionResult emission;
        try
        {
            var ooxmlLayout = options.AllowedImageRoot is { } allowedRoot
                ? AbsolutizeImageSources(layout, allowedRoot)
                : layout;
            emission = new OoxmlEmitter(new OoxmlEmitOptions { DocumentDirectory = options.DocumentDirectory }).Emit(ooxmlLayout);
        }
        catch (Exception ex) when (ex is ArgumentException or FormatException)
        {
            // Path traversal ("../../etc/passwd", absolute paths outside the allowed root)
            // and unsupported image extensions are document errors: reject the document
            // with an actionable message instead of faulting the request.
            return Rejected(
                [new GenerationIssue("$", ex.Message, null, GenerationIssueSeverity.Error)],
                validation.Warnings, totalTimer);
        }
        var generationMs = generationTimer.Elapsed.TotalMilliseconds;

        (IReadOnlyList<PptxSlidePreview> previews, string? previewError) = options.PreviewFormat is null
            ? ([], null)
            : RenderPreviews(layout, options);

        totalTimer.Stop();
        return new PptxGenerationResult
        {
            Success = true,
            PptxBytes = emission.Bytes,
            Layout = layout,
            Previews = previews,
            Warnings = validation.Warnings,
            PipelineWarnings = [.. layout.Warnings, .. emission.Warnings],
            PreviewError = previewError,
            NormalizedJson = validation.NormalizedJson,
            GenerationMilliseconds = generationMs,
            TotalMilliseconds = totalTimer.Elapsed.TotalMilliseconds
        };
    }

    private (IReadOnlyList<PptxSlidePreview> Previews, string? Error) RenderPreviews(
        LayoutResult layout, PptxGeneratorOptions options)
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

        using var ownedCompiler = _compiler is null ? new TypstCompilerService() : null;
        var result = (_compiler ?? ownedCompiler!).Compile(source, new CompileOptions
        {
            Format = options.PreviewFormat == "svg" ? OutputFormat.Svg : OutputFormat.Png,
            Ppi = options.Ppi,
            FontDirectory = options.FontDirectory,
            WorkingDirectory = options.AllowedImageRoot ?? options.DocumentDirectory
        });

        if (!result.Success)
        {
            return ([], result.ErrorMessage ?? "Typst preview compilation failed.");
        }
        if (result.Pages.Length != layout.Slides.Count)
        {
            return ([], $"Typst preview produced {result.Pages.Length} page(s) for {layout.Slides.Count} slide(s).");
        }

        var contentType = options.PreviewFormat == "svg" ? "image/svg+xml" : "image/png";
        var previews = new List<PptxSlidePreview>(result.Pages.Length);
        for (var i = 0; i < result.Pages.Length; i++)
        {
            previews.Add(new PptxSlidePreview(i + 1, options.PreviewFormat!, contentType, result.Pages[i]));
        }
        return (previews, null);
    }

    /// <summary>
    /// Returns a copy of the layout with every file image source resolved against — and
    /// confined to — <paramref name="allowedRoot"/> via
    /// <see cref="ImageSourceResolver.ResolveContainedCanonical"/>: relative sources are
    /// absolutized under the root; any source that would escape the root ("../../…" or
    /// an absolute path outside the allowed root) throws <see cref="ArgumentException"/>.
    /// Data URIs and URLs pass through untouched. Records make this a pure structural map.
    /// </summary>
    private static LayoutResult AbsolutizeImageSources(LayoutResult layout, string allowedRoot)
    {
        ResolvedElement Map(ResolvedElement element) => element switch
        {
            ResolvedImage image when IsFileSource(image.Source) =>
                image with { Source = ImageSourceResolver.ResolveContainedCanonical(allowedRoot, ImageSourceResolver.ResolveContained(allowedRoot, image.Source)) },
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

    private static PptxGenerationResult Rejected(
        IReadOnlyList<GenerationIssue> errors,
        IReadOnlyList<GenerationIssue> warnings,
        Stopwatch totalTimer)
    {
        totalTimer.Stop();
        return new PptxGenerationResult
        {
            Success = false,
            Errors = errors,
            Warnings = warnings,
            TotalMilliseconds = totalTimer.Elapsed.TotalMilliseconds
        };
    }
}
