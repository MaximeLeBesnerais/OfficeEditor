using OfficeEditor.Core.Services;
using PptxEditor.Core.Generation;
using PptxEditor.Core.Generation.Schema;

namespace OfficeEditor.Api.Services;

/// <summary>One rendered per-slide preview of a generated deck (1-based slide number).</summary>
public sealed record GeneratedSlidePreview(int Slide, string Format, string ContentType, byte[] Bytes);

/// <summary>
/// Outcome of a deck generation run (JSON in → PPTX + per-slide previews out).
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

    /// <summary>Warm-path metric (target &lt; 500ms): parse → expand → layout → PPTX, excluding preview render.</summary>
    public double GenerationMilliseconds { get; init; }

    /// <summary>End-to-end wall time including the preview render attempt.</summary>
    public double TotalMilliseconds { get; init; }
}

public interface IDeckGenerationService
{
    /// <summary>
    /// Runs the full generation pipeline: validate → archetype/component expand → measured layout →
    /// OOXML PPTX + per-slide Typst previews (SVG live-preview path through TypstBridge, the
    /// primary backend per the TypstBridge-primary backend convention). Preview rendering is best-effort.
    /// <paramref name="normalizedFormat"/> must already be normalized to "svg" or "png"
    /// (<see cref="DeckPreviewValidators.TryNormalizeFormat"/>).
    /// </summary>
    DeckGenerationResult Generate(string documentJson, string normalizedFormat, int ppi);
}

public sealed class DeckGenerationService : IDeckGenerationService
{
    private static readonly Lazy<string?> RepositoryRoot = new(RepositoryRootLocator.FindOrNull);
    private readonly PptxGenerator _generator;
    internal string? FontDirectory { get; }

    public DeckGenerationService(string? fontDirectory = null)
    {
        _generator = new PptxGenerator();
        FontDirectory = string.IsNullOrWhiteSpace(fontDirectory) ? null : fontDirectory;
    }

    internal DeckGenerationService(TypstCompilerService compiler, string? fontDirectory = null)
    {
        _generator = new PptxGenerator(compiler);
        FontDirectory = string.IsNullOrWhiteSpace(fontDirectory) ? null : fontDirectory;
    }

    public DeckGenerationResult Generate(string documentJson, string normalizedFormat, int ppi)
    {
        if (normalizedFormat is not ("svg" or "png"))
            throw new ArgumentException("Format must be normalized to 'svg' or 'png'.", nameof(normalizedFormat));

        var result = _generator.Generate(documentJson, new PptxGeneratorOptions
        {
            PreviewFormat = normalizedFormat,
            Ppi = ppi,
            FontDirectory = FontDirectory,
            AllowedImageRoot = RepositoryRoot.Value
        });
        return new DeckGenerationResult
        {
            Success = result.Success,
            PptxBytes = result.PptxBytes,
            SlideCount = result.SlideCount,
            Previews = result.Previews.Select(p => new GeneratedSlidePreview(p.Slide, p.Format, p.ContentType, p.Bytes)).ToArray(),
            Errors = result.Errors,
            Warnings = result.Warnings,
            PipelineWarnings = result.PipelineWarnings,
            PreviewError = result.PreviewError,
            GenerationMilliseconds = result.GenerationMilliseconds,
            TotalMilliseconds = result.TotalMilliseconds
        };
    }
}
