using DocxEditor.Core.Generation.Assets;
using DocxEditor.Core.Generation.Contracts;
using DocxEditor.Core.Generation.Emit.Ooxml;
using DocxEditor.Core.Generation.Model;
using DocxEditor.Core.Generation.Schema;

namespace DocxEditor.Core.Generation;

/// <summary>Options for the end-to-end declarative DOCX generator.</summary>
public sealed record DocxGeneratorOptions
{
    /// <summary>
    /// Optional template path override. When omitted, the model's <c>template</c> value is used.
    /// The template is copied into memory and is never modified in place.
    /// </summary>
    public string? TemplatePath { get; init; }

    /// <summary>Local image path policy. Network image fetching is never enabled.</summary>
    public ImageSourceOptions ImageSourceOptions { get; init; } = ImageSourceOptions.Default;

    /// <summary>Image payload, pixel, and media-type limits.</summary>
    public ImageAssetOptions ImageAssetOptions { get; init; } = ImageAssetOptions.Default;
}

/// <summary>A generated DOCX byte buffer together with its warnings and resolved model.</summary>
public sealed record GeneratedDocx
{
    /// <summary>The complete DOCX package bytes.</summary>
    public required byte[] Content { get; init; }

    /// <summary>Generation metadata and canonical warnings.</summary>
    public required DocxGenerationResult Result { get; init; }
}

/// <summary>
/// High-level declarative generator. It validates JSON, resolves design/styles/assets, emits
/// flow and positioned content into one OOXML package, and exposes file, stream, and byte outputs.
/// A generator instance is stateless and may be reused; each call creates an isolated package.
/// </summary>
public sealed class DocxGenerator
{
    private readonly IDocxGenerationParser _parser;
    private readonly ImageAssetLoader _imageAssetLoader;

    /// <summary>Creates a generator using the built-in parser and image loader.</summary>
    public DocxGenerator()
        : this(new DocxGenerationDocumentParser(), new ImageAssetLoader())
    {
    }

    /// <summary>Creates a generator with replaceable parser and image-loader dependencies.</summary>
    public DocxGenerator(IDocxGenerationParser parser, ImageAssetLoader imageAssetLoader)
    {
        ArgumentNullException.ThrowIfNull(parser);
        ArgumentNullException.ThrowIfNull(imageAssetLoader);
        _parser = parser;
        _imageAssetLoader = imageAssetLoader;
    }

    /// <summary>Validates declarative JSON and atomically writes a DOCX file.</summary>
    public DocxGenerationResult Generate(string json, string outputPath, DocxGeneratorOptions? options = null)
    {
        var validation = ValidateJson(json);
        return GenerateToFile(validation.Document!, outputPath, options, validation.Warnings);
    }

    /// <summary>Validates a parsed model and atomically writes a DOCX file.</summary>
    public DocxGenerationResult Generate(
        DocxGenerationDocument document,
        string outputPath,
        DocxGeneratorOptions? options = null)
    {
        var validation = ValidateModel(document);
        return GenerateToFile(validation.Document!, outputPath, options, validation.Warnings);
    }

    /// <summary>
    /// Validates declarative JSON, writes the complete package to <paramref name="output"/>,
    /// and leaves the caller-owned stream open at the end of the written bytes. The stream
    /// is flushed before returning, so the package is immediately readable (e.g. after the
    /// caller seeks back); its position has advanced by exactly the package length.
    /// </summary>
    public DocxGenerationResult Generate(string json, Stream output, DocxGeneratorOptions? options = null)
    {
        var validation = ValidateJson(json);
        return GenerateToStream(validation.Document!, output, options, validation.Warnings);
    }

    /// <summary>
    /// Validates a parsed model, writes the complete package to <paramref name="output"/>,
    /// and leaves the caller-owned stream open at the end of the written bytes. The stream
    /// is flushed before returning, so the package is immediately readable (e.g. after the
    /// caller seeks back); its position has advanced by exactly the package length.
    /// </summary>
    public DocxGenerationResult Generate(
        DocxGenerationDocument document,
        Stream output,
        DocxGeneratorOptions? options = null)
    {
        var validation = ValidateModel(document);
        return GenerateToStream(validation.Document!, output, options, validation.Warnings);
    }

    /// <summary>Validates declarative JSON and returns the complete DOCX package in memory.</summary>
    public GeneratedDocx GenerateToBytes(string json, DocxGeneratorOptions? options = null)
    {
        var validation = ValidateJson(json);
        return GenerateBytes(validation.Document!, options, validation.Warnings);
    }

    /// <summary>Validates a parsed model and returns the complete DOCX package in memory.</summary>
    public GeneratedDocx GenerateToBytes(
        DocxGenerationDocument document,
        DocxGeneratorOptions? options = null)
    {
        var validation = ValidateModel(document);
        return GenerateBytes(validation.Document!, options, validation.Warnings);
    }

    private DocxGenerationValidationResult ValidateJson(string json)
    {
        ArgumentNullException.ThrowIfNull(json);
        return _parser.Validate(json).ThrowIfInvalid();
    }

    private static DocxGenerationValidationResult ValidateModel(DocxGenerationDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);
        return DocxGenerationModelValidator.Validate(document).ThrowIfInvalid();
    }

    private DocxGenerationResult GenerateToFile(
        DocxGenerationDocument document,
        string outputPath,
        DocxGeneratorOptions? options,
        IReadOnlyList<DocxGenerationIssue> parserWarnings)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        var effectiveDocument = ApplyOptions(document, options);
        using var emitter = CreateEmitter(outputPath, writeOutput: true, options);
        return MergeWarnings(emitter.Emit(effectiveDocument), parserWarnings);
    }

    private DocxGenerationResult GenerateToStream(
        DocxGenerationDocument document,
        Stream output,
        DocxGeneratorOptions? options,
        IReadOnlyList<DocxGenerationIssue> parserWarnings)
    {
        ArgumentNullException.ThrowIfNull(output);
        if (!output.CanWrite)
        {
            throw new ArgumentException("The output stream must be writable.", nameof(output));
        }

        var generated = GenerateBytes(document, options, parserWarnings);
        output.Write(generated.Content);
        output.Flush();
        return generated.Result;
    }

    private GeneratedDocx GenerateBytes(
        DocxGenerationDocument document,
        DocxGeneratorOptions? options,
        IReadOnlyList<DocxGenerationIssue> parserWarnings)
    {
        var effectiveDocument = ApplyOptions(document, options);
        using var emitter = CreateEmitter(outputPath: null, writeOutput: false, options);
        var result = MergeWarnings(emitter.Emit(effectiveDocument), parserWarnings);
        return new GeneratedDocx { Content = emitter.SaveToBytes(), Result = result };
    }

    private DocxOoxmlEmitter CreateEmitter(string? outputPath, bool writeOutput, DocxGeneratorOptions? options)
    {
        var effectiveOptions = options ?? new DocxGeneratorOptions();
        return new DocxOoxmlEmitter(
            new DocxEmitOptions
            {
                OutputPath = outputPath,
                WriteOutput = writeOutput,
                ImageSourceOptions = effectiveOptions.ImageSourceOptions,
                ImageAssetOptions = effectiveOptions.ImageAssetOptions
            },
            _imageAssetLoader);
    }

    private static DocxGenerationDocument ApplyOptions(
        DocxGenerationDocument document,
        DocxGeneratorOptions? options)
    {
        if (options?.TemplatePath is null)
        {
            return document;
        }
        ArgumentException.ThrowIfNullOrWhiteSpace(options.TemplatePath);
        return document with { TemplatePath = options.TemplatePath };
    }

    private static DocxGenerationResult MergeWarnings(
        DocxGenerationResult result,
        IReadOnlyList<DocxGenerationIssue> parserWarnings) =>
        parserWarnings.Count == 0
            ? result
            : result with { Warnings = [.. parserWarnings, .. result.Warnings] };

}
