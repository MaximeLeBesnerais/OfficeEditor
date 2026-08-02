using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Generation.Assets;
using DocxEditor.Core.Generation.Contracts;
using DocxEditor.Core.Generation.Design;
using DocxEditor.Core.Generation.Emit.Ooxml.Design;
using DocxEditor.Core.Generation.Emit.Ooxml.Flow;
using DocxEditor.Core.Generation.Emit.Ooxml.Positioned;
using DocxEditor.Core.Generation.Model;
using DocxEditor.Core.Generation.Schema;
using OfficeEditor.Core.Exceptions;

namespace DocxEditor.Core.Generation.Emit.Ooxml;

/// <summary>
/// Emit options for the flow-first OOXML emitter. <see cref="OutputPath"/> is the final
/// destination of the generated package; when omitted and <see cref="WriteOutput"/> is true,
/// the emitter writes to a temporary
/// file (reported through <see cref="DocxGenerationResult.Outputs"/>) and still exposes the
/// package via <see cref="DocxOoxmlEmitter.SaveToBytes"/> / <see cref="DocxOoxmlEmitter.Save"/>.
/// </summary>
public sealed record DocxEmitOptions
{
    /// <summary>Destination for the generated .docx. Null = a temporary file.</summary>
    public string? OutputPath { get; init; }

    /// <summary>Whether <see cref="DocxOoxmlEmitter.Emit"/> writes a file artifact.</summary>
    public bool WriteOutput { get; init; } = true;

    /// <summary>Local image source policy. HTTP(S) sources remain unsupported.</summary>
    public ImageSourceOptions ImageSourceOptions { get; init; } = ImageSourceOptions.Default;

    /// <summary>Image payload and dimension limits.</summary>
    public ImageAssetOptions ImageAssetOptions { get; init; } = ImageAssetOptions.Default;
}

/// <summary>
/// Flow-first OOXML emitter: builds a real DOCX package in memory (optionally reusing a
/// template's styles part) from a validated <see cref="DocxGenerationDocument"/>, emitting
/// sections in order with valid section properties, flow blocks (paragraphs, headings, lists,
/// tables, callouts, page breaks, groups), headers/footers per section and inline images.
/// Existing template style definitions and numbering are never mutated: styles are referenced
/// by ID only and lists allocate fresh, collision-free numbering instances.
///
/// The emitter is disposable; after <see cref="Emit"/> succeeds the package stays open in
/// memory so callers can also save to a stream or byte array. On failure the package and any
/// partial output file are cleaned up before the original exception propagates.
/// </summary>
public sealed class DocxOoxmlEmitter : IDocxDocumentEmitter, IDisposable
{
    private readonly DocxEmitOptions _options;
    private readonly ImageAssetLoader _imageAssetLoader;
    private readonly List<DocxGenerationIssue> _warnings = [];

    private MemoryStream? _packageStream;
    private WordprocessingDocument? _document;
    private MainDocumentPart? _mainPart;
    private bool _emitted;
    private bool _disposed;

    /// <param name="options">Destination options. Null = defaults (temporary output).</param>
    /// <param name="imageAssetLoader">Optional asset loader; null uses the built-in safe loader.</param>
    public DocxOoxmlEmitter(
        DocxEmitOptions? options = null,
        ImageAssetLoader? imageAssetLoader = null)
    {
        _options = options ?? new DocxEmitOptions();
        _imageAssetLoader = imageAssetLoader ?? new ImageAssetLoader();
    }

    /// <summary>
    /// Emits <paramref name="document"/> to the configured (or a temporary) .docx path and
    /// returns the generation result shell. Emit is single-shot: a second call throws.
    /// The input model is never mutated.
    /// </summary>
    public DocxGenerationResult Emit(DocxGenerationDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);
        ObjectDisposedException.ThrowIf(_disposed, this);
        if (_emitted)
        {
            throw new InvalidOperationException("The emitter has already emitted a document; create a new emitter instance per document.");
        }
        _emitted = true;

        string? outputPath = null;
        try
        {
            CreatePackage(document.TemplatePath);
            var designResolver = new DocxDesignResolver(document.Design);
            _ = designResolver.ResolveAll();
            var styleManager = new DocxStyleManager(_document!, designResolver);
            var images = new DocxImagePipeline(
                _imageAssetLoader,
                _options.ImageSourceOptions,
                _options.ImageAssetOptions,
                _warnings);
            var positionedEmitter = new PositionedElementEmitter(new PositionedElementEmitOptions
            {
                Design = document.Design,
                ImageResolver = images
            });
            var context = new OoxmlEmitContext
            {
                Document = _document!,
                MainPart = _mainPart!,
                DesignResolver = designResolver,
                StyleManager = styleManager,
                Images = images,
                PositionedEmitter = positionedEmitter,
                FromTemplate = document.TemplatePath is not null,
                Warnings = _warnings
            };

            EmitSections(context, document);
            EmitMetadata(document.Metadata);
            AddDesignWarnings(designResolver.Warnings);

            IReadOnlyList<EmittedOutput> outputs = [];
            if (_options.WriteOutput)
            {
                outputPath = _options.OutputPath ?? DefaultOutputPath();
                WriteOutput(outputPath);
                outputPath = _outputPath!;
                outputs = [new EmittedOutput(DocxOutputKind.Document, outputPath)];
            }

            return new DocxGenerationResult
            {
                Document = document,
                Outputs = outputs,
                Warnings = [.. _warnings]
            };
        }
        catch
        {
            // Output is staged beside the destination and atomically renamed, so a failed emit
            // never touches an existing destination or leaves a partial final file.
            Dispose();
            throw;
        }
    }

    /// <summary>
    /// Saves the emitted package to <paramref name="stream"/> and leaves it open. The document
    /// must be emitted (or the package auto-creates a blank document so this never fails).
    /// </summary>
    public void Save(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);
        ObjectDisposedException.ThrowIf(_disposed, this);

        EnsurePackageExists();
        _document!.Save();
        _packageStream!.Position = 0;
        _packageStream.CopyTo(stream);
        _packageStream.Position = 0;
    }

    /// <summary>Returns the emitted package as a byte array.</summary>
    public byte[] SaveToBytes()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        EnsurePackageExists();
        _document!.Save();
        return _packageStream!.ToArray();
    }

    /// <summary>The full path the last <see cref="Emit"/> wrote to (null before emit).</summary>
    public string? OutputPath => _outputPath;

    private string? _outputPath;

    /// <summary>Warnings collected by the last emit, in emit order.</summary>
    public IReadOnlyList<DocxGenerationIssue> Warnings => _warnings;

    private void EnsurePackageExists()
    {
        if (_document is null)
        {
            CreatePackage(null);
        }
    }

    /// <summary>
    /// Creates the in-memory package: a fresh blank document, or — when a template is
    /// requested — a memory copy of the template that is opened for editing so the template
    /// file itself is never touched and its existing style definitions are never mutated.
    /// </summary>
    private void CreatePackage(string? templatePath)
    {
        if (templatePath is null)
        {
            _packageStream = new MemoryStream();
            _document = WordprocessingDocument.Create(_packageStream, WordprocessingDocumentType.Document);
            _mainPart = _document.AddMainDocumentPart();
            _mainPart.Document = new Document();
            _mainPart.Document.Append(new Body());
            _mainPart.Document.Save();
            return;
        }

        byte[] bytes;
        try
        {
            bytes = File.ReadAllBytes(templatePath);
        }
        catch (FileNotFoundException ex)
        {
            throw new OfficeEditorException(
                $"The template file '{templatePath}' does not exist; generate from a blank document by omitting 'template'.", ex);
        }

        _packageStream = new MemoryStream(bytes);
        try
        {
            _document = WordprocessingDocument.Open(_packageStream, true);
        }
        catch (Exception ex) when (ex is OpenXmlPackageException or FileFormatException or InvalidDataException or XmlException)
        {
            _packageStream.Dispose();
            _packageStream = null;
            throw new OfficeEditorException(
                $"The template file '{templatePath}' is not a valid DOCX package.", ex);
        }

        _mainPart = _document.MainDocumentPart
            ?? throw new OfficeEditorException(
                "The template package has no WordprocessingML main document part, so it cannot be used as a template.");
    }

    private void EmitSections(OoxmlEmitContext context, DocxGenerationDocument document)
    {
        if (document.Sections.Count == 0)
        {
            context.Warnings.Add(new DocxGenerationIssue(
                "$.sections", "the document has no sections; an empty A4 document was emitted.", null,
                DocxGenerationIssueSeverity.Warning));
            SectionEmitter.EmitEmpty(context);
            return;
        }

        SectionEmitter.EmitAll(context, document.Sections);
    }

    private void EmitMetadata(DocxMetadata? metadata)
    {
        if (metadata is null)
        {
            return;
        }

        var properties = _document?.PackageProperties;
        if (properties is null)
        {
            return;
        }

        properties.Title = metadata.Title;
        properties.Creator = metadata.Author;
        properties.Subject = metadata.Subject;
        properties.Keywords = metadata.Keywords;
        properties.Description = metadata.Description;
        properties.Language = metadata.Language;
    }

    private void WriteOutput(string outputPath)
    {
        var fullOutputPath = Path.GetFullPath(outputPath);
        var directory = Path.GetDirectoryName(fullOutputPath);
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var tempPath = Path.Combine(
            directory ?? string.Empty,
            $".{Path.GetFileName(fullOutputPath)}.{Guid.NewGuid():N}.tmp");
        try
        {
            using (var file = File.Create(tempPath))
            {
                Save(file);
            }
            File.Move(tempPath, fullOutputPath, overwrite: true);
            _outputPath = fullOutputPath;
        }
        finally
        {
            TryDeleteFile(tempPath);
        }
    }

    private void AddDesignWarnings(IReadOnlyList<DesignResolutionWarning> warnings)
    {
        foreach (var warning in warnings)
        {
            var path = warning.Context switch
            {
                null or "" => "$.design",
                var context when context.StartsWith('$') => context,
                var context => $"$.{context}"
            };
            _warnings.Add(new DocxGenerationIssue(
                path,
                $"[{warning.Code}] {warning.Message}",
                null,
                DocxGenerationIssueSeverity.Warning));
        }
    }

    private static string DefaultOutputPath() =>
        Path.Combine(Path.GetTempPath(), $"officeeditor-{Guid.NewGuid():N}.docx");

    private static void TryDeleteFile(string? path)
    {
        if (path is null)
        {
            return;
        }

        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
            // Best-effort cleanup must never mask the primary failure.
        }
    }

    /// <summary>Disposes the in-memory package (and its stream) if present.</summary>
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }
        _disposed = true;

        try
        {
            _document?.Dispose();
        }
        finally
        {
            _document = null;
            _packageStream?.Dispose();
            _packageStream = null;
            _mainPart = null;
        }
    }
}
