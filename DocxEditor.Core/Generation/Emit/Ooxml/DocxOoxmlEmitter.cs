using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Generation.Contracts;
using DocxEditor.Core.Generation.Emit.Ooxml.Flow;
using DocxEditor.Core.Generation.Model;
using DocxEditor.Core.Generation.Schema;
using OfficeEditor.Core.Exceptions;

namespace DocxEditor.Core.Generation.Emit.Ooxml;

/// <summary>
/// Emit options for the flow-first OOXML emitter. <see cref="OutputPath"/> is the final
/// destination of the generated package; when omitted the emitter writes to a temporary
/// file (reported through <see cref="DocxGenerationResult.Outputs"/>) and still exposes the
/// package via <see cref="DocxOoxmlEmitter.SaveToBytes"/> / <see cref="DocxOoxmlEmitter.Save"/>.
/// </summary>
public sealed record DocxEmitOptions
{
    /// <summary>Destination for the generated .docx. Null = a temporary file.</summary>
    public string? OutputPath { get; init; }
}

/// <summary>
/// Image-content seam for the OOXML emitter. The image workstream implements this to load
/// and decode image sources; until a resolver is wired, inline images emit an explicit
/// placeholder paragraph plus a warning instead of being silently dropped.
/// </summary>
public interface IImageContentResolver
{
    /// <summary>
    /// Resolves an image source (path, URL or base64 payload) to raw bytes, a file extension
    /// used to pick the package image part type (e.g. "png", "jpg") and the natural size in
    /// points used for fit math when the model omits explicit dimensions.
    /// </summary>
    bool TryResolveImage(
        string source,
        out byte[] content,
        out string extension,
        out double naturalWidthPt,
        out double naturalHeightPt);
}

/// <summary>
/// Positioned-tier seam: emits one absolutely anchored primitive (text box, floating picture,
/// rect, line, callout) into the section's <see cref="Body"/>. The positioned workstream owns
/// the anchored-emission implementation; until then the flow branch uses a default that warns
/// and skips so generated documents never silently lose positioned content.
/// </summary>
public interface IPositionedTierEmitter
{
    /// <summary>
    /// Emits a single positioned element as anchored drawing content inside
    /// <paramref name="body"/>. Implementations add warnings for anything they cannot emit.
    /// </summary>
    void EmitPositionedElement(
        PositionedElement element,
        Body body,
        string path,
        IList<DocxGenerationIssue> warnings);
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
    private readonly IImageContentResolver? _imageResolver;
    private readonly IPositionedTierEmitter? _positionedTierEmitter;
    private readonly List<DocxGenerationIssue> _warnings = [];

    private MemoryStream? _packageStream;
    private WordprocessingDocument? _document;
    private MainDocumentPart? _mainPart;
    private bool _emitted;
    private bool _disposed;

    /// <param name="options">Destination options. Null = defaults (temporary output).</param>
    /// <param name="imageResolver">
    /// Optional image-content seam. Null = inline images become explicit placeholders with a
    /// warning (see <see cref="IImageContentResolver"/>).
    /// </param>
    /// <param name="positionedTierEmitter">
    /// Optional positioned-tier seam. Null = positioned elements warn and are skipped until
    /// the positioned workstream provides an implementation.
    /// </param>
    public DocxOoxmlEmitter(
        DocxEmitOptions? options = null,
        IImageContentResolver? imageResolver = null,
        IPositionedTierEmitter? positionedTierEmitter = null)
    {
        _options = options ?? new DocxEmitOptions();
        _imageResolver = imageResolver;
        _positionedTierEmitter = positionedTierEmitter;
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
            var context = new OoxmlEmitContext
            {
                Document = _document!,
                MainPart = _mainPart!,
                Design = document.Design,
                FromTemplate = document.TemplatePath is not null,
                ImageResolver = _imageResolver,
                PositionedTierEmitter = _positionedTierEmitter,
                Warnings = _warnings
            };

            EnsureStylesPart(context);
            EmitSections(context, document);
            EmitMetadata(document.Metadata);

            outputPath = _options.OutputPath ?? DefaultOutputPath();
            WriteOutput(outputPath);

            return new DocxGenerationResult
            {
                Document = document,
                Outputs = [new EmittedOutput(DocxOutputKind.Document, outputPath)],
                Warnings = [.. _warnings]
            };
        }
        catch
        {
            // A failed emit must never leak the open package handle or a partial output file;
            // the original failure is preserved and rethrown unchanged after best-effort teardown.
            TryDeleteFile(outputPath);
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
    public string? OutputPath => _options.OutputPath ?? (_emitted ? _outputPath : null);

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

    /// <summary>
    /// Ensures a styles part exists. For blank documents a minimal default is created so
    /// direct formatting and style references have a base; template styles are read into the
    /// cache untouched and never written back.
    /// </summary>
    private static void EnsureStylesPart(OoxmlEmitContext context)
    {
        if (context.FromTemplate)
        {
            context.LoadTemplateStyles();
            return;
        }

        var mainPart = context.MainPart;
        var stylesPart = mainPart.StyleDefinitionsPart ?? mainPart.AddNewPart<StyleDefinitionsPart>();
        if (stylesPart.Styles is not null)
        {
            return;
        }

        var normal = FormattingHelpers.BuildNormalDefaultStyle(context.Design);
        stylesPart.Styles = new Styles(normal);
        context.RegisterStyle(normal);
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
        var directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        using var file = File.Create(outputPath);
        Save(file);
        _outputPath = outputPath;
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
