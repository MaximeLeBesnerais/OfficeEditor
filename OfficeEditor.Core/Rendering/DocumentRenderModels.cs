namespace OfficeEditor.Core.Rendering;

/// <summary>
/// Supported render output formats. <see cref="Typ"/> exports the document's Typst source
/// instead of compiling it (no Typst backend required); every other format compiles via the
/// repository's own Typst pipeline (<c>TypstCompilerService</c> — TypstBridge primary, CLI fallback).
/// </summary>
public enum DocumentOutputFormat
{
    Pdf,
    Png,
    Svg,
    Typ
}

/// <summary>
/// A render request against a document source file. <see cref="SourcePath"/> may point to a
/// .pptx, .docx, .xlsx or .json file — the facade detects the format by extension and routes
/// .json sources through the appropriate JSON generator first (see <see cref="DocumentRenderer"/>).
/// </summary>
public sealed record DocumentRenderRequest
{
    /// <summary>Absolute or relative path to the source document (.pptx/.docx/.xlsx/.json).</summary>
    public required string SourcePath { get; init; }

    /// <summary>Desired output format. Defaults to <see cref="DocumentOutputFormat.Pdf"/>.</summary>
    public DocumentOutputFormat Format { get; init; } = DocumentOutputFormat.Pdf;

    /// <summary>Raster density for PNG output (and SVG page geometry), in pixels per inch. Values ≤ 0 use the 150 default.</summary>
    public float Ppi { get; init; } = 150;

    /// <summary>
    /// Optional additional font directory (or path-separator-joined list) handed to the Typst
    /// compiler. For PPTX this extends — never replaces — the deck's embedded fonts.
    /// </summary>
    public string? FontPath { get; init; }
}

/// <summary>
/// The outcome of a render. <see cref="Pages"/> layout: <see cref="DocumentOutputFormat.Pdf"/> is a single
/// buffer; <see cref="DocumentOutputFormat.Png"/> and <see cref="DocumentOutputFormat.Svg"/> are one buffer per page;
/// <see cref="DocumentOutputFormat.Typ"/> is a single UTF-8 buffer of Typst source.
/// Failures are returned (never thrown) as <see cref="Success"/>=false with <see cref="ErrorMessage"/>.
/// </summary>
public sealed record DocumentRenderResult
{
    public bool Success { get; init; }

    public byte[][] Pages { get; init; } = Array.Empty<byte[]>();

    public string? ErrorMessage { get; init; }

    public IReadOnlyList<string> Warnings { get; init; } = Array.Empty<string>();
}

/// <summary>
/// Implemented exactly once per document format by the format core libraries
/// (PptxEditor.Core / DocxEditor.Core / XlsxEditor.Core). Internal: only the shared facade
/// (<see cref="DocumentRenderer"/>) discovers and drives implementations, so downstream
/// CLI/API/MCP agents code against <see cref="IDocumentRenderer"/> alone.
/// </summary>
internal interface IFormatRenderer
{
    DocumentRenderResult Render(DocumentRenderRequest request);

    /// <summary>
    /// True when this renderer owns the given source path: an exact extension match, or — for
    /// .json sources — the file matches this format's generation JSON vocabulary.
    /// </summary>
    bool CanRenderSource(string sourcePath);
}

/// <summary>
/// Declares the source extension a format renderer owns and the priority used when probing
/// .json sources: lower <see cref="JsonPriority"/> values are tried first (XLSX → DOCX → PPTX).
/// </summary>
[AttributeUsage(AttributeTargets.Class, Inherited = false, AllowMultiple = false)]
internal sealed class DocumentSourceFormatAttribute : Attribute
{
    public DocumentSourceFormatAttribute(string sourceExtension, int jsonPriority = 100)
    {
        SourceExtension = sourceExtension;
        JsonPriority = jsonPriority;
    }

    /// <summary>Native source extension without the leading dot (e.g. "pptx").</summary>
    public string SourceExtension { get; }

    /// <summary>JSON-vocabulary probe order; lower is tried first.</summary>
    public int JsonPriority { get; }
}
