namespace OfficeEditor.Core.Services;

/// <summary>
/// Identifies an Office document format supported by the binary stream APIs.
/// </summary>
public enum OfficeDocumentFormat
{
    Docx,
    Pptx,
    Xlsx,
    Markdown
}

/// <summary>
/// Lightweight metadata describing an in-memory Office document.
/// </summary>
public sealed record BinaryOfficeDocument
{
    public required byte[] Bytes { get; init; }

    public required OfficeDocumentFormat Format { get; init; }

    public required string ContentType { get; init; }

    public required string FileExtension { get; init; }
}
