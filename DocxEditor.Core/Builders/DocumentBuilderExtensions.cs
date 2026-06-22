using OfficeEditor.Core.Services;

namespace DocxEditor.Core.Builders;

/// <summary>
/// Extension methods for <see cref="IDocumentBuilder"/> to produce a <see cref="BinaryOfficeDocument"/>.
/// </summary>
public static class DocumentBuilderExtensions
{
    /// <summary>
    /// Saves the document built by <paramref name="builder"/> to a byte array and returns it as a
    /// <see cref="BinaryOfficeDocument"/> with DOCX metadata.
    /// </summary>
    /// <param name="builder">The document builder. Must not be <see langword="null"/>.</param>
    /// <returns>A binary office document containing the DOCX bytes and format metadata.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="builder"/> is <see langword="null"/>.</exception>
    public static BinaryOfficeDocument ToBinaryDocument(this IDocumentBuilder builder)
    {
        if (builder is null)
        {
            throw new ArgumentNullException(nameof(builder));
        }

        return new BinaryOfficeDocument
        {
            Bytes = builder.SaveToBytes(),
            Format = OfficeDocumentFormat.Docx,
            ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            FileExtension = ".docx"
        };
    }
}
