using OfficeEditor.Core.Services;

namespace PptxEditor.Core.Builders;

/// <summary>
/// Extension methods for <see cref="IPresentationBuilder"/> to produce a <see cref="BinaryOfficeDocument"/>.
/// </summary>
public static class PresentationBuilderExtensions
{
    /// <summary>
    /// Saves the presentation built by <paramref name="builder"/> to a byte array and returns it as a
    /// <see cref="BinaryOfficeDocument"/> with PPTX metadata.
    /// </summary>
    /// <param name="builder">The presentation builder. Must not be <see langword="null"/>.</param>
    /// <returns>A binary office document containing the PPTX bytes and format metadata.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="builder"/> is <see langword="null"/>.</exception>
    public static BinaryOfficeDocument ToBinaryDocument(this IPresentationBuilder builder)
    {
        if (builder is null)
        {
            throw new ArgumentNullException(nameof(builder));
        }

        return new BinaryOfficeDocument
        {
            Bytes = builder.SaveToBytes(),
            Format = OfficeDocumentFormat.Pptx,
            ContentType = "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            FileExtension = ".pptx"
        };
    }
}
