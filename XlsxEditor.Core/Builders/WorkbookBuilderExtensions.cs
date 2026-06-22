using OfficeEditor.Core.Services;

namespace XlsxEditor.Core.Builders;

/// <summary>
/// Extension methods for <see cref="IWorkbookBuilder"/> to produce a <see cref="BinaryOfficeDocument"/>.
/// </summary>
public static class WorkbookBuilderExtensions
{
    /// <summary>
    /// Saves the workbook built by <paramref name="builder"/> to a byte array and returns it as a
    /// <see cref="BinaryOfficeDocument"/> with XLSX metadata.
    /// </summary>
    /// <param name="builder">The workbook builder. Must not be <see langword="null"/>.</param>
    /// <returns>A binary office document containing the XLSX bytes and format metadata.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="builder"/> is <see langword="null"/>.</exception>
    public static BinaryOfficeDocument ToBinaryDocument(this IWorkbookBuilder builder)
    {
        if (builder is null)
        {
            throw new ArgumentNullException(nameof(builder));
        }

        return new BinaryOfficeDocument
        {
            Bytes = builder.SaveToBytes(),
            Format = OfficeDocumentFormat.Xlsx,
            ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            FileExtension = ".xlsx"
        };
    }
}
