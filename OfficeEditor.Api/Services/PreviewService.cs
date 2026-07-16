namespace OfficeEditor.Api.Services;

public interface IPreviewService
{
    bool CanPreview(string contentType);
}

public sealed class PreviewService : IPreviewService
{
    private static readonly HashSet<string> PreviewableContentTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "application/pdf",
        "image/png",
        "image/svg+xml"
    };

    public bool CanPreview(string contentType) =>
        !string.IsNullOrWhiteSpace(contentType) && PreviewableContentTypes.Contains(contentType);
}
