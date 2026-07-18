using OfficeEditor.Api.Services;

namespace OfficeEditor.Api.Tests.Unit;

public sealed class PreviewServiceTests
{
    private readonly PreviewService _service = new();

    [Theory]
    [InlineData("application/pdf")]
    [InlineData("image/png")]
    [InlineData("image/svg+xml")]
    public void CanPreview_AllowedContentTypes_ReturnsTrue(string contentType)
    {
        Assert.True(_service.CanPreview(contentType));
    }

    [Theory]
    [InlineData("APPLICATION/PDF")]
    [InlineData("Image/Png")]
    [InlineData("IMAGE/SVG+XML")]
    public void CanPreview_MatchesCaseInsensitively(string contentType)
    {
        Assert.True(_service.CanPreview(contentType));
    }

    [Theory]
    [InlineData("application/vnd.openxmlformats-officedocument.wordprocessingml.document")]
    [InlineData("application/vnd.openxmlformats-officedocument.presentationml.presentation")]
    [InlineData("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")]
    [InlineData("application/octet-stream")]
    [InlineData("text/plain")]
    [InlineData("image/jpeg")]
    [InlineData("image/gif")]
    [InlineData("application/pdf ")]
    [InlineData(" application/pdf")]
    public void CanPreview_DisallowedContentTypes_ReturnsFalse(string contentType)
    {
        Assert.False(_service.CanPreview(contentType));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void CanPreview_NullOrBlank_ReturnsFalse(string? contentType)
    {
        Assert.False(_service.CanPreview(contentType!));
    }
}
