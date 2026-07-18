using OfficeEditor.Api.Services;

namespace OfficeEditor.Api.Tests.Unit;

public sealed class DeckPreviewValidatorsTests
{
    [Theory]
    [InlineData(1, 3)]
    [InlineData(2, 3)]
    [InlineData(3, 3)]
    [InlineData(1, 1)]
    public void ValidateSlideNumber_InRange_ReturnsNull(int slideNumber, int slideCount)
    {
        Assert.Null(DeckPreviewValidators.ValidateSlideNumber(slideNumber, slideCount));
    }

    [Theory]
    [InlineData(0, 3)]
    [InlineData(-1, 3)]
    [InlineData(4, 3)]
    [InlineData(2, 1)]
    [InlineData(1, 0)]
    public void ValidateSlideNumber_OutOfRange_ReturnsError(int slideNumber, int slideCount)
    {
        var error = DeckPreviewValidators.ValidateSlideNumber(slideNumber, slideCount);

        Assert.NotNull(error);
        Assert.NotEmpty(error!);
    }

    [Theory]
    [InlineData(null, "png")]
    [InlineData("png", "png")]
    [InlineData("PNG", "png")]
    [InlineData(" svg ", "svg")]
    public void TryNormalizeFormat_SupportedFormats_NormalizeToLowerInvariant(string? input, string expected)
    {
        var ok = DeckPreviewValidators.TryNormalizeFormat(input, out var normalized, out var error);

        Assert.True(ok);
        Assert.Equal(expected, normalized);
        Assert.Null(error);
    }

    [Theory]
    [InlineData("")]
    [InlineData("gif")]
    [InlineData("pdf")]
    [InlineData("jpeg")]
    public void TryNormalizeFormat_UnsupportedFormats_ReturnFalseWithError(string? input)
    {
        var ok = DeckPreviewValidators.TryNormalizeFormat(input, out _, out var error);

        Assert.False(ok);
        Assert.NotNull(error);
    }

    [Theory]
    [InlineData(null, 150)]
    [InlineData(150, 150)]
    [InlineData(36, 36)]
    [InlineData(600, 600)]
    [InlineData(1, 36)]
    [InlineData(0, 36)]
    [InlineData(-50, 36)]
    [InlineData(601, 600)]
    [InlineData(10000, 600)]
    public void ClampPpi_ClampsTo36Through600(int? input, int expected)
    {
        Assert.Equal(expected, DeckPreviewValidators.ClampPpi(input));
    }

    [Theory]
    [InlineData("png", "image/png")]
    [InlineData("svg", "image/svg+xml")]
    public void ContentTypeForFormat_MapsToExpectedContentType(string format, string expected)
    {
        Assert.Equal(expected, DeckPreviewValidators.ContentTypeForFormat(format));
    }
}
