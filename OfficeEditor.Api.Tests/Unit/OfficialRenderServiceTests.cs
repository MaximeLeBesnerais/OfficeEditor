using Microsoft.Extensions.Logging.Abstractions;
using OfficeEditor.Api.Services;

namespace OfficeEditor.Api.Tests.Unit;

/// <summary>
/// Official-render (PowerPoint ground truth) service. Unknown-name and cache-shape tests
/// are environment-invariant; the happy path rasterizes the sibling REF PDF with
/// pdftoppm, so it skips gracefully when pdftoppm is absent (same gating convention as
/// LibreOfficeCompareServiceTests).
/// </summary>
public sealed class OfficialRenderServiceTests
{
    private static OfficialRenderService CreateService() =>
        new(new DemoDeckService(new StubDeckSessionStore()),
            NullLogger<OfficialRenderService>.Instance);

    [Theory]
    [InlineData("no-such-deck")]
    [InlineData("")]
    [InlineData("   ")]
    public void TryGetOfficialSlides_UnknownOrBlankName_ReturnsFalse(string name)
    {
        var service = CreateService();

        Assert.False(service.TryGetOfficialSlides(name, out var result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetOfficialSlides_SalesDeck_RasterizesAllSlides()
    {
        var service = CreateService();
        if (!service.PdfToPpmAvailable)
        {
            return; // no pdftoppm on this machine (see class summary)
        }

        Assert.True(service.TryGetOfficialSlides("sales", out var result));

        Assert.NotNull(result);
        Assert.Equal("sales", result.Name);
        Assert.Equal(16, result.SlideCount);
        Assert.Equal(result.SlideCount, result.Pages.Count);
        Assert.All(result.Pages, page =>
        {
            // PNG magic: 0x89 'P' 'N' 'G'
            Assert.True(page.Length > 8);
            Assert.Equal(0x89, page[0]);
            Assert.Equal((byte)'P', page[1]);
            Assert.Equal((byte)'N', page[2]);
            Assert.Equal((byte)'G', page[3]);
        });
    }

    [Fact]
    public void TryGetOfficialSlides_SecondCall_ReturnsCachedInstance()
    {
        var service = CreateService();
        if (!service.PdfToPpmAvailable)
        {
            return; // no pdftoppm on this machine (see class summary)
        }

        Assert.True(service.TryGetOfficialSlides("sales", out var first));
        Assert.True(service.TryGetOfficialSlides("SALES", out var second));

        Assert.Same(first, second);
    }
}
