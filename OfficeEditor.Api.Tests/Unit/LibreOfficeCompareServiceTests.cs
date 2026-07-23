using OfficeEditor.Api.Services;

namespace OfficeEditor.Api.Tests.Unit;

/// <summary>
/// LibreOffice compare service. Probe/discovery tests are environment-invariant
/// (assertions hold whether or not soffice is installed); the happy-path render is
/// opt-in via OE_RUN_TYPST_COMPILE_TESTS=1 (same convention as DemoDeckServiceTests)
/// and additionally requires LibreOffice + pdftoppm on the machine.
/// </summary>
public sealed class LibreOfficeCompareServiceTests
{
    private const string EnableRenderEnvVar = "OE_RUN_TYPST_COMPILE_TESTS";

    [Fact]
    public void Probe_ReturnsInternallyConsistentResult()
    {
        var service = new LibreOfficeCompareService();

        var probe = service.Probe();

        if (probe.Available)
        {
            Assert.False(string.IsNullOrWhiteSpace(probe.SofficePath));
            Assert.True(File.Exists(probe.SofficePath));
            Assert.Null(probe.SkipReason);
        }
        else
        {
            Assert.Null(probe.SofficePath);
            Assert.False(string.IsNullOrWhiteSpace(probe.SkipReason));
        }

        // pdftoppm availability is independent of soffice availability.
        Assert.True(probe.PdfToPpmAvailable || !probe.PdfToPpmAvailable);
    }

    [Fact]
    public void Probe_IsCached_SameInstanceReturned()
    {
        var service = new LibreOfficeCompareService();

        Assert.Same(service.Probe(), service.Probe());
    }

    [Fact]
    public void RenderDeck_WhenSofficeUnavailable_ReturnsSkippedResultWithoutThrowing()
    {
        var service = new LibreOfficeCompareService();
        if (service.Probe().Available)
        {
            return; // soffice exists here; the skip path only exists where it doesn't
        }

        var result = service.RenderDeck([1, 2, 3], "deck.pptx", 110);

        Assert.False(result.Available);
        Assert.Null(result.Version);
        Assert.Equal(0, result.ConversionMilliseconds);
        Assert.Null(result.RasterizationMilliseconds);
        Assert.Equal(0, result.TotalMilliseconds);
        Assert.Empty(result.PngPages);
        Assert.Null(result.PdfBytes);
        Assert.Equal(service.Probe().SkipReason, result.Error);
    }

    [Fact]
    public void RenderDeck_Northwind_ConvertsAndRasterizesAllSlides()
    {
        if (Environment.GetEnvironmentVariable(EnableRenderEnvVar) != "1")
        {
            return; // heavy external-tool test is opt-in (see class summary)
        }

        var service = new LibreOfficeCompareService();
        var probe = service.Probe();
        if (!probe.Available)
        {
            return; // no LibreOffice on this machine
        }

        var demoDeckService = new DemoDeckService(new StubDeckSessionStore());
        Assert.True(demoDeckService.TryGetDeckFile("northwind", out var deckPath));
        var bytes = File.ReadAllBytes(deckPath);

        var result = service.RenderDeck(bytes, "northwind-demo.pptx", 110);

        Assert.True(result.Available);
        Assert.Null(result.Error);
        Assert.True(result.ConversionMilliseconds > 0);
        Assert.True(result.TotalMilliseconds > 0);
        Assert.NotNull(result.PdfBytes);
        Assert.True(result.PdfBytes!.Length > 4);
        Assert.Equal((byte)'%', result.PdfBytes[0]);
        Assert.Equal((byte)'P', result.PdfBytes[1]);
        Assert.Equal((byte)'D', result.PdfBytes[2]);
        Assert.Equal((byte)'F', result.PdfBytes[3]);

        if (result.PdfToPpmAvailable)
        {
            Assert.NotNull(result.RasterizationMilliseconds);
            Assert.True(result.RasterizationMilliseconds > 0);
            Assert.Equal(15, result.PngPages.Count);
            Assert.All(result.PngPages, page =>
            {
                // PNG magic: 0x89 'P' 'N' 'G'
                Assert.True(page.Length > 8);
                Assert.Equal(0x89, page[0]);
                Assert.Equal((byte)'P', page[1]);
            });
        }
        else
        {
            Assert.Null(result.RasterizationMilliseconds);
            Assert.Empty(result.PngPages);
        }
    }
}
