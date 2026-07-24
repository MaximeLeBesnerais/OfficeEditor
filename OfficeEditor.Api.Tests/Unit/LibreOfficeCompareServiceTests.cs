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
    }

    [Fact]
    public void Probe_IsCached_SameInstanceReturned()
    {
        var service = new LibreOfficeCompareService();

        Assert.Same(service.Probe(), service.Probe());
    }

    [Fact]
    public void Probe_SeparateInstances_EachProbeIndependently()
    {
        var first = new LibreOfficeCompareService();
        var second = new LibreOfficeCompareService();

        var probe1 = first.Probe();
        var probe2 = second.Probe();

        Assert.Equal(probe1.Available, probe2.Available);
        Assert.Equal(probe1.SofficePath, probe2.SofficePath);
    }

    [Fact]
    public void Probe_WhenAvailable_VersionIsNotNullOrEmpty()
    {
        var service = new LibreOfficeCompareService();
        var probe = service.Probe();
        if (!probe.Available)
        {
            return;
        }

        Assert.NotNull(probe.Version);
        Assert.NotEmpty(probe.Version);
    }

    [Fact]
    public void RenderDeck_NullBytes_ThrowsArgumentNullException()
    {
        var service = new LibreOfficeCompareService();

        var ex = Assert.Throws<ArgumentNullException>(
            () => service.RenderDeck(null!, "deck.pptx", 110));

        Assert.Equal("pptxBytes", ex.ParamName);
    }

    [Fact]
    public void RenderDeck_NullFileName_ThrowsArgumentNullException()
    {
        var service = new LibreOfficeCompareService();

        var ex = Assert.Throws<ArgumentNullException>(
            () => service.RenderDeck([1, 2, 3], null!, 110));

        Assert.Equal("fileName", ex.ParamName);
    }

    [Fact]
    public void RenderDeck_EmptyFileName_ThrowsArgumentException()
    {
        var service = new LibreOfficeCompareService();

        var ex = Assert.Throws<ArgumentException>(
            () => service.RenderDeck([1, 2, 3], "", 110));

        Assert.Equal("fileName", ex.ParamName);
    }

    [Theory]
    [InlineData("   ")]
    [InlineData("\t")]
    public void RenderDeck_WhitespaceFileName_ThrowsArgumentException(string fileName)
    {
        var service = new LibreOfficeCompareService();

        var ex = Assert.Throws<ArgumentException>(
            () => service.RenderDeck([1, 2, 3], fileName, 110));

        Assert.Equal("fileName", ex.ParamName);
    }

    [Fact]
    public void RenderDeck_WhenSofficeUnavailable_ReturnsSkippedResultWithoutThrowing()
    {
        var service = new LibreOfficeCompareService();
        if (service.Probe().Available)
        {
            return;
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
            return;
        }

        var service = new LibreOfficeCompareService();
        var probe = service.Probe();
        if (!probe.Available)
        {
            return;
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

    [Fact]
    public void RenderDeck_WithValidDeck_SanitizesFileName()
    {
        if (Environment.GetEnvironmentVariable(EnableRenderEnvVar) != "1")
        {
            return;
        }

        var service = new LibreOfficeCompareService();
        if (!service.Probe().Available)
        {
            return;
        }

        var pptxContent = BuildMinimalPptx();

        // File name with path separators and special chars — SanitizeFileName replaces them.
        var result = service.RenderDeck(pptxContent, "deck/with\\invalid:chars.pptx", 72);

        Assert.True(result.Available);
        Assert.Null(result.Error);
        Assert.NotNull(result.PdfBytes);
        Assert.True(result.PdfBytes!.Length > 4);
        Assert.Equal("%PDF-", System.Text.Encoding.ASCII.GetString(result.PdfBytes, 0, 5));
    }

    [Fact]
    public void LibreOfficeProbe_AvailableRecord_PropertiesMatch()
    {
        var probe = new LibreOfficeProbe(
            Available: true,
            SofficePath: "/opt/homebrew/bin/soffice",
            Version: "24.8.3.2",
            PdfToPpmAvailable: true,
            SkipReason: null);

        Assert.True(probe.Available);
        Assert.Equal("/opt/homebrew/bin/soffice", probe.SofficePath);
        Assert.Equal("24.8.3.2", probe.Version);
        Assert.True(probe.PdfToPpmAvailable);
        Assert.Null(probe.SkipReason);
    }

    [Fact]
    public void LibreOfficeProbe_UnavailableRecord_PropertiesMatch()
    {
        var probe = new LibreOfficeProbe(
            Available: false,
            SofficePath: null,
            Version: null,
            PdfToPpmAvailable: false,
            SkipReason: "soffice not found");

        Assert.False(probe.Available);
        Assert.Null(probe.SofficePath);
        Assert.Null(probe.Version);
        Assert.False(probe.PdfToPpmAvailable);
        Assert.Equal("soffice not found", probe.SkipReason);
    }

    [Fact]
    public void LibreOfficeRenderResult_SuccessRecord_AllPropertiesSet()
    {
        var pages = new List<byte[]> { new byte[] { 0x89, (byte)'P', (byte)'N', (byte)'G' } };
        byte[] pdf = new byte[] { 0x25, 0x50, 0x44, 0x46 };

        var result = new LibreOfficeRenderResult(
            Available: true,
            Version: "24.8.3",
            PdfToPpmAvailable: true,
            ConversionMilliseconds: 2500.0,
            RasterizationMilliseconds: 800.0,
            TotalMilliseconds: 3300.0,
            PngPages: pages,
            PdfBytes: pdf,
            Error: null);

        Assert.True(result.Available);
        Assert.Equal("24.8.3", result.Version);
        Assert.True(result.PdfToPpmAvailable);
        Assert.Equal(2500.0, result.ConversionMilliseconds);
        Assert.Equal(800.0, result.RasterizationMilliseconds);
        Assert.Equal(3300.0, result.TotalMilliseconds);
        Assert.Single(result.PngPages);
        Assert.NotNull(result.PdfBytes);
        Assert.Null(result.Error);
    }

    [Fact]
    public void LibreOfficeRenderResult_ErrorRecord_CarriesError()
    {
        var result = new LibreOfficeRenderResult(
            Available: false,
            Version: null,
            PdfToPpmAvailable: false,
            ConversionMilliseconds: 0,
            RasterizationMilliseconds: null,
            TotalMilliseconds: 0,
            PngPages: [],
            PdfBytes: null,
            Error: "soffice conversion timed out");

        Assert.False(result.Available);
        Assert.Equal(0, result.ConversionMilliseconds);
        Assert.Null(result.RasterizationMilliseconds);
        Assert.Empty(result.PngPages);
        Assert.Null(result.PdfBytes);
        Assert.Equal("soffice conversion timed out", result.Error);
    }

    private static byte[] BuildMinimalPptx()
    {
        using var builder = PptxEditor.Core.Builders.PresentationBuilder.Create();
        builder.AddSlide();
        builder.CurrentSlide.AddTitle("Test Slide");
        return builder.SaveToBytes();
    }
}
