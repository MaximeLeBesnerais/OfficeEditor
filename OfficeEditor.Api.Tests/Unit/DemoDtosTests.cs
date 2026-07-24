using System.Text.Json;
using OfficeEditor.Api.Models;

namespace OfficeEditor.Api.Tests.Unit;

/// <summary>
/// Pins the shape of the demo DTO records via direct instantiation and JSON
/// round-trip tests. These records are the demo-endpoint wire contracts;
/// accidental shape changes break the web client.
/// </summary>
public sealed class DemoDtosTests
{
    [Fact]
    public void DemoDeckDto_MapsPositionalProperties()
    {
        var dto = new DemoDeckDto("Northwind", "northwind-demo.pptx", "A 15-slide demo deck", 15);

        Assert.Equal("Northwind", dto.Name);
        Assert.Equal("northwind-demo.pptx", dto.FileName);
        Assert.Equal("A 15-slide demo deck", dto.Description);
        Assert.Equal(15, dto.SlideCount);
    }

    [Fact]
    public void DemoDeckDto_Equality_HoldsForIdenticalValues()
    {
        var first = new DemoDeckDto("northwind", "demo.pptx", "desc", 15);
        var second = new DemoDeckDto("northwind", "demo.pptx", "desc", 15);

        Assert.Equal(first, second);
        Assert.Equal(first.GetHashCode(), second.GetHashCode());
    }

    [Fact]
    public void DemoDeckDto_Equality_FailsForDifferentValues()
    {
        var first = new DemoDeckDto("northwind", "demo.pptx", "desc", 15);
        var second = new DemoDeckDto("sales", "demo.pptx", "desc", 15);

        Assert.NotEqual(first, second);
    }

    [Fact]
    public void DemoRenderResponse_SuccessShape_AllPropertiesAreSet()
    {
        var deckId = Guid.NewGuid();
        var previews = new List<GeneratedSlidePreviewDto>
        {
            new(1, "png", "image/png", "base64content"),
            new(2, "png", "image/png", "base64content2")
        };

        var response = new DemoRenderResponse(
            Success: true,
            DeckId: deckId,
            SlideCount: 2,
            TotalMilliseconds: 1250.5,
            Previews: previews);

        Assert.True(response.Success);
        Assert.Equal(deckId, response.DeckId);
        Assert.Equal(2, response.SlideCount);
        Assert.Equal(1250.5, response.TotalMilliseconds);
        Assert.Equal(2, response.Previews.Count);
        Assert.Equal(1, response.Previews[0].Slide);
        Assert.Equal("png", response.Previews[0].Format);
        Assert.Equal("image/png", response.Previews[0].ContentType);
        Assert.Equal("base64content", response.Previews[0].ContentBase64);
    }

    [Fact]
    public void DemoRenderResponse_EmptyPreviews_HandlesEmptyList()
    {
        var response = new DemoRenderResponse(false, Guid.Empty, 0, 0, []);

        Assert.False(response.Success);
        Assert.Equal(Guid.Empty, response.DeckId);
        Assert.Equal(0, response.SlideCount);
        Assert.Equal(0, response.TotalMilliseconds);
        Assert.Empty(response.Previews);
    }

    [Fact]
    public void DemoRenderResponse_JsonRoundTrip_PreservesAllProperties()
    {
        var deckId = Guid.NewGuid();
        var original = new DemoRenderResponse(true, deckId, 3, 542.1,
            [new(1, "svg", "image/svg+xml", "PHN2Zy8+"), new(2, "svg", "image/svg+xml", "PHN2Zy8+")]);

        var json = JsonSerializer.Serialize(original);
        var deserialized = JsonSerializer.Deserialize<DemoRenderResponse>(json);

        Assert.NotNull(deserialized);
        Assert.Equal(original.Success, deserialized.Success);
        Assert.Equal(original.DeckId, deserialized.DeckId);
        Assert.Equal(original.SlideCount, deserialized.SlideCount);
        Assert.Equal(original.TotalMilliseconds, deserialized.TotalMilliseconds);
        Assert.Equal(original.Previews.Count, deserialized.Previews.Count);
        Assert.Equal(original.Previews[0], deserialized.Previews[0]);
        Assert.Contains("\"DeckId\"", json);
        Assert.Contains("\"SlideCount\"", json);
        Assert.Contains("\"TotalMilliseconds\"", json);
    }

    [Fact]
    public void OfficialSlidesResponse_MapsPositionalProperties()
    {
        var previews = new List<GeneratedSlidePreviewDto>
        {
            new(1, "png", "image/png", "base64slide1"),
            new(2, "png", "image/png", "base64slide2")
        };

        var response = new OfficialSlidesResponse(
            "northwind",
            2,
            "PowerPoint PDF export",
            previews);

        Assert.Equal("northwind", response.Name);
        Assert.Equal(2, response.SlideCount);
        Assert.Equal("PowerPoint PDF export", response.Source);
        Assert.Equal(2, response.Previews.Count);
        Assert.Equal(1, response.Previews[0].Slide);
        Assert.Equal("base64slide1", response.Previews[0].ContentBase64);
    }

    [Fact]
    public void OfficialSlidesResponse_JsonRoundTrip_PreservesAllProperties()
    {
        var original = new OfficialSlidesResponse("sales", 16, "PowerPoint PDF export",
            [new(1, "png", "image/png", "AAA="), new(16, "png", "image/png", "BBB=")]);

        var json = JsonSerializer.Serialize(original);
        var deserialized = JsonSerializer.Deserialize<OfficialSlidesResponse>(json);

        Assert.NotNull(deserialized);
        Assert.Equal(original.Name, deserialized.Name);
        Assert.Equal(original.SlideCount, deserialized.SlideCount);
        Assert.Equal(original.Source, deserialized.Source);
        Assert.Equal(original.Previews.Count, deserialized.Previews.Count);
        Assert.Equal(original.Previews[0], deserialized.Previews[0]);
        Assert.Contains("\"Name\"", json);
        Assert.Contains("\"SlideCount\"", json);
    }

    [Fact]
    public void CompareCapabilitiesDto_AvailableTrue_MapsAllProperties()
    {
        var dto = new CompareCapabilitiesDto(true, "24.8.3.2", true, null);

        Assert.True(dto.Available);
        Assert.Equal("24.8.3.2", dto.Version);
        Assert.True(dto.PdfToPpmAvailable);
        Assert.Null(dto.SkipReason);
    }

    [Fact]
    public void CompareCapabilitiesDto_AvailableFalse_CarriesSkipReason()
    {
        var dto = new CompareCapabilitiesDto(false, null, false, "soffice not found");

        Assert.False(dto.Available);
        Assert.Null(dto.Version);
        Assert.False(dto.PdfToPpmAvailable);
        Assert.Equal("soffice not found", dto.SkipReason);
    }

    [Fact]
    public void CompareCapabilitiesDto_JsonRoundTrip_PreservesAllProperties()
    {
        var original = new CompareCapabilitiesDto(true, "24.2.0.3", true, null);

        var json = JsonSerializer.Serialize(original);
        var deserialized = JsonSerializer.Deserialize<CompareCapabilitiesDto>(json);

        Assert.NotNull(deserialized);
        Assert.Equal(original.Available, deserialized.Available);
        Assert.Equal(original.Version, deserialized.Version);
        Assert.Equal(original.PdfToPpmAvailable, deserialized.PdfToPpmAvailable);
        Assert.Equal(original.SkipReason, deserialized.SkipReason);
        Assert.Contains("\"Available\"", json);
        Assert.Contains("\"PdfToPpmAvailable\"", json);
    }

    [Fact]
    public void CompareCapabilitiesDto_JsonRoundTrip_WithSkipReason_PreservesAllProperties()
    {
        var original = new CompareCapabilitiesDto(false, null, true, "LibreOffice not available on this machine");

        var json = JsonSerializer.Serialize(original);
        var deserialized = JsonSerializer.Deserialize<CompareCapabilitiesDto>(json);

        Assert.NotNull(deserialized);
        Assert.Equal(original.Available, deserialized.Available);
        Assert.Null(deserialized.Version);
        Assert.Equal(original.PdfToPpmAvailable, deserialized.PdfToPpmAvailable);
        Assert.Equal(original.SkipReason, deserialized.SkipReason);
    }

    [Fact]
    public void LibreOfficeLegResponse_SuccessShape_AllPropertiesSet()
    {
        var previews = new List<GeneratedSlidePreviewDto>
        {
            new(1, "png", "image/png", "YQ=="),
            new(2, "png", "image/png", "Yg=="),
            new(3, "png", "image/png", "Yw==")
        };

        var response = new LibreOfficeLegResponse(
            Success: true,
            Deck: "northwind",
            SlideCount: 3,
            Available: true,
            Version: "24.8.3",
            PdfToPpmAvailable: true,
            ConversionMilliseconds: 4500.2,
            RasterizationMilliseconds: 1200.5,
            TotalMilliseconds: 5700.7,
            Previews: previews,
            PdfDownloadUrl: "http://localhost:5001/api/download/abc-123",
            Error: null);

        Assert.True(response.Success);
        Assert.Equal("northwind", response.Deck);
        Assert.Equal(3, response.SlideCount);
        Assert.True(response.Available);
        Assert.Equal("24.8.3", response.Version);
        Assert.True(response.PdfToPpmAvailable);
        Assert.Equal(4500.2, response.ConversionMilliseconds);
        Assert.Equal(1200.5, response.RasterizationMilliseconds);
        Assert.Equal(5700.7, response.TotalMilliseconds);
        Assert.Equal(3, response.Previews!.Count);
        Assert.NotNull(response.PdfDownloadUrl);
        Assert.Null(response.Error);
    }

    [Fact]
    public void LibreOfficeLegResponse_UnavailableShape_AllOptionalPropertiesNull()
    {
        var response = new LibreOfficeLegResponse(
            Success: false,
            Deck: "northwind",
            SlideCount: 15,
            Available: false,
            Version: null,
            PdfToPpmAvailable: false,
            ConversionMilliseconds: null,
            RasterizationMilliseconds: null,
            TotalMilliseconds: null,
            Previews: null,
            PdfDownloadUrl: null,
            Error: "soffice not found");

        Assert.False(response.Success);
        Assert.Equal(15, response.SlideCount);
        Assert.False(response.Available);
        Assert.Null(response.Version);
        Assert.Null(response.ConversionMilliseconds);
        Assert.Null(response.RasterizationMilliseconds);
        Assert.Null(response.TotalMilliseconds);
        Assert.Null(response.Previews);
        Assert.Null(response.PdfDownloadUrl);
        Assert.Equal("soffice not found", response.Error);
    }

    [Fact]
    public void LibreOfficeLegResponse_JsonRoundTrip_PreservesAllProperties()
    {
        var original = new LibreOfficeLegResponse(
            true, "sales", 16, true, "24.8.3", true, 3100.0, 800.0, 3900.0,
            [new(1, "png", "image/png", "Zg=="), new(2, "png", "image/png", "Zw==")],
            "http://localhost:5001/api/download/x",
            null);

        var json = JsonSerializer.Serialize(original);
        var deserialized = JsonSerializer.Deserialize<LibreOfficeLegResponse>(json);

        Assert.NotNull(deserialized);
        Assert.Equal(original.Success, deserialized.Success);
        Assert.Equal(original.Deck, deserialized.Deck);
        Assert.Equal(original.SlideCount, deserialized.SlideCount);
        Assert.Equal(original.Available, deserialized.Available);
        Assert.Equal(original.Version, deserialized.Version);
        Assert.Equal(original.PdfToPpmAvailable, deserialized.PdfToPpmAvailable);
        Assert.Equal(original.ConversionMilliseconds, deserialized.ConversionMilliseconds);
        Assert.Equal(original.RasterizationMilliseconds, deserialized.RasterizationMilliseconds);
        Assert.Equal(original.TotalMilliseconds, deserialized.TotalMilliseconds);
        Assert.Equal(original.Previews!.Count, deserialized.Previews!.Count);
        Assert.Equal(original.Previews[0], deserialized.Previews[0]);
        Assert.Equal(original.PdfDownloadUrl, deserialized.PdfDownloadUrl);
        Assert.Equal(original.Error, deserialized.Error);
        Assert.Contains("\"Deck\"", json);
        Assert.Contains("\"SlideCount\"", json);
        Assert.Contains("\"ConversionMilliseconds\"", json);
    }

    [Fact]
    public void TypstLegResponse_SuccessShape_AllPropertiesSet()
    {
        var previews = new List<GeneratedSlidePreviewDto>
        {
            new(1, "png", "image/png", "AQ=="),
            new(2, "png", "image/png", "Ag==")
        };

        var response = new TypstLegResponse(
            Success: true,
            Deck: "aetherlink",
            SlideCount: 2,
            PngMilliseconds: 2300.0,
            PdfMilliseconds: 800.0,
            TotalMilliseconds: 2400.0,
            Previews: previews,
            PdfDownloadUrl: "http://localhost:5001/api/download/def-456",
            PdfError: null);

        Assert.True(response.Success);
        Assert.Equal("aetherlink", response.Deck);
        Assert.Equal(2, response.SlideCount);
        Assert.Equal(2300.0, response.PngMilliseconds);
        Assert.Equal(800.0, response.PdfMilliseconds);
        Assert.Equal(2400.0, response.TotalMilliseconds);
        Assert.Equal(2, response.Previews.Count);
        Assert.Equal("http://localhost:5001/api/download/def-456", response.PdfDownloadUrl);
        Assert.Null(response.PdfError);
    }

    [Fact]
    public void TypstLegResponse_PdfFailure_CarriesPdfError()
    {
        var response = new TypstLegResponse(
            Success: true,
            Deck: "northwind",
            SlideCount: 15,
            PngMilliseconds: 4000.0,
            PdfMilliseconds: null,
            TotalMilliseconds: 4200.0,
            Previews: [new(1, "png", "image/png", "Aw==")],
            PdfDownloadUrl: null,
            PdfError: "PDF compilation failed: Missing font family 'Aptos Display'");

        Assert.True(response.Success);
        Assert.Null(response.PdfMilliseconds);
        Assert.Null(response.PdfDownloadUrl);
        Assert.NotNull(response.PdfError);
        Assert.Contains("PDF compilation failed", response.PdfError);
    }

    [Fact]
    public void TypstLegResponse_JsonRoundTrip_PreservesAllProperties()
    {
        var original = new TypstLegResponse(
            true, "launch-review", 12, 3500.0, 1100.0, 3700.0,
            [new(1, "png", "image/png", "BA=="), new(2, "png", "image/png", "BQ==")],
            "http://localhost:5001/api/download/y",
            null);

        var json = JsonSerializer.Serialize(original);
        var deserialized = JsonSerializer.Deserialize<TypstLegResponse>(json);

        Assert.NotNull(deserialized);
        Assert.Equal(original.Success, deserialized.Success);
        Assert.Equal(original.Deck, deserialized.Deck);
        Assert.Equal(original.SlideCount, deserialized.SlideCount);
        Assert.Equal(original.PngMilliseconds, deserialized.PngMilliseconds);
        Assert.Equal(original.PdfMilliseconds, deserialized.PdfMilliseconds);
        Assert.Equal(original.TotalMilliseconds, deserialized.TotalMilliseconds);
        Assert.Equal(original.Previews.Count, deserialized.Previews.Count);
        Assert.Equal(original.Previews[0], deserialized.Previews[0]);
        Assert.Equal(original.PdfDownloadUrl, deserialized.PdfDownloadUrl);
        Assert.Equal(original.PdfError, deserialized.PdfError);
        Assert.Contains("\"PngMilliseconds\"", json);
        Assert.Contains("\"PdfDownloadUrl\"", json);
    }

    [Fact]
    public void TypstLegResponse_JsonRoundTrip_WithPdfError_PreservesAllProperties()
    {
        var original = new TypstLegResponse(
            true, "northwind", 15, 4200.0, null, 4200.0,
            [new(1, "png", "image/png", "Bg==")],
            null,
            "Font family not found");

        var json = JsonSerializer.Serialize(original);
        var deserialized = JsonSerializer.Deserialize<TypstLegResponse>(json);

        Assert.NotNull(deserialized);
        Assert.Null(deserialized.PdfMilliseconds);
        Assert.Null(deserialized.PdfDownloadUrl);
        Assert.Equal("Font family not found", deserialized.PdfError);
    }

    [Fact]
    public void DemoDeckDto_Deserialize_FromPinnedJson()
    {
        var json = """{"Name":"Northwind","FileName":"northwind-demo.pptx","Description":"A 15-slide demo deck","SlideCount":15}""";
        var dto = JsonSerializer.Deserialize<DemoDeckDto>(json);

        Assert.NotNull(dto);
        Assert.Equal("Northwind", dto.Name);
        Assert.Equal("northwind-demo.pptx", dto.FileName);
        Assert.Equal("A 15-slide demo deck", dto.Description);
        Assert.Equal(15, dto.SlideCount);
    }

    [Fact]
    public void CompareCapabilitiesDto_Deserialize_FromPinnedJson()
    {
        var json = """{"Available":true,"Version":"24.8.3.2","PdfToPpmAvailable":true,"SkipReason":null}""";
        var dto = JsonSerializer.Deserialize<CompareCapabilitiesDto>(json);

        Assert.NotNull(dto);
        Assert.True(dto.Available);
        Assert.Equal("24.8.3.2", dto.Version);
        Assert.True(dto.PdfToPpmAvailable);
        Assert.Null(dto.SkipReason);
    }

    [Fact]
    public void OfficialSlidesResponse_Deserialize_FromPinnedJson()
    {
        var json = """{"Name":"sales","SlideCount":16,"Source":"PowerPoint PDF export","Previews":[{"Slide":1,"Format":"png","ContentType":"image/png","ContentBase64":"AAA="}]}""";
        var dto = JsonSerializer.Deserialize<OfficialSlidesResponse>(json);

        Assert.NotNull(dto);
        Assert.Equal("sales", dto.Name);
        Assert.Equal(16, dto.SlideCount);
        Assert.Equal("PowerPoint PDF export", dto.Source);
        Assert.Single(dto.Previews);
        Assert.Equal(1, dto.Previews[0].Slide);
        Assert.Equal("png", dto.Previews[0].Format);
        Assert.Equal("AAA=", dto.Previews[0].ContentBase64);
    }

    [Fact]
    public void DemoRenderResponse_Deserialize_FromPinnedJson()
    {
        var json = """{"Success":true,"DeckId":"00000000-0000-0000-0000-000000000001","SlideCount":3,"TotalMilliseconds":542.1,"Previews":[{"Slide":1,"Format":"svg","ContentType":"image/svg+xml","ContentBase64":"PHN2Zy8+"}]}""";
        var dto = JsonSerializer.Deserialize<DemoRenderResponse>(json);

        Assert.NotNull(dto);
        Assert.True(dto.Success);
        Assert.Equal(Guid.Parse("00000000-0000-0000-0000-000000000001"), dto.DeckId);
        Assert.Equal(3, dto.SlideCount);
        Assert.Equal(542.1, dto.TotalMilliseconds);
        Assert.Single(dto.Previews);
        Assert.Equal(1, dto.Previews[0].Slide);
        Assert.Equal("svg", dto.Previews[0].Format);
        Assert.Equal("image/svg+xml", dto.Previews[0].ContentType);
        Assert.Equal("PHN2Zy8+", dto.Previews[0].ContentBase64);
    }

    [Fact]
    public void AllDemoDtos_AreRecordTypes_WithValueEquality()
    {
        var dto1 = new DemoDeckDto("a", "b.pptx", "c", 1);
        var dto2 = new DemoDeckDto("a", "b.pptx", "c", 1);
        Assert.Equal(dto1, dto2);

        var caps1 = new CompareCapabilitiesDto(true, "v1", true, null);
        var caps2 = new CompareCapabilitiesDto(true, "v1", true, null);
        Assert.Equal(caps1, caps2);

        // OfficialSlidesResponse and other IReadOnlyList-bearing records are
        // excluded — C# record equality uses reference equality for interfaces,
        // so two instances with identical list contents are not .Equal.
    }
}
