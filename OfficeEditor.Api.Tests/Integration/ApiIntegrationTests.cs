using System.Net;
using System.Net.Http.Json;
using System.Text.Json;
using Microsoft.AspNetCore.Mvc.Testing;
using OfficeEditor.Api.Services;
using PptxEditor.Core.Builders;

namespace OfficeEditor.Api.Tests.Integration;

public sealed class ApiIntegrationTests : IClassFixture<WebApplicationFactory<Program>>
{
    private readonly HttpClient _client;

    public ApiIntegrationTests(WebApplicationFactory<Program> factory)
    {
        _client = factory.CreateClient();
    }

    [Fact]
    public async Task Health_ReturnsOk()
    {
        var response = await _client.GetAsync("/api/health");

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        var body = await response.Content.ReadAsStringAsync();
        Assert.Contains("ok", body);
    }

    [Fact]
    public async Task RequestLargerThanConfiguredLimit_ReturnsPayloadTooLarge()
    {
        using var content = new StringContent("{}");
        content.Headers.ContentLength = ApiResourceLimits.MaxRequestBodyBytes + 1;

        var response = await _client.PostAsync("/api/decks/generate", content);

        Assert.Equal(HttpStatusCode.RequestEntityTooLarge, response.StatusCode);
    }

    [Fact]
    public async Task Samples_ReturnsNonEmptyList()
    {
        var response = await _client.GetAsync("/api/samples");

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        var body = await response.Content.ReadAsStringAsync();
        Assert.Contains("\"name\"", body);
        Assert.Contains("\"format\"", body);
    }

    [Fact]
    public async Task DemoDecks_ReturnsNonEmptyList()
    {
        var response = await _client.GetAsync("/api/demo/decks");

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        var body = await response.Content.ReadAsStringAsync();
        Assert.Contains("\"name\"", body);
        Assert.Contains("\"decks\"", body);
    }

    [Fact]
    public async Task DeckTemplate_ReturnsValidJson()
    {
        var response = await _client.GetAsync("/api/demo/deck-template");

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        var json = await response.Content.ReadAsStringAsync();
        Assert.Contains("\"version\"", json);
        Assert.Contains("\"slides\"", json);
    }

    [Fact]
    public async Task CreateDeck_WithValidPptx_ReturnsSession()
    {
        var pptxBytes = CreateMinimalPptx();
        using var content = new MultipartFormDataContent();
        var fileContent = new ByteArrayContent(pptxBytes);
        fileContent.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.presentationml.presentation");
        content.Add(fileContent, "file", "test.pptx");

        var response = await _client.PostAsync("/api/decks", content);

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        var result = await response.Content.ReadFromJsonAsync<JsonElement>();
        Assert.True(result.TryGetProperty("deckId", out _));
        Assert.True(result.TryGetProperty("slideCount", out var slideCount));
        Assert.True(slideCount.GetInt32() > 0);
    }

    [Fact]
    public async Task GetDeckAnatomy_WithValidDeck_ReturnsSlideData()
    {
        var deckId = await CreateDeckSession();
        ArgumentNullException.ThrowIfNull(deckId);

        var response = await _client.GetAsync($"/api/decks/{deckId}/anatomy");

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        var anatomy = await response.Content.ReadFromJsonAsync<JsonElement>();
        Assert.True(anatomy.TryGetProperty("slides", out var slides));
        Assert.True(slides.GetArrayLength() > 0);
    }

    [Fact]
    public async Task GetDeckFile_WithValidDeck_ReturnsPptxBytes()
    {
        var deckId = await CreateDeckSession();
        ArgumentNullException.ThrowIfNull(deckId);

        var response = await _client.GetAsync($"/api/decks/{deckId}/file");

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        var bytes = await response.Content.ReadAsByteArrayAsync();
        Assert.True(bytes.Length > 0);
        Assert.StartsWith("application/vnd.openxmlformats-officedocument.presentationml.presentation",
            response.Content.Headers.ContentType?.MediaType ?? "");
    }

    [Fact]
    public async Task GenerateDeck_WithValidDocument_ReturnsSuccess()
    {
        var genJson = """
        {
          "document": {
            "version": "2.0",
            "design": {
              "palette": { "primary": "#0B3D91", "accent": "#1E7BC6", "ink": "#1A1A1A", "paper": "#FFFFFF", "muted": "#8A94A6" },
              "fonts": { "display": "Aptos Display", "body": "Aptos" },
              "metrics": { "marginPt": 43, "gutterPt": 18, "titleSizePt": 30, "bodySizePt": 14 }
            },
            "slides": [
              {
                "type": "container",
                "fill": "paper",
                "overflow": "clip",
                "layout": { "mode": "row", "gap": 18 },
                "padding": 43,
                "children": [
                  {
                    "type": "card",
                    "size": { "grow": 1, "aspect": "4:3" },
                    "content": { "title": "Generated", "subtitle": "Integration Test" }
                  }
                ]
              }
            ]
          },
          "previewFormat": "svg",
          "ppi": 75
        }
        """;
        using var content = new StringContent(genJson, System.Text.Encoding.UTF8, "application/json");

        var response = await _client.PostAsync("/api/decks/generate", content);

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        var result = await response.Content.ReadFromJsonAsync<JsonElement>();
        Assert.True(result.TryGetProperty("deckId", out var deckId));
        Assert.NotEqual(Guid.Empty.ToString(), deckId.GetString());
        Assert.True(result.TryGetProperty("slideCount", out var slideCount));
        Assert.Equal(1, slideCount.GetInt32());
    }

    [Fact]
    public async Task GenerateDeck_InvalidJson_ReturnsBadRequest()
    {
        var genJson = """{ "document": "not a valid generation document" }""";
        using var content = new StringContent(genJson, System.Text.Encoding.UTF8, "application/json");

        var response = await _client.PostAsync("/api/decks/generate", content);

        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
    }

    [Fact]
    public async Task Convert_SampleFile_ReturnsResult()
    {
        const string enableEnvVar = "OE_RUN_TYPST_COMPILE_TESTS";
        if (Environment.GetEnvironmentVariable(enableEnvVar) != "1")
        {
            return;
        }

        using var content = new MultipartFormDataContent();
        content.Add(new StringContent("northwind-demo"), "sampleName");
        content.Add(new StringContent("pdf"), "targetFormat");

        var response = await _client.PostAsync("/api/convert", content);

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        var bytes = await response.Content.ReadAsByteArrayAsync();
        Assert.True(bytes.Length > 0, "Converted PDF payload should be non-empty");
    }

    [Fact]
    public async Task Convert_JsonFileToXlsx_ReturnsDownloadUrl()
    {
        var json = """{ "version": "1.0", "worksheets": [ { "name": "S", "headers": ["A"], "rows": [["1"]] } ] }""";
        using var content = new MultipartFormDataContent();
        var fileContent = new StringContent(json, System.Text.Encoding.UTF8, "application/json");
        content.Add(fileContent, "file", "workbook.json");
        content.Add(new StringContent("xlsx"), "targetFormat");

        var response = await _client.PostAsync("/api/convert", content);

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        var result = await response.Content.ReadFromJsonAsync<JsonElement>();
        Assert.True(result.TryGetProperty("success", out var success) && success.GetBoolean());
        Assert.True(result.TryGetProperty("downloadUrl", out var url) && !string.IsNullOrEmpty(url.GetString()));
    }

    [Fact]
    public async Task Convert_JsonFileToDocx_ReturnsDownloadUrl()
    {
        var json = """{ "version": "1.0", "sections": [ { "blocks": [ { "type": "paragraph", "text": "Hello" } ] } ] }""";
        using var content = new MultipartFormDataContent();
        var fileContent = new StringContent(json, System.Text.Encoding.UTF8, "application/json");
        content.Add(fileContent, "file", "report.json");
        content.Add(new StringContent("docx"), "targetFormat");

        var response = await _client.PostAsync("/api/convert", content);

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        var result = await response.Content.ReadFromJsonAsync<JsonElement>();
        Assert.True(result.TryGetProperty("success", out var success) && success.GetBoolean());
        Assert.True(result.TryGetProperty("downloadUrl", out var url) && !string.IsNullOrEmpty(url.GetString()));
    }

    [Fact]
    public async Task Convert_InvalidJsonToXlsx_ReturnsBadRequestWithError()
    {
        var json = """{ "version": "9.9", "worksheets": [] }""";
        using var content = new MultipartFormDataContent();
        var fileContent = new StringContent(json, System.Text.Encoding.UTF8, "application/json");
        content.Add(fileContent, "file", "bad.json");
        content.Add(new StringContent("xlsx"), "targetFormat");

        var response = await _client.PostAsync("/api/convert", content);

        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
        var body = await response.Content.ReadAsStringAsync();
        Assert.Contains("Invalid XLSX instruction JSON", body);
    }

    [Fact]
    public async Task Download_InvalidId_ReturnsNotFound()
    {
        var response = await _client.GetAsync($"/api/download/{Guid.NewGuid()}");

        Assert.Equal(HttpStatusCode.NotFound, response.StatusCode);
    }

    [Fact]
    public async Task Preview_InvalidId_ReturnsNotFound()
    {
        var response = await _client.GetAsync($"/api/preview/{Guid.NewGuid()}");

        Assert.Equal(HttpStatusCode.NotFound, response.StatusCode);
    }

    [Fact]
    public async Task DeckAnatomy_InvalidId_ReturnsNotFound()
    {
        var response = await _client.GetAsync($"/api/decks/{Guid.NewGuid()}/anatomy");

        Assert.Equal(HttpStatusCode.NotFound, response.StatusCode);
    }

    [Fact]
    public async Task DeckPreview_InvalidId_ReturnsNotFound()
    {
        var response = await _client.GetAsync($"/api/decks/{Guid.NewGuid()}/slides/1/preview");

        Assert.Equal(HttpStatusCode.NotFound, response.StatusCode);
    }

    [Fact]
    public async Task PostInstructions_InvalidDeck_ReturnsNotFound()
    {
        var instructions = """{ "operations": [] }""";
        using var content = new StringContent(instructions, System.Text.Encoding.UTF8, "application/json");

        var response = await _client.PostAsync($"/api/decks/{Guid.NewGuid()}/instructions", content);

        Assert.Equal(HttpStatusCode.NotFound, response.StatusCode);
    }

    [Fact]
    public async Task DemoRender_InvalidName_ReturnsBadRequest()
    {
        var body = """{ "name": "nonexistent", "ppi": 110, "format": "svg" }""";
        using var content = new StringContent(body, System.Text.Encoding.UTF8, "application/json");

        var response = await _client.PostAsync("/api/demo/render", content);

        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
    }

    [Fact]
    public async Task CompareCapabilities_ReturnsResult()
    {
        var response = await _client.GetAsync("/api/demo/compare/capabilities");

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        var caps = await response.Content.ReadFromJsonAsync<JsonElement>();
        Assert.True(caps.TryGetProperty("available", out _));
    }

    [Fact]
    public async Task CreateDeck_InvalidFile_ReturnsBadRequest()
    {
        using var content = new MultipartFormDataContent();
        var fileContent = new ByteArrayContent("not a pptx"u8.ToArray());
        fileContent.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/octet-stream");
        content.Add(fileContent, "file", "test.pptx");

        var response = await _client.PostAsync("/api/decks", content);

        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
    }

    [Fact]
    public async Task PostInstructions_WithValidDeck_ReturnsResult()
    {
        var deckId = await CreateDeckSession();
        ArgumentNullException.ThrowIfNull(deckId);

        var instructions = """{ "operations": [] }""";
        using var content = new StringContent(instructions, System.Text.Encoding.UTF8, "application/json");

        var response = await _client.PostAsync($"/api/decks/{deckId}/instructions", content);

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
    }

    [Fact]
    public async Task DemoCompare_LibreOffice_MissingDeck_ReturnsBadRequest()
    {
        var body = """{ "name": "nonexistent-deck", "ppi": 110 }""";
        using var content = new StringContent(body, System.Text.Encoding.UTF8, "application/json");

        var response = await _client.PostAsync("/api/demo/compare/libreoffice", content);

        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
    }

    [Fact]
    public async Task DemoCompare_Typst_MissingDeck_ReturnsBadRequest()
    {
        var body = """{ "name": "nonexistent-deck", "ppi": 110 }""";
        using var content = new StringContent(body, System.Text.Encoding.UTF8, "application/json");

        var response = await _client.PostAsync("/api/demo/compare/typst", content);

        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
    }

    private async Task<string?> CreateDeckSession()
    {
        var pptxBytes = CreateMinimalPptx();
        using var content = new MultipartFormDataContent();
        var fileContent = new ByteArrayContent(pptxBytes);
        fileContent.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.presentationml.presentation");
        content.Add(fileContent, "file", "test.pptx");

        var response = await _client.PostAsync("/api/decks", content);
        response.EnsureSuccessStatusCode();

        var result = await response.Content.ReadFromJsonAsync<JsonElement>();
        return result.TryGetProperty("deckId", out var id) ? id.GetString() : null;
    }

    private static byte[] CreateMinimalPptx()
    {
        using var builder = PresentationBuilder.Create();
        builder.AddSlide();
        builder.CurrentSlide.AddTitle("Integration Test Slide");
        return builder.SaveToBytes();
    }
}
