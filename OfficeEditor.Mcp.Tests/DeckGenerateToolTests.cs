using System.Text.Json.Nodes;
using OfficeEditor.Mcp.JsonRpc;
using OfficeEditor.Mcp.Tools;

namespace OfficeEditor.Mcp.Tests;

/// <summary>
/// deck_generate (P9, plan.md §7.1): generation JSON in → PPTX + per-slide previews out,
/// invalid documents rejected with the P1 validator's errors verbatim. Preview rendering
/// needs a Typst backend (absent in this sandbox), so render assertions are
/// environment-invariant (previews xor previewError) with an opt-in happy-path check via
/// OE_RUN_TYPST_COMPILE_TESTS=1 (same convention as TypstEmitterCompileTests).
/// </summary>
public sealed class DeckGenerateToolTests
{
    private const string EnableRenderEnvVar = "OE_RUN_TYPST_COMPILE_TESTS";

    private const string ValidDocument = """
        {
          "version": "2.0",
          "design": {
            "palette": { "primary": "#0B3D91", "accent": "#FF6B00", "ink": "#1A1A1A", "paper": "#FFFFFF", "muted": "#8A94A6" },
            "fonts": { "display": "Aptos Display", "body": "Aptos" },
            "metrics": { "marginPt": 43, "gutterPt": 18, "titleSizePt": 30, "bodySizePt": 14 }
          },
          "slides": [
            {
              "type": "container",
              "layout": { "mode": "row", "gap": 18, "justify": "space-evenly", "align": "center" },
              "padding": 43,
              "fill": "paper",
              "children": [
                { "type": "card", "size": { "grow": 1, "aspect": "4:3" },
                  "content": { "title": "+34%", "subtitle": "Revenue" } },
                { "type": "card", "size": { "grow": 1, "aspect": "4:3" },
                  "content": { "title": "12k", "subtitle": "Users" } }
              ]
            },
            {
              "type": "container",
              "fill": "primary",
              "children": [
                { "type": "text", "text": "Slide two", "color": "paper", "fontSize": 30,
                  "anchor": "middle", "textAlign": "center",
                  "at": { "x": 0, "y": 250 }, "size": { "w": 960, "h": 40 } }
              ]
            }
          ]
        }
        """;

    private static JsonObject GenerateArgs(string? document = null) => new()
    {
        ["document"] = JsonNode.Parse(document ?? ValidDocument)
    };

    [Fact]
    public void ListTools_IncludesDeckGenerate()
    {
        using var server = TestHost.CreateServer();

        var response = TestHost.Send(server, "tools/list");

        var tools = response["result"]!["tools"]!.AsArray();
        var generate = tools.Single(t => (string)t!["name"]! == ToolSchemas.GenerateName);
        var schema = generate!["inputSchema"]!;
        Assert.Contains("document", schema["required"]!.AsArray().Select(r => (string)r!));
        Assert.NotNull(schema["properties"]!["document"]);
        Assert.NotNull(schema["properties"]!["previewFormat"]);
        Assert.NotNull(schema["properties"]!["ppi"]);
    }

    [Fact]
    public void Generate_ValidDocument_ReturnsPptxAndSession()
    {
        using var server = TestHost.CreateServer();

        var response = TestHost.SendToolCall(server, "deck_generate", GenerateArgs());

        var payload = TestHost.AssertToolPayload(response);
        Assert.True((bool)payload["success"]!);
        Assert.Equal(2, (int)payload["slideCount"]!);
        Assert.True(Guid.TryParse((string?)payload["deckHandle"], out _));
        Assert.Equal(0, (int)payload["revision"]!);
        Assert.Equal("svg", (string?)payload["previewFormat"]);
        Assert.True((double)payload["generationMs"]! >= 0);

        var pptx = Convert.FromBase64String((string)payload["pptxBase64"]!);
        Assert.True(pptx.Length > 0);
        Assert.Equal((byte)'P', pptx[0]); // zip magic: a real package
        Assert.Equal((byte)'K', pptx[1]);
    }

    [Fact]
    public void Generate_SessionHandle_WorksWithOtherTools()
    {
        using var server = TestHost.CreateServer();

        var generated = TestHost.AssertToolPayload(
            TestHost.SendToolCall(server, "deck_generate", GenerateArgs()));
        var handle = (string)generated["deckHandle"]!;

        var anatomy = TestHost.AssertToolPayload(TestHost.SendToolCall(
            server, "deck_anatomize", new JsonObject { ["deckHandle"] = handle }));
        Assert.Equal(2, (int)anatomy["slideCount"]!);
    }

    [Fact]
    public void Generate_PreviewContractIsInvariant()
    {
        using var server = TestHost.CreateServer();

        var payload = TestHost.AssertToolPayload(
            TestHost.SendToolCall(server, "deck_generate", GenerateArgs()));

        var previews = payload["previews"]!.AsArray();
        var previewError = (string?)payload["previewError"];
        if (previewError is null)
        {
            Assert.Equal(2, previews.Count);
            Assert.All(previews, p =>
            {
                Assert.Equal("image/svg+xml", (string)p!["contentType"]!);
                Assert.NotEmpty(Convert.FromBase64String((string)p["contentBase64"]!));
            });
        }
        else
        {
            // No Typst backend (sandbox): explicit error, empty previews, deck still delivered.
            Assert.Empty(previews);
            Assert.NotEmpty(previewError);
        }
    }

    [Fact]
    public void Generate_InvalidDocument_ReturnsValidatorErrorsVerbatim()
    {
        var document = ValidDocument.Replace("\"text\": \"Slide two\"", "\"text\": \"Slide two\", \"colour\": \"paper\"");
        using var server = TestHost.CreateServer();

        var response = TestHost.SendToolCall(server, "deck_generate", GenerateArgs(document));

        var (code, message) = TestHost.AssertError(response);
        Assert.Equal(JsonRpcErrorCodes.InvalidParams, code);
        Assert.Contains("$.slides[1].children[0].colour", message);
        Assert.Contains("unknown property 'colour'", message);
        Assert.Contains("Did you mean 'color'?", message);
    }

    [Fact]
    public void Generate_CssIsm_ReturnsActionableError()
    {
        using var server = TestHost.CreateServer();

        var response = TestHost.SendToolCall(server, "deck_generate", GenerateArgs("""
            {
              "version": "2.0",
              "design": { "palette": { "primary": "#0B3D91" } },
              "slides": [ { "type": "container", "zIndex": 3, "children": [] } ]
            }
            """));

        var (code, message) = TestHost.AssertError(response);
        Assert.Equal(JsonRpcErrorCodes.InvalidParams, code);
        Assert.Contains("CSS 'z-index' is not supported", message);
    }

    [Fact]
    public void Generate_LayoutOverflow_ReturnsErrorWithPath()
    {
        using var server = TestHost.CreateServer();

        var response = TestHost.SendToolCall(server, "deck_generate", GenerateArgs("""
            {
              "version": "2.0",
              "design": { "palette": { "primary": "#0B3D91" } },
              "slides": [
                {
                  "type": "container",
                  "layout": { "mode": "row" },
                  "children": [
                    { "type": "rect", "size": { "w": 600, "h": 100 }, "fill": "primary" },
                    { "type": "rect", "size": { "w": 600, "h": 100 }, "fill": "primary" }
                  ]
                }
              ]
            }
            """));

        var (code, message) = TestHost.AssertError(response);
        Assert.Equal(JsonRpcErrorCodes.InvalidParams, code);
        Assert.Contains("slides[0]", message);
        Assert.Contains("overflow", message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Generate_MissingDocument_ReturnsInvalidParams()
    {
        using var server = TestHost.CreateServer();

        var response = TestHost.SendToolCall(server, "deck_generate", new JsonObject());

        var (code, message) = TestHost.AssertError(response);
        Assert.Equal(JsonRpcErrorCodes.InvalidParams, code);
        Assert.Contains("'document' is required", message);
    }

    [Fact]
    public void Generate_InvalidPreviewFormat_ReturnsInvalidParams()
    {
        using var server = TestHost.CreateServer();
        var args = GenerateArgs();
        args["previewFormat"] = "jpeg";

        var response = TestHost.SendToolCall(server, "deck_generate", args);

        var (code, message) = TestHost.AssertError(response);
        Assert.Equal(JsonRpcErrorCodes.InvalidParams, code);
        Assert.Contains("previewFormat", message);
    }

    [Fact]
    public void Generate_PpiOutOfRange_ReturnsInvalidParams()
    {
        using var server = TestHost.CreateServer();
        var args = GenerateArgs();
        args["ppi"] = 1200;

        var response = TestHost.SendToolCall(server, "deck_generate", args);

        var (code, message) = TestHost.AssertError(response);
        Assert.Equal(JsonRpcErrorCodes.InvalidParams, code);
        Assert.Contains("ppi", message);
    }

    [Fact]
    public void Generate_Previews_RenderedPerSlide_WhenBackendAvailable()
    {
        if (Environment.GetEnvironmentVariable(EnableRenderEnvVar) != "1")
        {
            return; // no Typst backend in this environment (see class summary)
        }

        using var server = TestHost.CreateServer();
        var payload = TestHost.AssertToolPayload(
            TestHost.SendToolCall(server, "deck_generate", GenerateArgs()));

        Assert.Null((string?)payload["previewError"]);
        var previews = payload["previews"]!.AsArray();
        Assert.Equal(2, previews.Count);
        Assert.All(previews, p =>
            Assert.Contains("<svg", System.Text.Encoding.UTF8.GetString(
                Convert.FromBase64String((string)p!["contentBase64"]!))));
    }
}
