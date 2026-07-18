using System.Text.Json.Nodes;
using OfficeEditor.Mcp.JsonRpc;

namespace OfficeEditor.Mcp.Tests;

/// <summary>
/// Tool-level tests over the real reference deck (no rendering — the render path is
/// exercised by the demo script, keeping unit tests free of the native dependency).
/// Covers the full replaceText round-trip plus argument/error mapping.
/// </summary>
public sealed class DeckToolsTests
{
    [Fact]
    public void Anatomize_Upload_ReturnsSlidesElementsAndCreatesSession()
    {
        using var server = TestHost.CreateServer();

        var response = TestHost.SendToolCall(server, "deck_anatomize", new JsonObject
        {
            ["pptxBase64"] = TestHost.LoadReferenceDeckBase64()
        });

        var payload = TestHost.AssertToolPayload(response);
        Assert.True((int)payload["slideCount"]! > 0);
        Assert.True(Guid.TryParse((string?)payload["deckHandle"], out _));
        Assert.Equal(0, (int)payload["revision"]!);

        var slides = payload["slides"]!.AsArray();
        Assert.Equal((int)payload["slideCount"]!, slides.Count);
        var element = slides.SelectMany(s => s!["elements"]!.AsArray()).First()!;
        Assert.NotNull(element["id"]);
        Assert.NotNull(element["type"]);
        Assert.NotNull(element["name"]);
        Assert.NotNull(element["location"]);
    }

    [Fact]
    public void Anatomize_ByHandle_ReusesUploadedSession()
    {
        using var server = TestHost.CreateServer();
        var handle = Upload(server);

        var response = TestHost.SendToolCall(server, "deck_anatomize", new JsonObject
        {
            ["deckHandle"] = handle
        });

        var payload = TestHost.AssertToolPayload(response);
        Assert.Equal(handle, (string?)payload["deckHandle"]);
        Assert.True((int)payload["slideCount"]! > 0);
    }

    [Fact]
    public void Anatomize_UnknownHandle_ReturnsDeckNotFound()
    {
        using var server = TestHost.CreateServer();

        var response = TestHost.SendToolCall(server, "deck_anatomize", new JsonObject
        {
            ["deckHandle"] = Guid.NewGuid().ToString()
        });

        var (code, _) = TestHost.AssertError(response);
        Assert.Equal(JsonRpcErrorCodes.DeckNotFound, code);
    }

    [Fact]
    public void Anatomize_BothSources_ReturnsInvalidParams()
    {
        using var server = TestHost.CreateServer();

        var response = TestHost.SendToolCall(server, "deck_anatomize", new JsonObject
        {
            ["deckHandle"] = Guid.NewGuid().ToString(),
            ["pptxBase64"] = TestHost.LoadReferenceDeckBase64()
        });

        var (code, message) = TestHost.AssertError(response);
        Assert.Equal(JsonRpcErrorCodes.InvalidParams, code);
        Assert.Contains("Exactly one", message);
    }

    [Fact]
    public void Anatomize_InvalidBase64_ReturnsInvalidParams()
    {
        using var server = TestHost.CreateServer();

        var response = TestHost.SendToolCall(server, "deck_anatomize", new JsonObject
        {
            ["pptxBase64"] = "!!! not base64 !!!"
        });

        var (code, _) = TestHost.AssertError(response);
        Assert.Equal(JsonRpcErrorCodes.InvalidParams, code);
    }

    [Fact]
    public void ReplaceElement_ReplaceText_RoundTripsAndBumpsRevision()
    {
        using var server = TestHost.CreateServer();
        var handle = Upload(server);

        var anatomy = TestHost.AssertToolPayload(TestHost.SendToolCall(
            server, "deck_anatomize", new JsonObject { ["deckHandle"] = handle }));
        var textElement = anatomy["slides"]!.AsArray()
            .SelectMany(s => s!["elements"]!.AsArray())
            .First(e => (string)e!["type"]! == "Text");
        var slideIndex = (int)anatomy["slides"]!.AsArray()
            .First(s => s!["elements"]!.AsArray().Any(e => (string)e!["type"]! == "Text"))!["slideIndex"]!;

        var edit = TestHost.AssertToolPayload(TestHost.SendToolCall(
            server, "deck_replace_element", new JsonObject
            {
                ["deckHandle"] = handle,
                ["operations"] = new JsonArray(new JsonObject
                {
                    ["type"] = "replaceText",
                    ["slide"] = slideIndex,
                    ["elementId"] = (uint)textElement!["id"]!,
                    ["text"] = "Edited by unit test"
                })
            }));

        Assert.True((bool)edit["success"]!);
        Assert.Equal(1, (int)edit["appliedOps"]!);
        Assert.Equal(1, (int)edit["revision"]!);
        Assert.Empty(edit["errors"]!.AsArray());
        Assert.Contains(slideIndex, edit["changedSlides"]!.AsArray().Select(s => (int)s!));

        // The updated deck bytes are kept in the session: re-anatomize sees the text.
        var reAnatomy = TestHost.AssertToolPayload(TestHost.SendToolCall(
            server, "deck_anatomize", new JsonObject { ["deckHandle"] = handle }));
        Assert.Equal(1, (int)reAnatomy["revision"]!);
        var editedText = reAnatomy["slides"]!.AsArray()
            .SelectMany(s => s!["elements"]!.AsArray())
            .Where(e => (string)e!["type"]! == "Text")
            .Select(e => (string?)e!["text"])
            .FirstOrDefault(t => t?.Contains("Edited by unit test") == true);
        Assert.NotNull(editedText);
    }

    [Fact]
    public void ReplaceElement_StatelessUpload_ReturnsNewPptxBase64()
    {
        using var server = TestHost.CreateServer();

        // Upload once to learn a valid element target.
        var handle = Upload(server);
        var anatomy = TestHost.AssertToolPayload(TestHost.SendToolCall(
            server, "deck_anatomize", new JsonObject { ["deckHandle"] = handle }));
        var textElement = anatomy["slides"]!.AsArray()
            .SelectMany(s => s!["elements"]!.AsArray())
            .First(e => (string)e!["type"]! == "Text");

        var edit = TestHost.AssertToolPayload(TestHost.SendToolCall(
            server, "deck_replace_element", new JsonObject
            {
                ["pptxBase64"] = TestHost.LoadReferenceDeckBase64(),
                ["operations"] = new JsonArray(new JsonObject
                {
                    ["type"] = "replaceText",
                    ["slide"] = 1,
                    ["elementId"] = (uint)textElement!["id"]!,
                    ["text"] = "Stateless edit"
                })
            }));

        Assert.True((bool)edit["success"]!);
        var newBytes = Convert.FromBase64String((string)edit["newPptxBase64"]!);
        Assert.True(newBytes.Length > 1000);
        Assert.Equal(new byte[] { (byte)'P', (byte)'K' }, newBytes.Take(2).ToArray());
    }

    [Fact]
    public void ReplaceElement_ValidationFailure_ReturnsSuccessFalseWithErrors()
    {
        using var server = TestHost.CreateServer();
        var handle = Upload(server);

        var edit = TestHost.AssertToolPayload(TestHost.SendToolCall(
            server, "deck_replace_element", new JsonObject
            {
                ["deckHandle"] = handle,
                ["operations"] = new JsonArray(new JsonObject
                {
                    ["type"] = "replaceText",
                    ["slide"] = 1,
                    ["elementId"] = 999999,
                    ["text"] = "nope"
                })
            }));

        Assert.False((bool)edit["success"]!);
        var errors = edit["errors"]!.AsArray();
        Assert.Single(errors);
        Assert.Equal(0, (int)errors[0]!["index"]!);
        Assert.Equal("replaceText", (string?)errors[0]!["type"]);
        // Nothing applied → revision untouched.
        Assert.Equal(0, (int)edit["revision"]!);
    }

    [Fact]
    public void ReplaceElement_UnknownOperationType_ReturnsInvalidParams()
    {
        using var server = TestHost.CreateServer();
        var handle = Upload(server);

        var response = TestHost.SendToolCall(server, "deck_replace_element", new JsonObject
        {
            ["deckHandle"] = handle,
            ["operations"] = new JsonArray(new JsonObject { ["type"] = "nuke" })
        });

        var (code, message) = TestHost.AssertError(response);
        Assert.Equal(JsonRpcErrorCodes.InvalidParams, code);
        Assert.Contains("operations[0]", message);
    }

    [Fact]
    public void ReplaceElement_MissingOperations_ReturnsInvalidParams()
    {
        using var server = TestHost.CreateServer();
        var handle = Upload(server);

        var response = TestHost.SendToolCall(server, "deck_replace_element", new JsonObject
        {
            ["deckHandle"] = handle
        });

        var (code, _) = TestHost.AssertError(response);
        Assert.Equal(JsonRpcErrorCodes.InvalidParams, code);
    }

    [Fact]
    public void RenderSlide_SlideOutOfRange_ReturnsInvalidParams()
    {
        using var server = TestHost.CreateServer();
        var handle = Upload(server);

        var response = TestHost.SendToolCall(server, "deck_render_slide", new JsonObject
        {
            ["deckHandle"] = handle,
            ["slide"] = 999
        });

        var (code, message) = TestHost.AssertError(response);
        Assert.Equal(JsonRpcErrorCodes.InvalidParams, code);
        Assert.Contains("out of range", message);
    }

    [Fact]
    public void RenderSlide_InvalidFormat_ReturnsInvalidParams()
    {
        using var server = TestHost.CreateServer();
        var handle = Upload(server);

        var response = TestHost.SendToolCall(server, "deck_render_slide", new JsonObject
        {
            ["deckHandle"] = handle,
            ["slide"] = 1,
            ["format"] = "gif"
        });

        var (code, _) = TestHost.AssertError(response);
        Assert.Equal(JsonRpcErrorCodes.InvalidParams, code);
    }

    private static string Upload(McpServer server)
    {
        var response = TestHost.SendToolCall(server, "deck_anatomize", new JsonObject
        {
            ["pptxBase64"] = TestHost.LoadReferenceDeckBase64()
        });
        return (string)TestHost.AssertToolPayload(response)["deckHandle"]!;
    }
}
