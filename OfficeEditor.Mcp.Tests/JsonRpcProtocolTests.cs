using System.Text.Json.Nodes;
using OfficeEditor.Mcp.JsonRpc;

namespace OfficeEditor.Mcp.Tests;

/// <summary>
/// Protocol-layer tests: JSON-RPC framing, the initialize handshake, notification
/// tolerance, tools/list shape, and error-code mapping. No deck content needed.
/// </summary>
public sealed class JsonRpcProtocolTests
{
    [Fact]
    public void InvalidJson_ReturnsParseError()
    {
        using var server = TestHost.CreateServer();

        var response = server.HandleLine("this is not json");

        var parsed = JsonNode.Parse(response!)!.AsObject();
        Assert.Equal(JsonRpcErrorCodes.ParseError, (int)parsed["error"]!["code"]!);
        Assert.Null(parsed["id"]);
    }

    [Fact]
    public void NonObjectFrame_ReturnsInvalidRequest()
    {
        using var server = TestHost.CreateServer();

        var response = server.HandleLine("[1,2,3]");

        var parsed = JsonNode.Parse(response!)!.AsObject();
        Assert.Equal(JsonRpcErrorCodes.InvalidRequest, (int)parsed["error"]!["code"]!);
    }

    [Fact]
    public void Initialize_ReturnsProtocolVersionCapabilitiesAndServerInfo()
    {
        using var server = TestHost.CreateServer();

        var response = TestHost.Send(server, "initialize", new JsonObject
        {
            ["protocolVersion"] = "2024-11-05",
            ["capabilities"] = new JsonObject(),
            ["clientInfo"] = new JsonObject { ["name"] = "test", ["version"] = "1.0" }
        });

        var result = response["result"]!.AsObject();
        Assert.Equal(McpServer.ProtocolVersion, (string?)result["protocolVersion"]);
        Assert.NotNull(result["capabilities"]!["tools"]);
        Assert.Equal("officeeditor-mcp", (string?)result["serverInfo"]!["name"]);
        Assert.Null(response["error"]);
    }

    [Fact]
    public void Notification_GetsNoResponse()
    {
        using var server = TestHost.CreateServer();

        var response = server.HandleLine(
            """{"jsonrpc":"2.0","method":"notifications/initialized"}""");

        Assert.Null(response);
    }

    [Fact]
    public void UnknownMethod_ReturnsMethodNotFound()
    {
        using var server = TestHost.CreateServer();

        var response = TestHost.Send(server, "resources/list");

        var (code, _) = TestHost.AssertError(response);
        Assert.Equal(JsonRpcErrorCodes.MethodNotFound, code);
    }

    [Fact]
    public void Ping_ReturnsEmptyResult()
    {
        using var server = TestHost.CreateServer();

        var response = TestHost.Send(server, "ping");

        Assert.NotNull(response["result"]);
        Assert.Null(response["error"]);
    }

    [Fact]
    public void ToolsList_ReturnsFourToolsWithSchemas()
    {
        using var server = TestHost.CreateServer();

        var response = TestHost.Send(server, "tools/list");

        var tools = response["result"]!["tools"]!.AsArray();
        Assert.Equal(4, tools.Count);
        var names = tools.Select(t => (string)t!["name"]!).ToList();
        Assert.Equal(
            new[] { "deck_anatomize", "deck_replace_element", "deck_render_slide", "deck_generate" },
            names);
        foreach (var tool in tools)
        {
            Assert.NotNull(tool!["description"]);
            Assert.Equal("object", (string?)tool["inputSchema"]!["type"]);
        }
    }

    [Fact]
    public void ToolsCall_WithoutParams_ReturnsInvalidParams()
    {
        using var server = TestHost.CreateServer();

        var response = TestHost.Send(server, "tools/call");

        var (code, _) = TestHost.AssertError(response);
        Assert.Equal(JsonRpcErrorCodes.InvalidParams, code);
    }

    [Fact]
    public void ToolsCall_UnknownTool_ReturnsInvalidParams()
    {
        using var server = TestHost.CreateServer();

        var response = TestHost.SendToolCall(server, "deck_explode", new JsonObject());

        var (code, message) = TestHost.AssertError(response);
        Assert.Equal(JsonRpcErrorCodes.InvalidParams, code);
        Assert.Contains("deck_explode", message);
    }

    [Fact]
    public void ToolsCall_WithoutArguments_ReturnsInvalidParams()
    {
        using var server = TestHost.CreateServer();

        var response = TestHost.Send(server, "tools/call", new JsonObject
        {
            ["name"] = "deck_anatomize"
        });

        var (code, _) = TestHost.AssertError(response);
        Assert.Equal(JsonRpcErrorCodes.InvalidParams, code);
    }

    [Fact]
    public void ErrorMessages_DoNotContainStackTraces()
    {
        using var server = TestHost.CreateServer();

        // Force a domain error (unknown deck handle) and inspect the frame text.
        var response = TestHost.SendToolCall(server, "deck_anatomize", new JsonObject
        {
            ["deckHandle"] = Guid.NewGuid().ToString()
        });

        var frame = response.ToJsonString();
        Assert.DoesNotContain(" at ", frame);
        Assert.DoesNotContain("StackTrace", frame);
        Assert.DoesNotContain(".cs:line", frame);
    }
}
