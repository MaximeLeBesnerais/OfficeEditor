using System.Text.Json;
using System.Text.Json.Nodes;
using OfficeEditor.Mcp.JsonRpc;
using OfficeEditor.Mcp.Sessions;

namespace OfficeEditor.Mcp.Tests;

/// <summary>
/// Shared helpers: build a disposable server over a timer-less session store,
/// send request frames through <see cref="McpServer.HandleLine"/>, and locate the
/// reference deck. No stdio, no environment mutation.
/// </summary>
internal static class TestHost
{
    public static McpServer CreateServer() =>
        new(sessions: new DeckSessionStore(enableSweepTimer: false));

    public static JsonObject Send(McpServer server, string method, JsonObject? parameters = null, int id = 1)
    {
        var frame = new JsonObject
        {
            ["jsonrpc"] = "2.0",
            ["id"] = id,
            ["method"] = method
        };
        if (parameters is not null)
        {
            frame["params"] = parameters;
        }

        var response = server.HandleLine(frame.ToJsonString());
        Assert.NotNull(response);
        var parsed = JsonNode.Parse(response)!.AsObject();
        Assert.Equal("2.0", (string?)parsed["jsonrpc"]);
        Assert.Equal(id, (int?)parsed["id"]);
        return parsed;
    }

    public static JsonObject SendToolCall(McpServer server, string toolName, JsonObject arguments)
    {
        return Send(server, "tools/call", new JsonObject
        {
            ["name"] = toolName,
            ["arguments"] = arguments
        });
    }

    /// <summary>Asserts the frame is a JSON-RPC error and returns (code, message).</summary>
    public static (int Code, string Message) AssertError(JsonObject response)
    {
        var error = response["error"]!.AsObject();
        return ((int)error["code"]!, (string)error["message"]!);
    }

    /// <summary>Extracts the structuredContent payload of a successful tools/call result.</summary>
    public static JsonObject AssertToolPayload(JsonObject response)
    {
        var result = response["result"]!.AsObject();
        Assert.False((bool)result["isError"]!);
        return result["structuredContent"]!.AsObject();
    }

    public static byte[] LoadReferenceDeck()
    {
        // Same resolution pattern as DocxEditor.Tests: bin/Debug/net9.0 -> repo root.
        var path = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory, "..", "..", "..", "..",
            "examples", "REF", "PPTX", "northwind-demo.pptx"));
        Assert.True(File.Exists(path), $"Reference deck not found at {path}");
        return File.ReadAllBytes(path);
    }

    public static string LoadReferenceDeckBase64() => Convert.ToBase64String(LoadReferenceDeck());
}
