using System.Text.Json;
using System.Text.Json.Nodes;

namespace OfficeEditor.Mcp.JsonRpc;

/// <summary>JSON-RPC 2.0 error codes (standard plus the MCP server-error range).</summary>
public static class JsonRpcErrorCodes
{
    public const int ParseError = -32700;
    public const int InvalidRequest = -32600;
    public const int MethodNotFound = -32601;
    public const int InvalidParams = -32602;
    public const int InternalError = -32603;

    /// <summary>Domain failure surfaced by a tool (invalid deck, render failure, ...).</summary>
    public const int DomainError = -32000;

    /// <summary>Unknown or expired deck handle.</summary>
    public const int DeckNotFound = -32001;
}

/// <summary>
/// A protocol-level failure. The host maps <see cref="Code"/> and <see cref="Exception.Message"/>
/// onto a JSON-RPC error object. Messages are written to be client-safe — they never
/// contain stack traces or internal paths beyond what the caller supplied.
/// </summary>
public sealed class McpException : Exception
{
    public int Code { get; }

    public McpException(int code, string message)
        : base(message)
    {
        Code = code;
    }
}

/// <summary>Helpers for building JSON-RPC 2.0 frames over System.Text.Json nodes.</summary>
public static class JsonRpcFrame
{
    /// <summary>Wraps a JsonElement id into a JsonNode (ids may be string, number, or null).</summary>
    public static JsonNode? IdToNode(JsonElement? id)
    {
        if (id is null || id.Value.ValueKind is JsonValueKind.Undefined or JsonValueKind.Null)
        {
            return null;
        }
        return JsonNode.Parse(id.Value.GetRawText());
    }

    public static JsonObject Result(JsonNode? id, JsonNode result)
    {
        return new JsonObject
        {
            ["jsonrpc"] = "2.0",
            ["id"] = id,
            ["result"] = result
        };
    }

    public static JsonObject Error(JsonNode? id, int code, string message)
    {
        return new JsonObject
        {
            ["jsonrpc"] = "2.0",
            ["id"] = id,
            ["error"] = new JsonObject
            {
                ["code"] = code,
                ["message"] = message
            }
        };
    }
}
