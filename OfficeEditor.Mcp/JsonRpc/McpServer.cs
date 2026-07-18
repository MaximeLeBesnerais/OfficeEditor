using System.Text.Json;
using System.Text.Json.Nodes;
using OfficeEditor.Mcp.Sessions;
using OfficeEditor.Mcp.Tools;

namespace OfficeEditor.Mcp.JsonRpc;

/// <summary>
/// Minimal MCP (Model Context Protocol) server over newline-delimited JSON-RPC 2.0.
/// <para>
/// Supported surface: <c>initialize</c>, <c>ping</c>, <c>tools/list</c>,
/// <c>tools/call</c>, and notifications (tolerated, no response — including
/// <c>notifications/initialized</c>). Framing is one JSON value per line, per the MCP
/// stdio transport: responses are single-line JSON; log output never touches stdout.
/// </para>
/// <para>
/// Dispatch is synchronous and the host runs a single stdin read loop, so tool calls
/// are naturally serialized — a feature, not a limitation: the rendering pipeline
/// (TypstBridge native library, OpenXML packages) is not guaranteed thread-safe, and
/// one-at-a-time execution keeps every deck mutation and render on a single thread.
/// </para>
/// </summary>
public sealed class McpServer : IDisposable
{
    public const string ProtocolVersion = "2024-11-05";
    public const string ServerName = "officeeditor-mcp";
    public const string ServerVersion = "0.1.0";

    private readonly DeckSessionStore _sessions;
    private readonly DeckTools _tools;
    private readonly TextWriter _log;

    public McpServer(TextWriter? log = null, DeckSessionStore? sessions = null)
    {
        _log = log ?? TextWriter.Null;
        _sessions = sessions ?? new DeckSessionStore();
        _tools = new DeckTools(_sessions);
    }

    /// <summary>Handles one input line; returns the response line, or null for notifications.</summary>
    public string? HandleLine(string line)
    {
        JsonDocument document;
        try
        {
            document = JsonDocument.Parse(line);
        }
        catch (JsonException ex)
        {
            _log.WriteLine($"[mcp] parse error: {ex.Message}");
            return JsonRpcFrame.Error(null, JsonRpcErrorCodes.ParseError, "Parse error: invalid JSON.").ToJsonString();
        }

        using (document)
        {
            return HandleMessage(document.RootElement)?.ToJsonString();
        }
    }

    private JsonObject? HandleMessage(JsonElement root)
    {
        if (root.ValueKind != JsonValueKind.Object)
        {
            return JsonRpcFrame.Error(null, JsonRpcErrorCodes.InvalidRequest,
                "Invalid request: expected a JSON-RPC object.");
        }

        JsonNode? id = null;
        var hasId = root.TryGetProperty("id", out var idElement)
                    && idElement.ValueKind is not (JsonValueKind.Null or JsonValueKind.Undefined);
        if (hasId)
        {
            id = JsonRpcFrame.IdToNode(idElement);
        }

        if (!root.TryGetProperty("method", out var methodElement)
            || methodElement.ValueKind != JsonValueKind.String)
        {
            return JsonRpcFrame.Error(id, JsonRpcErrorCodes.InvalidRequest,
                "Invalid request: missing 'method'.");
        }
        var method = methodElement.GetString()!;

        // Notifications (no id) are tolerated and never answered.
        if (!hasId)
        {
            if (!method.StartsWith("notifications/", StringComparison.Ordinal))
            {
                _log.WriteLine($"[mcp] ignoring notification '{method}'.");
            }
            return null;
        }

        JsonElement? parameters = root.TryGetProperty("params", out var paramsElement)
            ? paramsElement
            : null;

        try
        {
            return method switch
            {
                "initialize" => JsonRpcFrame.Result(id, Initialize()),
                "ping" => JsonRpcFrame.Result(id, new JsonObject()),
                "tools/list" => JsonRpcFrame.Result(id, new JsonObject { ["tools"] = _tools.ListTools() }),
                "tools/call" => JsonRpcFrame.Result(id, CallTool(parameters)),
                _ => JsonRpcFrame.Error(id, JsonRpcErrorCodes.MethodNotFound,
                    $"Method not found: '{method}'.")
            };
        }
        catch (McpException ex)
        {
            _log.WriteLine($"[mcp] {method} failed ({ex.Code}): {ex.Message}");
            return JsonRpcFrame.Error(id, ex.Code, ex.Message);
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException
            or InvalidDataException or IOException or KeyNotFoundException or NotSupportedException)
        {
            // Domain failures: message is client-safe, stack trace stays on stderr.
            _log.WriteLine($"[mcp] {method} domain failure: {ex}");
            return JsonRpcFrame.Error(id, JsonRpcErrorCodes.DomainError, ex.Message);
        }
        catch (Exception ex)
        {
            // Unexpected: never leak internals into protocol frames.
            _log.WriteLine($"[mcp] {method} internal error: {ex}");
            return JsonRpcFrame.Error(id, JsonRpcErrorCodes.InternalError,
                "Internal error while executing the request; see server stderr for details.");
        }
    }

    private static JsonObject Initialize()
    {
        return new JsonObject
        {
            ["protocolVersion"] = ProtocolVersion,
            ["capabilities"] = new JsonObject
            {
                ["tools"] = new JsonObject()
            },
            ["serverInfo"] = new JsonObject
            {
                ["name"] = ServerName,
                ["version"] = ServerVersion
            }
        };
    }

    private JsonObject CallTool(JsonElement? parameters)
    {
        if (parameters is null || parameters.Value.ValueKind != JsonValueKind.Object)
        {
            throw new McpException(JsonRpcErrorCodes.InvalidParams,
                "tools/call requires a 'params' object with 'name' and 'arguments'.");
        }

        if (!parameters.Value.TryGetProperty("name", out var nameElement)
            || nameElement.ValueKind != JsonValueKind.String)
        {
            throw new McpException(JsonRpcErrorCodes.InvalidParams,
                "tools/call requires a 'name' string.");
        }

        JsonElement? arguments = parameters.Value.TryGetProperty("arguments", out var argsElement)
            ? argsElement
            : null;

        var payload = _tools.Call(nameElement.GetString()!, arguments);

        // MCP tool-result envelope: structured payload plus a text content part for
        // clients that only render content blocks.
        return new JsonObject
        {
            ["content"] = new JsonArray(new JsonObject
            {
                ["type"] = "text",
                ["text"] = payload.ToJsonString()
            }),
            ["structuredContent"] = payload,
            ["isError"] = false
        };
    }

    public void Dispose() => _sessions.Dispose();
}
