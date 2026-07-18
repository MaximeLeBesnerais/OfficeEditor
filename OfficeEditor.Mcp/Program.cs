using OfficeEditor.Mcp.JsonRpc;

// OfficeEditor.Mcp — MCP stdio host.
// stdout carries ONLY JSON-RPC protocol frames (one single-line JSON value per line);
// all diagnostics go to stderr, so stdout can be piped straight into an MCP client.
var log = Console.Error;
log.WriteLine($"[mcp] {McpServer.ServerName} {McpServer.ServerVersion} ready (protocol {McpServer.ProtocolVersion}).");

using var server = new McpServer(log);

// Single stdin loop: tool calls execute synchronously one at a time. This is
// deliberate — the render pipeline (TypstBridge native library) and OpenXML
// packages are not guaranteed thread-safe, so serialization is a safety feature.
string? line;
while ((line = Console.In.ReadLine()) is not null)
{
    if (string.IsNullOrWhiteSpace(line))
    {
        continue;
    }

    try
    {
        var response = server.HandleLine(line);
        if (response is not null)
        {
            Console.Out.WriteLine(response);
            Console.Out.Flush();
        }
    }
    catch (Exception ex)
    {
        // The loop must survive any single-frame failure.
        log.WriteLine($"[mcp] unhandled error while processing a frame: {ex}");
    }
}

log.WriteLine("[mcp] stdin closed; shutting down.");
