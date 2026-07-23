# OfficeEditor.Mcp

A thin [Model Context Protocol](https://modelcontextprotocol.io) (MCP) stdio host that
exposes the OfficeEditor PPTX pipeline to MCP clients (Claude Desktop, IDEs, agents)
as three tools:

| Tool | Purpose | Backed by |
|---|---|---|
| `deck_anatomize` | Slide/element anatomy of a deck (id, type, name, EMU position, text/table data) | `PptxAnatomizer` (via `IPresentationBuilder.Analyze`) |
| `deck_replace_element` | Batch edit ops: `replaceText`, `replaceImage`, `replaceTable`, `moveSlide`, `duplicateSlide`, `deleteSlide` | `PptxInstructionEngine` (validate-then-execute) |
| `deck_render_slide` | Render one slide (1-based) to PNG/SVG bytes | `PresentationBuilder.ExportThumbnail` |

**Zero third-party dependencies.** The JSON-RPC 2.0 / MCP protocol layer is
hand-rolled over `System.Text.Json` (in-box); the only references are
`PptxEditor.Core` and `OfficeEditor.Core`. The official `ModelContextProtocol` SDK
was deliberately not taken to keep the dependency footprint at zero.

## Protocol shape

- **Transport:** stdio. One single-line JSON value per line (MCP stdio framing).
  **stdout carries only protocol frames; all logs go to stderr.**
- **Handshake:** `initialize` → (`notifications/initialized`, tolerated) →
  `tools/list` → `tools/call`. `ping` is also answered.
- **Errors:** domain failures surface as JSON-RPC error objects (`code` + `message`),
  never as stack traces:
  - `-32602` invalid params (bad arguments, malformed base64, out-of-range slide)
  - `-32001` unknown/expired deck handle
  - `-32000` domain failure (unreadable deck, render failure)
  - `-32700` parse error, `-32601` unknown method, `-32603` unexpected internal error
    (details stay on stderr)
- **Serialization:** a single stdin loop executes tool calls one at a time. This is a
  feature: the render pipeline (TypstBridge native library) and OpenXML packages are
  not guaranteed thread-safe, so serialized execution is the safe default.

## Tools

### `deck_anatomize`

```json
{ "pptxBase64": "UEsDBB..." }
```
or
```json
{ "deckHandle": "9a1c98db-d221-4a63-aee1-265474ef068a" }
```

Exactly one of `pptxBase64` / `deckHandle`. An upload **creates a session** and the
response carries its `deckHandle`. Response payload:

```json
{
  "slideCount": 8,
  "slides": [
    { "slideIndex": 1,
      "elements": [
        { "type": "Text", "id": 7, "name": "TextBox 7",
          "location": "slide:1:shape:TextBox 7",
          "text": "Lorem Ipsum ...", "x": 123456, "y": 234567, "cx": 3456789, "cy": 456789 }
      ] }
  ],
  "deckHandle": "9a1c98db-...",
  "revision": 0
}
```

`x`/`y`/`cx`/`cy` are EMU and may be absent (placeholders inherit layout positions).

### `deck_replace_element`

```json
{
  "deckHandle": "9a1c98db-...",
  "operations": [
    { "type": "replaceText", "slide": 1, "elementId": 7, "text": "Edited via MCP" },
    { "type": "moveSlide", "from": 2, "to": 1 }
  ]
}
```

Response payload:

```json
{
  "success": true,
  "appliedOps": 2,
  "changedSlides": [1, 2],
  "errors": [],
  "deckHandle": "9a1c98db-...",
  "revision": 1
}
```

Per-op failures (validation or execution) are reported in `errors`
(`{index, type, error}`) with `success: false` — they are **not** protocol errors.
Stateless mode: pass `pptxBase64` instead of `deckHandle`; the response then includes
`newPptxBase64` with the edited deck bytes (and still creates a session you can
continue from via the returned `deckHandle`).

### `deck_render_slide`

```json
{ "deckHandle": "9a1c98db-...", "slide": 1, "format": "png", "ppi": 150 }
```

- `slide` — 1-based, required. `format` — `png` (default) or `svg`.
  `ppi` — 36–600, default 150.
- Response payload: `{ "contentBase64": "...", "contentType": "image/png",
  "slide": 1, "format": "png", "ppi": 150.0, "deckHandle": "...", "revision": 1 }`

All tool results use the MCP envelope: `content: [{type: "text", text: <json>}]` plus
the same payload under `structuredContent`.

## Session semantics

Mirrors the API's `DeckSessionStore` (upload → handle; 30-minute sliding lifetime;
temp dirs swept on eviction), reimplemented in-process without the ASP.NET project:

- Any tool call with `pptxBase64` uploads the deck and returns a `deckHandle`
  (GUID). Subsequent calls can use the handle — no need to re-upload megabytes.
- **30-minute sliding lifetime:** every access refreshes the expiry. Expired handles
  fail with `-32001`.
- **Revisions:** each successfully applied `deck_replace_element` batch bumps the
  session's monotonic `revision` and stores the new deck bytes (slide count is
  re-derived, since slide ops can change it).
- **Temp hygiene:** each session owns a scratch directory under
  `$TMPDIR/officeeditor-mcp/<pid>/<guid>`. A sweep timer (every 5 min) plus lazy
  sweeps on access delete expired sessions' directories; remaining directories are
  removed on host shutdown. Cleanup is best-effort with the OS temp cleaner as
  backstop.
- Deck bytes live only in memory; nothing is written to disk by the host itself
  (the converter/engine manage their own internal scratch).

## Running

```bash
dotnet run --project OfficeEditor.Mcp
```

(For scripted use, build first and run the DLL directly so `dotnet run` build output
never pollutes stdout — see `scripts/demo.py`.)

### MCP client configuration (stdio)

```json
{
  "mcpServers": {
    "officeeditor": {
      "command": "dotnet",
      "args": ["run", "--project", "/path/to/DocxEditor/OfficeEditor.Mcp"]
    }
  }
}
```

### Demo

```bash
python3 OfficeEditor.Mcp/scripts/demo.py
```

Runs the full workflow against `examples/REF/PPTX/northwind-demo.pptx`:
initialize → tools/list → anatomize → replaceText on slide 1 → render slide 1 to
`/tmp/officeeditor-mcp-slide1.png`.

## Security notes

- **Local stdio only.** The host opens no sockets and performs no network access;
  it is only as exposed as the process that spawns it.
- Deck content arrives as base64 over stdin; uploads are validated as readable PPTX
  before entering a session.
- Protocol frames never contain stack traces; internal exception details go to
  stderr only.
- The host trusts the local caller: it executes arbitrary deck edits and renders on
  demand, which is CPU/memory intensive by nature. Do not expose it behind an
  unsandboxed multi-tenant transport.
