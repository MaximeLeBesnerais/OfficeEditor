using System.Text.Json.Nodes;

namespace OfficeEditor.Mcp.Tools;

/// <summary>
/// Static MCP tool definitions (name, description, JSON Schema for the arguments).
/// The schemas mirror the argument validation performed in <see cref="DeckTools"/>.
/// </summary>
public static class ToolSchemas
{
    public const string AnatomizeName = "deck_anatomize";
    public const string ReplaceElementName = "deck_replace_element";
    public const string RenderSlideName = "deck_render_slide";
    public const string GenerateName = "deck_generate";

    private const string DeckSourceProperties = """
        "pptxBase64": {
          "type": "string",
          "description": "Base64-encoded .pptx bytes. Supplying an upload creates a deck session and returns its deckHandle."
        },
        "deckHandle": {
          "type": "string",
          "description": "Handle (GUID) of a previously uploaded deck session."
        }
        """;

    private static readonly JsonObject Anatomize = new()
    {
        ["name"] = AnatomizeName,
        ["description"] = "Analyzes a PPTX deck: returns every slide with its elements (id, type, name, EMU position/extent, text or table data). Exactly one of pptxBase64 or deckHandle is required.",
        ["inputSchema"] = JsonNode.Parse($$"""
        {
          "type": "object",
          "properties": {
            {{DeckSourceProperties}}
          }
        }
        """)
    };

    private static readonly JsonObject ReplaceElement = new()
    {
        ["name"] = ReplaceElementName,
        ["description"] = "Applies a batch of edit operations (replaceText, replaceImage, replaceTable, moveSlide, duplicateSlide, deleteSlide) to a deck. Stateful: pass deckHandle — the session keeps the updated bytes. Stateless: pass pptxBase64 — the response includes newPptxBase64.",
        ["inputSchema"] = JsonNode.Parse($$"""
        {
          "type": "object",
          "properties": {
            {{DeckSourceProperties}},
            "operations": {
              "type": "array",
              "description": "Edit operations, applied in order against the pre-batch deck (validate-then-execute).",
              "items": {
                "type": "object",
                "properties": {
                  "type": { "type": "string", "enum": ["replaceText", "replaceImage", "replaceTable", "moveSlide", "duplicateSlide", "deleteSlide"] },
                  "slide": { "type": "integer", "minimum": 1, "description": "1-based slide index (replaceText/replaceImage/replaceTable/duplicateSlide/deleteSlide)." },
                  "elementId": { "type": "integer", "minimum": 1, "description": "cNvPr element id from deck_anatomize (replace* ops)." },
                  "text": { "type": "string", "description": "Replacement text (replaceText)." },
                  "image": { "type": "string", "description": "Base64 image bytes (replaceImage)." },
                  "fit": { "type": "string", "enum": ["stretch", "fill", "crop", "contain"], "description": "Image fit mode (replaceImage, forward-compatible)." },
                  "rows": { "type": "array", "items": { "type": "array", "items": { "type": "string" } }, "description": "Table cell data (replaceTable)." },
                  "from": { "type": "integer", "minimum": 1, "description": "1-based source slot (moveSlide)." },
                  "to": { "type": "integer", "minimum": 1, "description": "1-based destination slot (moveSlide)." },
                  "position": { "type": "integer", "minimum": 1, "description": "1-based insert position (duplicateSlide, optional)." }
                },
                "required": ["type"]
              }
            }
          },
          "required": ["operations"]
        }
        """)
    };

    private static readonly JsonObject RenderSlide = new()
    {
        ["name"] = RenderSlideName,
        ["description"] = "Renders one slide (1-based) of a deck to PNG or SVG via the Typst pipeline and returns the image bytes as base64.",
        ["inputSchema"] = JsonNode.Parse($$"""
        {
          "type": "object",
          "properties": {
            {{DeckSourceProperties}},
            "slide": { "type": "integer", "minimum": 1, "description": "1-based slide index." },
            "format": { "type": "string", "enum": ["png", "svg"], "default": "png" },
            "ppi": { "type": "number", "minimum": 36, "maximum": 600, "default": 150 }
          },
          "required": ["slide"]
        }
        """)
    };

    private static readonly JsonObject Generate = new()
    {
        ["name"] = GenerateName,
        ["description"] = "Generates a PPTX deck from scratch out of a generation document (deck.schema.json v2.0: design tokens + container tree; layout resolved server-side). Returns the deck as pptxBase64 plus per-slide SVG/PNG previews rendered through the Typst pipeline (best-effort: previewError is set when no Typst backend is available). Invalid documents are rejected with the validator's actionable errors. A deck session is created; the returned deckHandle works with every other deck_* tool.",
        ["inputSchema"] = JsonNode.Parse("""
        {
          "type": "object",
          "properties": {
            "document": {
              "type": "object",
              "description": "Generation document per deck.schema.json v2.0 (version/design/slides; pt units only). Unknown properties are rejected with path + suggestion."
            },
            "previewFormat": { "type": "string", "enum": ["svg", "png"], "default": "svg" },
            "ppi": { "type": "number", "minimum": 36, "maximum": 600, "default": 150 }
          },
          "required": ["document"]
        }
        """)
    };

    public static JsonArray ListAll()
    {
        return new JsonArray(
            Anatomize.DeepClone(),
            ReplaceElement.DeepClone(),
            RenderSlide.DeepClone(),
            Generate.DeepClone());
    }
}
