using System.Text.Json;
using System.Text.Json.Serialization;
using DocxEditor.Core.Models;

namespace DocxEditor.Core.Serialization;

public class DocxJsonInstructionParser
{
    private readonly JsonSerializerOptions _options;

    public DocxJsonInstructionParser()
    {
        _options = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
        };
    }

    public DocumentInstructions Parse(string json)
    {
        var wrapper = JsonSerializer.Deserialize<JsonInstructionWrapper>(json, _options);
        if (wrapper?.Operations == null)
        {
            throw new ArgumentException("Invalid JSON instruction file.");
        }

        var instructions = new List<Instruction>();
        foreach (var op in wrapper.Operations)
        {
            instructions.Add(ParseInstruction(op));
        }

        return new DocumentInstructions { Operations = instructions };
    }

    private Instruction ParseInstruction(JsonInstructionDto dto)
    {
        return dto.Type?.ToLowerInvariant() switch
        {
            "create" => new CreateDocumentInstruction(),
            "addparagraph" => new AddParagraphInstruction
            {
                Text = dto.Text ?? throw new ArgumentException("Text is required for addParagraph."),
                Style = dto.Style
            },
            "replacetext" => new ReplaceTextInstruction
            {
                Find = dto.Find ?? throw new ArgumentException("Find is required for replaceText."),
                Replace = dto.Replace ?? throw new ArgumentException("Replace is required for replaceText.")
            },
            "insertafter" => new InsertAfterInstruction
            {
                Target = dto.Target ?? throw new ArgumentException("Target is required for insertAfter."),
                Content = dto.Content ?? throw new ArgumentException("Content is required for insertAfter.")
            },
            "addrichcontent" => new AddRichContentInstruction
            {
                Blocks = ParseContentBlocks(dto.Blocks ?? throw new ArgumentException("Blocks is required for addRichContent."))
            },
            "replacewithrichcontent" => new ReplaceWithRichContentInstruction
            {
                Target = dto.Target ?? throw new ArgumentException("Target is required for replaceWithRichContent."),
                Blocks = ParseContentBlocks(dto.Blocks ?? throw new ArgumentException("Blocks is required for replaceWithRichContent."))
            },
            _ => throw new NotSupportedException($"Instruction type '{dto.Type}' is not supported.")
        };
    }

    internal static List<ContentBlock> ParseContentBlocks(JsonElement blocksElement)
    {
        if (blocksElement.ValueKind != JsonValueKind.Array)
        {
            throw new ArgumentException("Blocks must be a JSON array.");
        }

        var blocks = new List<ContentBlock>();
        foreach (var blockElement in blocksElement.EnumerateArray())
        {
            if (blockElement.ValueKind != JsonValueKind.Object)
            {
                throw new ArgumentException("Each block must be a JSON object.");
            }

            if (!blockElement.TryGetProperty("type", out var typeProp)
                || typeProp.ValueKind != JsonValueKind.String
                || string.IsNullOrEmpty(typeProp.GetString()))
            {
                throw new ArgumentException("Each block must have a 'type' field.");
            }

            var type = typeProp.GetString()!;

            blocks.Add(ParseBlock(type, blockElement));
        }

        return blocks;
    }

    private static ContentBlock ParseBlock(string type, JsonElement element)
    {
        return type.ToLowerInvariant() switch
        {
            "paragraph" => new ParagraphBlock
            {
                Text = element.TryGetProperty("text", out var text) ? text.GetString() ?? string.Empty : string.Empty,
                Style = element.TryGetProperty("style", out var style) ? style.GetString() : null,
                InlineFormats = element.TryGetProperty("inlineFormats", out var formats) ? ParseInlineFormats(formats) : null
            },
            "heading" => new HeadingBlock
            {
                Level = element.TryGetProperty("level", out var level) && level.ValueKind == JsonValueKind.Number ? level.GetInt32() : 1,
                Text = element.TryGetProperty("text", out var hText) ? hText.GetString() ?? string.Empty : string.Empty,
                Style = element.TryGetProperty("style", out var hStyle) ? hStyle.GetString() : null
            },
            "list" => new ListBlock
            {
                Ordered = element.TryGetProperty("ordered", out var ordered) && ordered.ValueKind == JsonValueKind.True,
                Items = element.TryGetProperty("items", out var items) ? items.EnumerateArray().Select(i => i.GetString() ?? string.Empty).ToList() : [],
                Style = element.TryGetProperty("style", out var lStyle) ? lStyle.GetString() : null
            },
            "table" => new TableBlock
            {
                Rows = element.TryGetProperty("rows", out var rows) ? ParseTableRows(rows) : []
            },
            "blockquote" => new BlockquoteBlock
            {
                Text = element.TryGetProperty("text", out var bqText) ? bqText.GetString() ?? string.Empty : string.Empty,
                Style = element.TryGetProperty("style", out var bqStyle) ? bqStyle.GetString() : null
            },
            "code" => new CodeBlock
            {
                Text = element.TryGetProperty("text", out var cText) ? cText.GetString() ?? string.Empty : string.Empty,
                Language = element.TryGetProperty("language", out var lang) ? lang.GetString() : null,
                Style = element.TryGetProperty("style", out var cStyle) ? cStyle.GetString() : null
            },
            "horizontalrule" => new HorizontalRuleBlock(),
            "custom" => new CustomBlock
            {
                CustomType = element.TryGetProperty("customType", out var ct) ? ct.GetString() ?? string.Empty : string.Empty,
                Text = element.TryGetProperty("text", out var cuText) ? cuText.GetString() ?? string.Empty : string.Empty,
                Style = element.TryGetProperty("style", out var cuStyle) ? cuStyle.GetString() : null
            },
            _ => throw new NotSupportedException($"Block type '{type}' is not supported. Valid types: paragraph, heading, list, table, blockquote, code, horizontalRule, custom.")
        };
    }

    private static List<InlineFormat>? ParseInlineFormats(JsonElement formats)
    {
        if (formats.ValueKind != JsonValueKind.Array || formats.GetArrayLength() == 0)
            return null;

        var result = new List<InlineFormat>();
        foreach (var f in formats.EnumerateArray())
        {
            result.Add(new InlineFormat
            {
                Type = f.TryGetProperty("type", out var t) ? t.GetString() ?? string.Empty : string.Empty,
                Text = f.TryGetProperty("text", out var txt) ? txt.GetString() ?? string.Empty : string.Empty
            });
        }

        return result.Count > 0 ? result : null;
    }

    private static List<TableRow> ParseTableRows(JsonElement rows)
    {
        var tableRows = new List<TableRow>();
        foreach (var row in rows.EnumerateArray())
        {
            var cells = new List<TableCell>();
            if (row.TryGetProperty("cells", out var cellsElement))
            {
                foreach (var cell in cellsElement.EnumerateArray())
                {
                    cells.Add(new TableCell
                    {
                        Text = cell.TryGetProperty("text", out var t) ? t.GetString() ?? string.Empty : string.Empty
                    });
                }
            }

            tableRows.Add(new TableRow { Cells = cells });
        }

        return tableRows;
    }

    private class JsonInstructionWrapper
    {
        public List<JsonInstructionDto>? Operations { get; set; }
    }

    private class JsonInstructionDto
    {
        public string? Type { get; set; }
        public string? Text { get; set; }
        public string? Style { get; set; }
        public string? Find { get; set; }
        public string? Replace { get; set; }
        public string? Target { get; set; }
        public ParagraphContent? Content { get; set; }
        public JsonElement? Blocks { get; set; }
    }
}
