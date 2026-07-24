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
                // Empty find is rejected: string.Replace("", x) would insert the
                // replacement between every character of the document.
                Find = string.IsNullOrEmpty(dto.Find)
                    ? throw new ArgumentException("Find must be a non-empty string for replaceText.")
                    : dto.Find,
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

            blocks.Add(ParseBlock(type, blockElement, $"blocks[{blocks.Count}] ('{type}')"));
        }

        return blocks;
    }

    private static string? GetOptionalString(JsonElement element, string propertyName, string path)
    {
        if (!element.TryGetProperty(propertyName, out var prop) || prop.ValueKind == JsonValueKind.Null)
            return null;
        if (prop.ValueKind != JsonValueKind.String)
            throw new ArgumentException($"{path}: field '{propertyName}' must be a string, got {prop.ValueKind}.");
        return prop.GetString();
    }

    private static string GetStringOrEmpty(JsonElement element, string propertyName, string path)
        => GetOptionalString(element, propertyName, path) ?? string.Empty;

    private static int GetOptionalInt32(JsonElement element, string propertyName, int defaultValue, string path)
    {
        if (!element.TryGetProperty(propertyName, out var prop) || prop.ValueKind == JsonValueKind.Null)
            return defaultValue;
        if (prop.ValueKind == JsonValueKind.Number && prop.TryGetInt32(out var value))
            return value;
        throw new ArgumentException($"{path}: field '{propertyName}' must be an integer, got {prop}.");
    }

    private static bool GetOptionalBool(JsonElement element, string propertyName, string path)
    {
        if (!element.TryGetProperty(propertyName, out var prop) || prop.ValueKind == JsonValueKind.Null)
            return false;
        if (prop.ValueKind is JsonValueKind.True or JsonValueKind.False)
            return prop.GetBoolean();
        throw new ArgumentException($"{path}: field '{propertyName}' must be a boolean, got {prop.ValueKind}.");
    }

    private static List<string> GetOptionalStringList(JsonElement element, string propertyName, string path)
    {
        if (!element.TryGetProperty(propertyName, out var prop) || prop.ValueKind == JsonValueKind.Null)
            return [];
        if (prop.ValueKind != JsonValueKind.Array)
            throw new ArgumentException($"{path}: field '{propertyName}' must be an array, got {prop.ValueKind}.");

        var items = new List<string>();
        int index = 0;
        foreach (var item in prop.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.String)
                throw new ArgumentException($"{path}: field '{propertyName}[{index}]' must be a string, got {item.ValueKind}.");
            items.Add(item.GetString()!);
            index++;
        }

        return items;
    }

    private static ContentBlock ParseBlock(string type, JsonElement element, string path)
    {
        return type.ToLowerInvariant() switch
        {
            "paragraph" => new ParagraphBlock
            {
                Text = GetStringOrEmpty(element, "text", path),
                Style = GetOptionalString(element, "style", path),
                InlineFormats = element.TryGetProperty("inlineFormats", out var formats) ? ParseInlineFormats(formats, path) : null
            },
            "heading" => new HeadingBlock
            {
                Level = GetOptionalInt32(element, "level", defaultValue: 1, path),
                Text = GetStringOrEmpty(element, "text", path),
                Style = GetOptionalString(element, "style", path)
            },
            "list" => new ListBlock
            {
                Ordered = GetOptionalBool(element, "ordered", path),
                Items = GetOptionalStringList(element, "items", path),
                Style = GetOptionalString(element, "style", path)
            },
            "table" => new TableBlock
            {
                Rows = element.TryGetProperty("rows", out var rows) ? ParseTableRows(rows, path) : []
            },
            "blockquote" => new BlockquoteBlock
            {
                Text = GetStringOrEmpty(element, "text", path),
                Style = GetOptionalString(element, "style", path)
            },
            "code" => new CodeBlock
            {
                Text = GetStringOrEmpty(element, "text", path),
                Language = GetOptionalString(element, "language", path),
                Style = GetOptionalString(element, "style", path)
            },
            "horizontalrule" => new HorizontalRuleBlock(),
            "custom" => new CustomBlock
            {
                CustomType = GetStringOrEmpty(element, "customType", path),
                Text = GetStringOrEmpty(element, "text", path),
                Style = GetOptionalString(element, "style", path)
            },
            _ => throw new NotSupportedException($"Block type '{type}' is not supported. Valid types: paragraph, heading, list, table, blockquote, code, horizontalRule, custom.")
        };
    }

    private static List<InlineFormat>? ParseInlineFormats(JsonElement formats, string path)
    {
        if (formats.ValueKind == JsonValueKind.Null)
            return null;
        if (formats.ValueKind != JsonValueKind.Array)
            throw new ArgumentException($"{path}: field 'inlineFormats' must be an array, got {formats.ValueKind}.");
        if (formats.GetArrayLength() == 0)
            return null;

        var result = new List<InlineFormat>();
        int index = 0;
        foreach (var f in formats.EnumerateArray())
        {
            var formatPath = $"{path}.inlineFormats[{index}]";
            if (f.ValueKind != JsonValueKind.Object)
                throw new ArgumentException($"{formatPath}: each inline format must be a JSON object, got {f.ValueKind}.");
            result.Add(new InlineFormat
            {
                Type = GetStringOrEmpty(f, "type", formatPath),
                Text = GetStringOrEmpty(f, "text", formatPath)
            });
            index++;
        }

        return result.Count > 0 ? result : null;
    }

    private static List<TableRow> ParseTableRows(JsonElement rows, string path)
    {
        if (rows.ValueKind != JsonValueKind.Array)
            throw new ArgumentException($"{path}: field 'rows' must be an array, got {rows.ValueKind}.");

        var tableRows = new List<TableRow>();
        int rowIndex = 0;
        foreach (var row in rows.EnumerateArray())
        {
            var rowPath = $"{path}.rows[{rowIndex}]";
            if (row.ValueKind != JsonValueKind.Object)
                throw new ArgumentException($"{rowPath}: each row must be a JSON object, got {row.ValueKind}.");

            var cells = new List<TableCell>();
            if (row.TryGetProperty("cells", out var cellsElement) && cellsElement.ValueKind != JsonValueKind.Null)
            {
                if (cellsElement.ValueKind != JsonValueKind.Array)
                    throw new ArgumentException($"{rowPath}: field 'cells' must be an array, got {cellsElement.ValueKind}.");

                int cellIndex = 0;
                foreach (var cell in cellsElement.EnumerateArray())
                {
                    var cellPath = $"{rowPath}.cells[{cellIndex}]";
                    if (cell.ValueKind != JsonValueKind.Object)
                        throw new ArgumentException($"{cellPath}: each cell must be a JSON object, got {cell.ValueKind}.");
                    cells.Add(new TableCell
                    {
                        Text = GetStringOrEmpty(cell, "text", cellPath)
                    });
                    cellIndex++;
                }
            }

            tableRows.Add(new TableRow { Cells = cells });
            rowIndex++;
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
