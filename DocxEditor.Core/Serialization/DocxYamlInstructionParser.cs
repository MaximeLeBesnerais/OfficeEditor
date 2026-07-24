using DocxEditor.Core.Models;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace DocxEditor.Core.Serialization;

public class DocxYamlInstructionParser
{
    private readonly IDeserializer _deserializer;

    public DocxYamlInstructionParser()
    {
        _deserializer = new DeserializerBuilder()
            .WithNamingConvention(CamelCaseNamingConvention.Instance)
            .Build();
    }

    public DocumentInstructions Parse(string yaml)
    {
        var wrapper = _deserializer.Deserialize<YamlInstructionWrapper>(yaml);
        if (wrapper?.Operations == null)
        {
            throw new ArgumentException("Invalid YAML instruction file.");
        }

        var instructions = new List<Instruction>();
        foreach (var op in wrapper.Operations)
        {
            instructions.Add(ParseInstruction(op));
        }

        return new DocumentInstructions { Operations = instructions };
    }

    private Instruction ParseInstruction(YamlInstructionDto dto)
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
                // YamlDotNet ignores C# 'required', so Content.Text can be null at runtime.
                Content = dto.Content?.Text != null
                    ? dto.Content
                    : throw new ArgumentException("Content with a 'text' field is required for insertAfter.")
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

    internal static List<ContentBlock> ParseContentBlocks(List<object> blocksList)
    {
        var blocks = new List<ContentBlock>();
        foreach (var item in blocksList)
        {
            if (item is not Dictionary<object, object> dict)
            {
                throw new ArgumentException("Each block must be an object.");
            }

            var typeObj2 = dict.TryGetValue("type", out var typeObj) ? typeObj?.ToString() : null;
            var type = typeObj2 ?? throw new ArgumentException("Each block must have a 'type' field.");

            blocks.Add(ParseBlock(type, dict));
        }

        return blocks;
    }

    private static ContentBlock ParseBlock(string type, Dictionary<object, object> dict)
    {
        // YAML scalars arrive as plain objects; lists/mappings in a scalar field are a
        // malformed value kind and must fail loudly instead of ToString()'ing into garbage.
        string? GetScalar(string key)
        {
            if (!dict.TryGetValue(key, out var v) || v == null)
                return null;
            if (v is List<object> or Dictionary<object, object>)
                throw new ArgumentException($"Block '{type}': field '{key}' must be a scalar value.");
            return v.ToString();
        }

        int GetInt(string key, int defaultValue)
        {
            if (!dict.TryGetValue(key, out var v) || v == null)
                return defaultValue;
            var scalar = GetScalar(key);
            if (int.TryParse(scalar, System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out var i))
                return i;
            throw new ArgumentException($"Block '{type}': field '{key}' must be an integer, got '{scalar}'.");
        }

        bool GetBool(string key)
        {
            if (!dict.TryGetValue(key, out var v) || v == null)
                return false;
            var scalar = GetScalar(key);
            if (bool.TryParse(scalar, out var b))
                return b;
            throw new ArgumentException($"Block '{type}': field '{key}' must be a boolean, got '{scalar}'.");
        }

        List<object>? GetList(string key)
        {
            if (!dict.TryGetValue(key, out var v) || v == null)
                return null;
            return v as List<object>
                ?? throw new ArgumentException($"Block '{type}': field '{key}' must be a list.");
        }

        string ScalarItem(object? item, string key)
            => item is List<object> or Dictionary<object, object>
                ? throw new ArgumentException($"Block '{type}': entries of '{key}' must be scalar values.")
                : item?.ToString() ?? string.Empty;

        return type.ToLowerInvariant() switch
        {
            "paragraph" => new ParagraphBlock
            {
                Text = GetScalar("text") ?? string.Empty,
                Style = GetScalar("style"),
                InlineFormats = GetList("inlineFormats") is List<object> formats ? ParseInlineFormats(formats) : null
            },
            "heading" => new HeadingBlock
            {
                // Default 1, matching the JSON parser and the validator.
                Level = GetInt("level", defaultValue: 1),
                Text = GetScalar("text") ?? string.Empty,
                Style = GetScalar("style")
            },
            "list" => new ListBlock
            {
                Ordered = GetBool("ordered"),
                Items = GetList("items") is List<object> items ? items.Select(i => ScalarItem(i, "items")).ToList() : [],
                Style = GetScalar("style")
            },
            "table" => new TableBlock
            {
                Rows = GetList("rows") is List<object> rows ? ParseTableRows(rows, type) : []
            },
            "blockquote" => new BlockquoteBlock
            {
                Text = GetScalar("text") ?? string.Empty,
                Style = GetScalar("style")
            },
            "code" => new CodeBlock
            {
                Text = GetScalar("text") ?? string.Empty,
                Language = GetScalar("language"),
                Style = GetScalar("style")
            },
            "horizontalrule" => new HorizontalRuleBlock(),
            "custom" => new CustomBlock
            {
                CustomType = GetScalar("customType") ?? string.Empty,
                Text = GetScalar("text") ?? string.Empty,
                Style = GetScalar("style")
            },
            _ => throw new NotSupportedException($"Block type '{type}' is not supported. Valid types: paragraph, heading, list, table, blockquote, code, horizontalRule, custom.")
        };
    }

    private static List<InlineFormat>? ParseInlineFormats(List<object> formats)
    {
        if (formats.Count == 0)
            return null;

        var result = new List<InlineFormat>();
        foreach (var item in formats)
        {
            if (item is not Dictionary<object, object> dict)
                throw new ArgumentException("Each inline format entry must be an object with 'type' and 'text' fields.");

            result.Add(new InlineFormat
            {
                Type = (dict.TryGetValue("type", out var t) ? t?.ToString() : null) ?? string.Empty,
                Text = (dict.TryGetValue("text", out var txt) ? txt?.ToString() : null) ?? string.Empty
            });
        }

        return result.Count > 0 ? result : null;
    }

    private static List<TableRow> ParseTableRows(List<object> rows, string blockType)
    {
        var tableRows = new List<TableRow>();
        foreach (var row in rows)
        {
            if (row is not Dictionary<object, object> rowDict)
                throw new ArgumentException($"Block '{blockType}': each row must be an object with a 'cells' list.");

            var cells = new List<TableCell>();
            if (rowDict.TryGetValue("cells", out var cellsObj) && cellsObj != null)
            {
                if (cellsObj is not List<object> cellsList)
                    throw new ArgumentException($"Block '{blockType}': field 'cells' must be a list.");

                foreach (var cell in cellsList)
                {
                    if (cell is not Dictionary<object, object> cellDict)
                        throw new ArgumentException($"Block '{blockType}': each cell must be an object with a 'text' field.");

                    cells.Add(new TableCell
                    {
                        Text = (cellDict.TryGetValue("text", out var t) ? t?.ToString() : null) ?? string.Empty
                    });
                }
            }

            tableRows.Add(new TableRow { Cells = cells });
        }

        return tableRows;
    }

    private class YamlInstructionWrapper
    {
        public List<YamlInstructionDto>? Operations { get; set; }
    }

    private class YamlInstructionDto
    {
        public string? Type { get; set; }
        public string? Text { get; set; }
        public string? Style { get; set; }
        public string? Find { get; set; }
        public string? Replace { get; set; }
        public string? Target { get; set; }
        public ParagraphContent? Content { get; set; }
        public List<object>? Blocks { get; set; }
    }
}
