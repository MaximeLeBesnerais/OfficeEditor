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
        string? GetString(string key) => dict.TryGetValue(key, out var v) ? v?.ToString() : null;
        int GetInt(string key) => dict.TryGetValue(key, out var v) && int.TryParse(v?.ToString(), out var i) ? i : 0;
        bool GetBool(string key) => dict.TryGetValue(key, out var v) && bool.TryParse(v?.ToString(), out var b) && b;
        List<object>? GetList(string key) => dict.TryGetValue(key, out var v) ? v as List<object> : null;

        return type.ToLowerInvariant() switch
        {
            "paragraph" => new ParagraphBlock
            {
                Text = GetString("text") ?? string.Empty,
                Style = GetString("style"),
                InlineFormats = GetList("inlineFormats") is List<object> formats ? ParseInlineFormats(formats) : null
            },
            "heading" => new HeadingBlock
            {
                Level = GetInt("level"),
                Text = GetString("text") ?? string.Empty,
                Style = GetString("style")
            },
            "list" => new ListBlock
            {
                Ordered = GetBool("ordered"),
                Items = GetList("items") is List<object> items ? items.Select(i => i?.ToString() ?? string.Empty).ToList() : [],
                Style = GetString("style")
            },
            "table" => new TableBlock
            {
                Rows = GetList("rows") is List<object> rows ? ParseTableRows(rows) : []
            },
            "blockquote" => new BlockquoteBlock
            {
                Text = GetString("text") ?? string.Empty,
                Style = GetString("style")
            },
            "code" => new CodeBlock
            {
                Text = GetString("text") ?? string.Empty,
                Language = GetString("language"),
                Style = GetString("style")
            },
            "horizontalrule" => new HorizontalRuleBlock(),
            "custom" => new CustomBlock
            {
                CustomType = GetString("customType") ?? string.Empty,
                Text = GetString("text") ?? string.Empty,
                Style = GetString("style")
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
                continue;

            result.Add(new InlineFormat
            {
                Type = (dict.TryGetValue("type", out var t) ? t?.ToString() : null) ?? string.Empty,
                Text = (dict.TryGetValue("text", out var txt) ? txt?.ToString() : null) ?? string.Empty
            });
        }

        return result.Count > 0 ? result : null;
    }

    private static List<TableRow> ParseTableRows(List<object> rows)
    {
        var tableRows = new List<TableRow>();
        foreach (var row in rows)
        {
            if (row is not Dictionary<object, object> rowDict)
                continue;

            var cells = new List<TableCell>();
            if (rowDict.TryGetValue("cells", out var cellsObj) && cellsObj is List<object> cellsList)
            {
                foreach (var cell in cellsList)
                {
                    if (cell is not Dictionary<object, object> cellDict)
                        continue;

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
