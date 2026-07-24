using System.Text.Json;

namespace DocxEditor.Core.Serialization;

public class DocxInstructionValidator
{
    private static readonly HashSet<string> KnownOpTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "create", "addParagraph", "addRichContent", "replaceText",
        "replaceWithRichContent", "insertAfter"
    };

    private static readonly HashSet<string> KnownBlockTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "paragraph", "heading", "list", "table", "blockquote", "code", "horizontalRule", "custom"
    };

    private readonly List<string> _errors;

    public DocxInstructionValidator()
    {
        _errors = [];
    }

    public IReadOnlyList<string> Validate(string json)
    {
        _errors.Clear();

        JsonDocument document;
        try
        {
            document = JsonDocument.Parse(json);
        }
        catch (JsonException ex)
        {
            _errors.Add($"Invalid JSON: {ex.Message}");
            return _errors.AsReadOnly();
        }

        using (document)
        {
            var root = document.RootElement;
            if (root.ValueKind != JsonValueKind.Object)
            {
                _errors.Add("Root element must be a JSON object with an 'operations' array.");
                return _errors.AsReadOnly();
            }

            if (!root.TryGetProperty("operations", out var operations))
            {
                _errors.Add("Missing required field 'operations' at root.");
                return _errors.AsReadOnly();
            }

            if (operations.ValueKind != JsonValueKind.Array)
            {
                _errors.Add("Field 'operations' must be a JSON array.");
                return _errors.AsReadOnly();
            }

            for (int i = 0; i < operations.GetArrayLength(); i++)
            {
                ValidateOperation(operations[i], i);
            }
        }

        return _errors.AsReadOnly();
    }

    private void ValidateOperation(JsonElement op, int index)
    {
        var prefix = $"operations[{index}]";

        if (op.ValueKind != JsonValueKind.Object)
        {
            _errors.Add($"{prefix}: must be a JSON object.");
            return;
        }

        if (!op.TryGetProperty("type", out var typeProp))
        {
            _errors.Add($"{prefix}: missing required field 'type'.");
            return;
        }

        var type = typeProp.GetString();
        if (string.IsNullOrEmpty(type))
        {
            _errors.Add($"{prefix}.type: must be a non-empty string.");
            return;
        }

        if (!KnownOpTypes.Contains(type))
        {
            _errors.Add($"{prefix}.type: '{type}' is not a supported operation type. Valid types: {string.Join(", ", KnownOpTypes.OrderBy(t => t))}.");
            return;
        }

        switch (type.ToLowerInvariant())
        {
            case "addparagraph":
                if (!op.TryGetProperty("text", out _))
                    _errors.Add($"{prefix}: missing required field 'text' for addParagraph.");
                break;
            case "replacetext":
                if (!op.TryGetProperty("find", out _))
                    _errors.Add($"{prefix}: missing required field 'find' for replaceText.");
                if (!op.TryGetProperty("replace", out var rep) || rep.ValueKind != JsonValueKind.String)
                    _errors.Add($"{prefix}: missing required field 'replace' for replaceText.");
                break;
            case "insertafter":
                if (!op.TryGetProperty("target", out _))
                    _errors.Add($"{prefix}: missing required field 'target' for insertAfter.");
                if (!op.TryGetProperty("content", out var content) || content.ValueKind != JsonValueKind.Object)
                    _errors.Add($"{prefix}: missing required field 'content' for insertAfter.");
                else
                {
                    if (!content.TryGetProperty("text", out _))
                        _errors.Add($"{prefix}.content: missing required field 'text'.");
                }
                break;
            case "addrichcontent":
                if (!op.TryGetProperty("blocks", out var blocks) || blocks.ValueKind != JsonValueKind.Array)
                    _errors.Add($"{prefix}: missing required field 'blocks' for addRichContent (must be an array).");
                else
                    ValidateBlocks(blocks, $"{prefix}.blocks");
                break;
            case "replacewithrichcontent":
                if (!op.TryGetProperty("target", out _))
                    _errors.Add($"{prefix}: missing required field 'target' for replaceWithRichContent.");
                if (!op.TryGetProperty("blocks", out var rBlocks) || rBlocks.ValueKind != JsonValueKind.Array)
                    _errors.Add($"{prefix}: missing required field 'blocks' for replaceWithRichContent (must be an array).");
                else
                    ValidateBlocks(rBlocks, $"{prefix}.blocks");
                break;
        }

        var knownFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "type", "text", "style", "find", "replace", "target", "content", "blocks"
        };

        foreach (var property in op.EnumerateObject())
        {
            if (!knownFields.Contains(property.Name))
            {
                _errors.Add($"{prefix}: unknown field '{property.Name}' for operation type '{type}'.");
            }
        }
    }

    private void ValidateBlocks(JsonElement blocks, string path)
    {
        for (int i = 0; i < blocks.GetArrayLength(); i++)
        {
            var block = blocks[i];
            var blockPath = $"{path}[{i}]";

            if (block.ValueKind != JsonValueKind.Object)
            {
                _errors.Add($"{blockPath}: must be a JSON object.");
                continue;
            }

            if (!block.TryGetProperty("type", out var typeProp))
            {
                _errors.Add($"{blockPath}: missing required field 'type'.");
                continue;
            }

            var type = typeProp.GetString();
            if (string.IsNullOrEmpty(type))
            {
                _errors.Add($"{blockPath}.type: must be a non-empty string.");
                continue;
            }

            if (!KnownBlockTypes.Contains(type))
            {
                _errors.Add($"{blockPath}.type: '{type}' is not a supported block type. Valid types: {string.Join(", ", KnownBlockTypes.OrderBy(t => t))}.");
                continue;
            }

            switch (type.ToLowerInvariant())
            {
                case "heading":
                    if (!block.TryGetProperty("level", out var level) || level.ValueKind != JsonValueKind.Number)
                        _errors.Add($"{blockPath}: missing or invalid 'level' field for heading block.");
                    break;
                case "list":
                    if (!block.TryGetProperty("items", out var items) || items.ValueKind != JsonValueKind.Array)
                        _errors.Add($"{blockPath}: missing 'items' array field for list block.");
                    break;
                case "table":
                    if (!block.TryGetProperty("rows", out var rows) || rows.ValueKind != JsonValueKind.Array)
                        _errors.Add($"{blockPath}: missing 'rows' array field for table block.");
                    break;
                case "custom":
                    if (!block.TryGetProperty("customType", out _))
                        _errors.Add($"{blockPath}: missing required field 'customType' for custom block.");
                    break;
            }

            var blockKnownFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "type", "text", "style", "level", "ordered", "items", "rows",
                "language", "customType", "inlineFormats"
            };

            foreach (var property in block.EnumerateObject())
            {
                if (!blockKnownFields.Contains(property.Name))
                {
                    _errors.Add($"{blockPath}: unknown field '{property.Name}' for block type '{type}'.");
                }
            }
        }
    }
}
