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

    /// <summary>
    /// Validates instruction JSON and returns field-level errors. Each call returns a new
    /// list; callers may keep or mutate it without affecting subsequent validations.
    /// </summary>
    public IReadOnlyList<string> Validate(string json)
    {
        var errors = new List<string>();

        if (json is null)
        {
            // The validation contract is result-based: a null argument is an invalid
            // document, not an argument error, and must never throw.
            errors.Add("Invalid JSON: input is null.");
            return errors;
        }

        JsonDocument document;
        try
        {
            document = JsonDocument.Parse(json);
        }
        catch (JsonException ex)
        {
            errors.Add($"Invalid JSON: {ex.Message}");
            return errors;
        }

        using (document)
        {
            var root = document.RootElement;
            if (root.ValueKind != JsonValueKind.Object)
            {
                errors.Add("Root element must be a JSON object with an 'operations' array.");
                return errors;
            }

            if (!root.TryGetProperty("operations", out var operations))
            {
                errors.Add("Missing required field 'operations' at root.");
                return errors;
            }

            if (operations.ValueKind != JsonValueKind.Array)
            {
                errors.Add("Field 'operations' must be a JSON array.");
                return errors;
            }

            for (int i = 0; i < operations.GetArrayLength(); i++)
            {
                ValidateOperation(operations[i], i, errors);
            }
        }

        return errors;
    }

    private static void ValidateOperation(JsonElement op, int index, List<string> errors)
    {
        var prefix = $"operations[{index}]";

        if (op.ValueKind != JsonValueKind.Object)
        {
            errors.Add($"{prefix}: must be a JSON object.");
            return;
        }

        if (!op.TryGetProperty("type", out var typeProp))
        {
            errors.Add($"{prefix}: missing required field 'type'.");
            return;
        }

        // A 'type' that exists but is not a string must be reported as a validation error
        // instead of throwing InvalidOperationException from GetString().
        if (typeProp.ValueKind != JsonValueKind.String)
        {
            errors.Add($"{prefix}.type: must be a string.");
            return;
        }

        var type = typeProp.GetString();
        if (string.IsNullOrEmpty(type))
        {
            errors.Add($"{prefix}.type: must be a non-empty string.");
            return;
        }

        if (!KnownOpTypes.Contains(type))
        {
            errors.Add($"{prefix}.type: '{type}' is not a supported operation type. Valid types: {string.Join(", ", KnownOpTypes.OrderBy(t => t))}.");
            return;
        }

        switch (type.ToLowerInvariant())
        {
            case "addparagraph":
                if (!op.TryGetProperty("text", out _))
                    errors.Add($"{prefix}: missing required field 'text' for addParagraph.");
                break;
            case "replacetext":
                // Empty find corrupts: string.Replace("", x) inserts between every character.
                if (!op.TryGetProperty("find", out var find) || find.ValueKind != JsonValueKind.String
                    || string.IsNullOrEmpty(find.GetString()))
                    errors.Add($"{prefix}: missing or empty required field 'find' for replaceText.");
                if (!op.TryGetProperty("replace", out var rep) || rep.ValueKind != JsonValueKind.String)
                    errors.Add($"{prefix}: missing required field 'replace' for replaceText.");
                break;
            case "insertafter":
                if (!op.TryGetProperty("target", out _))
                    errors.Add($"{prefix}: missing required field 'target' for insertAfter.");
                if (!op.TryGetProperty("content", out var content) || content.ValueKind != JsonValueKind.Object)
                    errors.Add($"{prefix}: missing required field 'content' for insertAfter.");
                else
                {
                    if (!content.TryGetProperty("text", out _))
                        errors.Add($"{prefix}.content: missing required field 'text'.");
                }
                break;
            case "addrichcontent":
                if (!op.TryGetProperty("blocks", out var blocks) || blocks.ValueKind != JsonValueKind.Array)
                    errors.Add($"{prefix}: missing required field 'blocks' for addRichContent (must be an array).");
                else
                    ValidateBlocks(blocks, $"{prefix}.blocks", errors);
                break;
            case "replacewithrichcontent":
                if (!op.TryGetProperty("target", out _))
                    errors.Add($"{prefix}: missing required field 'target' for replaceWithRichContent.");
                if (!op.TryGetProperty("blocks", out var rBlocks) || rBlocks.ValueKind != JsonValueKind.Array)
                    errors.Add($"{prefix}: missing required field 'blocks' for replaceWithRichContent (must be an array).");
                else
                    ValidateBlocks(rBlocks, $"{prefix}.blocks", errors);
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
                errors.Add($"{prefix}: unknown field '{property.Name}' for operation type '{type}'.");
            }
        }
    }

    private static void ValidateBlocks(JsonElement blocks, string path, List<string> errors)
    {
        for (int i = 0; i < blocks.GetArrayLength(); i++)
        {
            var block = blocks[i];
            var blockPath = $"{path}[{i}]";

            if (block.ValueKind != JsonValueKind.Object)
            {
                errors.Add($"{blockPath}: must be a JSON object.");
                continue;
            }

            if (!block.TryGetProperty("type", out var typeProp))
            {
                errors.Add($"{blockPath}: missing required field 'type'.");
                continue;
            }

            if (typeProp.ValueKind != JsonValueKind.String)
            {
                errors.Add($"{blockPath}.type: must be a string.");
                continue;
            }

            var type = typeProp.GetString();
            if (string.IsNullOrEmpty(type))
            {
                errors.Add($"{blockPath}.type: must be a non-empty string.");
                continue;
            }

            if (!KnownBlockTypes.Contains(type))
            {
                errors.Add($"{blockPath}.type: '{type}' is not a supported block type. Valid types: {string.Join(", ", KnownBlockTypes.OrderBy(t => t))}.");
                continue;
            }

            switch (type.ToLowerInvariant())
            {
                case "heading":
                    // Level is optional (parsers default it to 1); when present it must be an integer.
                    if (block.TryGetProperty("level", out var level)
                        && (level.ValueKind != JsonValueKind.Number || !level.TryGetInt32(out _)))
                        errors.Add($"{blockPath}: 'level' must be an integer for heading block.");
                    break;
                case "list":
                    if (!block.TryGetProperty("items", out var items) || items.ValueKind != JsonValueKind.Array)
                        errors.Add($"{blockPath}: missing 'items' array field for list block.");
                    break;
                case "table":
                    if (!block.TryGetProperty("rows", out var rows) || rows.ValueKind != JsonValueKind.Array)
                        errors.Add($"{blockPath}: missing 'rows' array field for table block.");
                    break;
                case "custom":
                    if (!block.TryGetProperty("customType", out _))
                        errors.Add($"{blockPath}: missing required field 'customType' for custom block.");
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
                    errors.Add($"{blockPath}: unknown field '{property.Name}' for block type '{type}'.");
                }
            }
        }
    }
}
