using System.Text.Json;
using PptxEditor.Core.Models;

namespace PptxEditor.Core.Serialization;

/// <summary>
/// Manual JSON parser for the PPTX instruction vocabulary, mirroring
/// DocxJsonInstructionParser: a switch on the op "type" builds the typed
/// instruction records. Each operation element is parsed individually so
/// missing/invalid fields raise an <see cref="ArgumentException"/> that names
/// the operation index ("operations[3]: ...").
/// </summary>
public class PptxJsonInstructionParser
{
    private static readonly HashSet<string> KnownFitModes = new(StringComparer.OrdinalIgnoreCase)
    {
        "stretch", "fill", "crop", "contain"
    };

    public PptxInstructionSet Parse(string json)
    {
        JsonDocument document;
        try
        {
            document = JsonDocument.Parse(json);
        }
        catch (JsonException ex)
        {
            throw new ArgumentException($"Invalid JSON instruction file: {ex.Message}");
        }

        using (document)
        {
            var root = document.RootElement;
            if (root.ValueKind != JsonValueKind.Object
                || !TryGetProperty(root, "operations", out var operations)
                || operations.ValueKind != JsonValueKind.Array)
            {
                throw new ArgumentException(
                    "Invalid JSON instruction file: expected an object with an 'operations' array.");
            }

            var instructions = new List<PptxInstruction>();
            var index = 0;
            foreach (var op in operations.EnumerateArray())
            {
                instructions.Add(ParseInstruction(op, index));
                index++;
            }

            return new PptxInstructionSet { Operations = instructions };
        }
    }

    private PptxInstruction ParseInstruction(JsonElement op, int index)
    {
        if (op.ValueKind != JsonValueKind.Object)
        {
            throw new ArgumentException($"operations[{index}]: each operation must be a JSON object.");
        }

        var type = GetRequiredString(op, "type", index);
        return type.ToLowerInvariant() switch
        {
            "replacetext" => new PptxReplaceTextInstruction
            {
                Slide = GetRequiredPositiveInt(op, "slide", index),
                ElementId = GetRequiredPositiveUInt(op, "elementId", index),
                Text = GetRequiredString(op, "text", index)
            },
            "replaceimage" => new PptxReplaceImageInstruction
            {
                Slide = GetRequiredPositiveInt(op, "slide", index),
                ElementId = GetRequiredPositiveUInt(op, "elementId", index),
                Image = GetRequiredString(op, "image", index),
                Fit = GetOptionalFit(op, index)
            },
            "replacetable" => new PptxReplaceTableInstruction
            {
                Slide = GetRequiredPositiveInt(op, "slide", index),
                ElementId = GetRequiredPositiveUInt(op, "elementId", index),
                Rows = GetRequiredRows(op, index)
            },
            "moveslide" => new PptxMoveSlideInstruction
            {
                From = GetRequiredPositiveInt(op, "from", index),
                To = GetRequiredPositiveInt(op, "to", index)
            },
            "duplicateslide" => new PptxDuplicateSlideInstruction
            {
                Slide = GetRequiredPositiveInt(op, "slide", index),
                Position = GetOptionalPositiveInt(op, "position", index)
            },
            "deleteslide" => new PptxDeleteSlideInstruction
            {
                Slide = GetRequiredPositiveInt(op, "slide", index)
            },
            _ => throw new ArgumentException(
                $"operations[{index}]: instruction type '{type}' is not supported. " +
                "Supported types: replaceText, replaceImage, replaceTable, moveSlide, duplicateSlide, deleteSlide.")
        };
    }

    private static string GetRequiredString(JsonElement op, string name, int index)
    {
        if (!TryGetProperty(op, name, out var value) || value.ValueKind != JsonValueKind.String)
        {
            throw new ArgumentException($"operations[{index}]: '{name}' is required and must be a string.");
        }
        return value.GetString()!;
    }

    private static int GetRequiredPositiveInt(JsonElement op, string name, int index)
    {
        if (!TryGetProperty(op, name, out var value)
            || value.ValueKind != JsonValueKind.Number
            || !value.TryGetInt32(out var number)
            || number < 1)
        {
            throw new ArgumentException(
                $"operations[{index}]: '{name}' is required and must be a positive integer (1-based).");
        }
        return number;
    }

    private static int? GetOptionalPositiveInt(JsonElement op, string name, int index)
    {
        if (!TryGetProperty(op, name, out var value) || value.ValueKind == JsonValueKind.Null)
        {
            return null;
        }
        if (value.ValueKind != JsonValueKind.Number || !value.TryGetInt32(out var number) || number < 1)
        {
            throw new ArgumentException(
                $"operations[{index}]: '{name}' must be a positive integer (1-based) when present.");
        }
        return number;
    }

    private static uint GetRequiredPositiveUInt(JsonElement op, string name, int index)
    {
        if (!TryGetProperty(op, name, out var value)
            || value.ValueKind != JsonValueKind.Number
            || !value.TryGetUInt32(out var number)
            || number < 1)
        {
            throw new ArgumentException(
                $"operations[{index}]: '{name}' is required and must be a positive integer.");
        }
        return number;
    }

    private static string? GetOptionalFit(JsonElement op, int index)
    {
        if (!TryGetProperty(op, "fit", out var value) || value.ValueKind == JsonValueKind.Null)
        {
            return null;
        }
        if (value.ValueKind != JsonValueKind.String)
        {
            throw new ArgumentException($"operations[{index}]: 'fit' must be a string when present.");
        }
        var fit = value.GetString()!;
        if (!KnownFitModes.Contains(fit))
        {
            throw new ArgumentException(
                $"operations[{index}]: 'fit' must be one of stretch|fill|crop|contain (got '{fit}').");
        }
        return fit.ToLowerInvariant();
    }

    private static List<List<string>> GetRequiredRows(JsonElement op, int index)
    {
        if (!TryGetProperty(op, "rows", out var value) || value.ValueKind != JsonValueKind.Array)
        {
            throw new ArgumentException($"operations[{index}]: 'rows' is required and must be an array of string arrays.");
        }

        var rows = new List<List<string>>();
        var rowIndex = 0;
        foreach (var rowElement in value.EnumerateArray())
        {
            if (rowElement.ValueKind != JsonValueKind.Array)
            {
                throw new ArgumentException($"operations[{index}]: 'rows[{rowIndex}]' must be an array of strings.");
            }
            var row = new List<string>();
            foreach (var cell in rowElement.EnumerateArray())
            {
                if (cell.ValueKind != JsonValueKind.String)
                {
                    throw new ArgumentException($"operations[{index}]: 'rows[{rowIndex}]' cells must all be strings.");
                }
                row.Add(cell.GetString()!);
            }
            rows.Add(row);
            rowIndex++;
        }

        if (rows.Count == 0 || rows[0].Count == 0)
        {
            throw new ArgumentException(
                $"operations[{index}]: 'rows' must contain at least one row and one column.");
        }
        if (rows.Any(row => row.Count != rows[0].Count))
        {
            throw new ArgumentException($"operations[{index}]: all 'rows' must have the same number of columns.");
        }

        return rows;
    }

    private static bool TryGetProperty(JsonElement element, string name, out JsonElement value)
    {
        foreach (var property in element.EnumerateObject())
        {
            if (string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase))
            {
                value = property.Value;
                return true;
            }
        }
        value = default;
        return false;
    }
}
