using System.Text.Json.Serialization;

namespace XlsxEditor.Core.Instructions;

public sealed record XlsxInstructionSet
{
    [JsonPropertyName("version")]
    public string Version { get; init; } = "1.0";

    [JsonPropertyName("description")]
    public string? Description { get; init; }

    [JsonPropertyName("worksheets")]
    public List<WorksheetInstruction> Worksheets { get; init; } = new();

    [JsonPropertyName("variables")]
    public Dictionary<string, string>? Variables { get; init; }
}

public sealed record WorksheetInstruction
{
    /// <summary>Worksheet name (required, validated).</summary>
    [JsonPropertyName("name")]
    public string Name { get; init; } = string.Empty;

    /// <summary>Header row values (optional). Written to row 1.</summary>
    [JsonPropertyName("headers")]
    public List<string>? Headers { get; init; }

    /// <summary>
    /// Row data as a list of row-arrays. Each inner array is one row's cell values.
    /// Values may be plain text, numbers, formulas (starting with "="), or variables ({{…}}).
    /// </summary>
    [JsonPropertyName("rows")]
    public List<List<string>>? Rows { get; init; }

    /// <summary>Discrete cell-level instructions (address + value OR formula).</summary>
    [JsonPropertyName("cells")]
    public List<CellInstruction>? Cells { get; init; }
}

public sealed record CellInstruction
{
    [JsonPropertyName("address")]
    public string Address { get; init; } = string.Empty;

    [JsonPropertyName("value")]
    public string? Value { get; init; }

    [JsonPropertyName("formula")]
    public string? Formula { get; init; }

    /// <summary>Optional type hint: number, string, boolean, date.</summary>
    [JsonPropertyName("type")]
    public string? Type { get; init; }

    /// <summary>Excel number format string (e.g. "0.00", "$#,##0.00").</summary>
    [JsonPropertyName("numberFormat")]
    public string? NumberFormat { get; init; }

    /// <summary>Style reference (planned for Phase 3 style builder).</summary>
    [JsonPropertyName("style")]
    public string? Style { get; init; }
}
