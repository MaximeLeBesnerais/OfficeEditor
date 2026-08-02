using System.Text.Json.Serialization;

namespace XlsxEditor.Core.Instructions;

public sealed record XlsxInstructionSet
{
    [JsonPropertyName("version")]
    public string Version { get; init; } = string.Empty;

    [JsonPropertyName("description")]
    public string? Description { get; init; }

    /// <summary>Workbook-level document properties (title, author, …).</summary>
    [JsonPropertyName("metadata")]
    public WorkbookMetadata? Metadata { get; init; }

    /// <summary>
    /// Named style definitions. Cells reference styles by name (see
    /// <see cref="CellInstruction.Style"/> and <see cref="WorksheetInstruction.HeaderStyle"/>);
    /// the legacy numeric style-id form also remains accepted.
    /// </summary>
    [JsonPropertyName("styles")]
    public List<NamedStyle>? Styles { get; init; }

    [JsonPropertyName("worksheets")]
    public List<WorksheetInstruction> Worksheets { get; init; } = new();

    [JsonPropertyName("variables")]
    public Dictionary<string, string>? Variables { get; init; }
}

/// <summary>Workbook document properties, mapped onto the OOXML core properties.</summary>
public sealed record WorkbookMetadata
{
    [JsonPropertyName("title")]
    public string? Title { get; init; }

    [JsonPropertyName("subject")]
    public string? Subject { get; init; }

    [JsonPropertyName("author")]
    public string? Author { get; init; }

    [JsonPropertyName("category")]
    public string? Category { get; init; }

    [JsonPropertyName("keywords")]
    public string? Keywords { get; init; }

    [JsonPropertyName("comments")]
    public string? Comments { get; init; }
}

/// <summary>
/// A named, reusable style definition. Every aspect is optional; unset aspects inherit
/// the workbook's defaults. The executor's style machinery (owned by another branch)
/// maps these onto cellXfs entries; the planner carries them unchanged.
/// </summary>
public sealed record NamedStyle
{
    [JsonPropertyName("name")]
    public string Name { get; init; } = string.Empty;

    [JsonPropertyName("font")]
    public FontStyleInstruction? Font { get; init; }

    [JsonPropertyName("fill")]
    public FillStyleInstruction? Fill { get; init; }

    [JsonPropertyName("border")]
    public BorderStyleInstruction? Border { get; init; }

    [JsonPropertyName("alignment")]
    public AlignmentStyleInstruction? Alignment { get; init; }

    [JsonPropertyName("numberFormat")]
    public string? NumberFormat { get; init; }
}

public sealed record FontStyleInstruction
{
    [JsonPropertyName("bold")]
    public bool? Bold { get; init; }

    [JsonPropertyName("italic")]
    public bool? Italic { get; init; }

    /// <summary>ARGB hex color ("FF0000") or a documented CSS/Excel color name.</summary>
    [JsonPropertyName("color")]
    public string? Color { get; init; }

    [JsonPropertyName("size")]
    public double? Size { get; init; }
}

public sealed record FillStyleInstruction
{
    [JsonPropertyName("color")]
    public string? Color { get; init; }

    /// <summary>Excel pattern type: "none", "solid", "gray125", …</summary>
    [JsonPropertyName("pattern")]
    public string? Pattern { get; init; }
}

public sealed record BorderStyleInstruction
{
    [JsonPropertyName("left")]
    public BorderEdgeInstruction? Left { get; init; }

    [JsonPropertyName("right")]
    public BorderEdgeInstruction? Right { get; init; }

    [JsonPropertyName("top")]
    public BorderEdgeInstruction? Top { get; init; }

    [JsonPropertyName("bottom")]
    public BorderEdgeInstruction? Bottom { get; init; }
}

public sealed record BorderEdgeInstruction
{
    /// <summary>Excel border style: "thin", "medium", "thick", "double", "dashed", …</summary>
    [JsonPropertyName("style")]
    public string? Style { get; init; }

    [JsonPropertyName("color")]
    public string? Color { get; init; }
}

public sealed record AlignmentStyleInstruction
{
    /// <summary>"left", "center", "right", "fill", "justify", "centerContinuous".</summary>
    [JsonPropertyName("horizontal")]
    public string? Horizontal { get; init; }

    /// <summary>"top", "center", "bottom", "justify", "distributed".</summary>
    [JsonPropertyName("vertical")]
    public string? Vertical { get; init; }

    [JsonPropertyName("wrapText")]
    public bool? WrapText { get; init; }
}

public sealed record WorksheetInstruction
{
    /// <summary>Worksheet name (required, validated).</summary>
    [JsonPropertyName("name")]
    public string Name { get; init; } = string.Empty;

    /// <summary>Header row values (optional). Written to the <see cref="StartRow"/>.</summary>
    [JsonPropertyName("headers")]
    public List<string>? Headers { get; init; }

    /// <summary>
    /// Named style (or legacy numeric style id) applied to the header row.
    /// </summary>
    [JsonPropertyName("headerStyle")]
    public string? HeaderStyle { get; init; }

    /// <summary>
    /// 1-based row the header row is written to (default 1). Data rows follow at
    /// <c>startRow + (headers present ? 1 : 0)</c>. Discrete cells keep their
    /// explicit addresses and are unaffected.
    /// </summary>
    [JsonPropertyName("startRow")]
    public int? StartRow { get; init; }

    /// <summary>
    /// Positional typed column definitions. <c>columns[0]</c> describes column A,
    /// <c>columns[1]</c> column B, and so on. Each entry may set a display name, an
    /// explicit width, a style, and a default cell type for the column.
    /// </summary>
    [JsonPropertyName("columns")]
    public List<ColumnInstruction>? Columns { get; init; }

    /// <summary>
    /// Row data as a list of row-arrays. Each inner array is one row's cell values.
    /// Values may be plain text, numbers, formulas (starting with "="), or variables ({{…}}).
    /// </summary>
    [JsonPropertyName("rows")]
    public List<List<string>>? Rows { get; init; }

    /// <summary>Explicit row heights, in points.</summary>
    [JsonPropertyName("rowHeights")]
    public List<RowHeightInstruction>? RowHeights { get; init; }

    /// <summary>Discrete cell-level instructions (address + value OR formula).</summary>
    [JsonPropertyName("cells")]
    public List<CellInstruction>? Cells { get; init; }

    /// <summary>Cell ranges to merge (A1 notation, e.g. "A1:C3").</summary>
    [JsonPropertyName("merges")]
    public List<string>? Merges { get; init; }

    /// <summary>Freeze panes: rows above <c>row</c> and columns left of <c>column</c> are frozen.</summary>
    [JsonPropertyName("freezePanes")]
    public FreezePanesInstruction? FreezePanes { get; init; }

    /// <summary>Standalone autofilter range (A1 notation, e.g. "A1:D20"). Not a table.</summary>
    [JsonPropertyName("autoFilter")]
    public string? AutoFilter { get; init; }

    /// <summary>Excel table declarations (display name + range).</summary>
    [JsonPropertyName("tables")]
    public List<TableInstruction>? Tables { get; init; }
}

public sealed record ColumnInstruction
{
    /// <summary>Optional display name for the column (used by table headers).</summary>
    [JsonPropertyName("name")]
    public string? Name { get; init; }

    /// <summary>Explicit column width in Excel column-width units (0 exclusive, 255 inclusive).</summary>
    [JsonPropertyName("width")]
    public double? Width { get; init; }

    /// <summary>Named style (or legacy numeric style id) applied to the column's cells.</summary>
    [JsonPropertyName("style")]
    public string? Style { get; init; }

    /// <summary>Default cell type for the column: auto/string/number/boolean/date/datetime.</summary>
    [JsonPropertyName("type")]
    public string? Type { get; init; }
}

public sealed record RowHeightInstruction
{
    [JsonPropertyName("row")]
    public int? Row { get; init; }

    [JsonPropertyName("height")]
    public double? Height { get; init; }
}

public sealed record FreezePanesInstruction
{
    /// <summary>
    /// Top-left cell of the scrollable region, e.g. "A2" freezes row 1, "B1" freezes column A.
    /// Mutually exclusive with <see cref="Row"/>/<see cref="Column"/>.
    /// </summary>
    [JsonPropertyName("cell")]
    public string? Cell { get; init; }

    /// <summary>Freeze rows 1..<c>row</c>-1. Defaults to 1 when <see cref="Cell"/> is unset.</summary>
    [JsonPropertyName("row")]
    public int? Row { get; init; }

    /// <summary>Freeze columns 1..<c>column</c>-1. Defaults to 1 when <see cref="Cell"/> is unset.</summary>
    [JsonPropertyName("column")]
    public int? Column { get; init; }
}

public sealed record TableInstruction
{
    /// <summary>Table display name (unique workbook-wide, case-insensitive).</summary>
    [JsonPropertyName("name")]
    public string Name { get; init; } = string.Empty;

    /// <summary>Table range in A1 notation (e.g. "A1:D20").</summary>
    [JsonPropertyName("range")]
    public string Range { get; init; } = string.Empty;
}

public sealed record CellInstruction
{
    [JsonPropertyName("address")]
    public string Address { get; init; } = string.Empty;

    /// <summary>
    /// The cell's literal value. An explicit empty string is a valid, intentionally
    /// blank cell (distinct from the property being absent, which is an error unless
    /// <see cref="Formula"/> is set).
    /// </summary>
    [JsonPropertyName("value")]
    public string? Value { get; init; }

    [JsonPropertyName("formula")]
    public string? Formula { get; init; }

    /// <summary>
    /// Cell type: "auto", "string", "number", "boolean", "date", or "datetime".
    /// When unset (auto), the type is inferred from the resolved value.
    /// </summary>
    [JsonPropertyName("type")]
    public string? Type { get; init; }

    /// <summary>Excel number format (e.g. "0.00", "$#,##0.00", "yyyy-mm-dd").</summary>
    [JsonPropertyName("numberFormat")]
    public string? NumberFormat { get; init; }

    /// <summary>
    /// Style reference: a named style from the set's <c>styles</c> block, or the
    /// legacy numeric style-id string (e.g. "0").
    /// </summary>
    [JsonPropertyName("style")]
    public string? Style { get; init; }
}
