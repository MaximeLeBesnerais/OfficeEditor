namespace DocxEditor.Core.Generation.Model;

/// <summary>A table with rows/cells, a header flag per row, optional column widths and
/// alignment. v1 requires a rectangular grid: every row has the same number of cells.</summary>
public sealed record TableBlock : FlowBlock
{
    /// <summary>Rows in document order. Required, at least one.</summary>
    public required IReadOnlyList<TableRow> Rows { get; init; }

    /// <summary>Optional per-column widths in points; length must equal the column count.</summary>
    public IReadOnlyList<double>? ColumnWidthsPt { get; init; }

    /// <summary>Optional table alignment.</summary>
    public TextAlignment? Alignment { get; init; }
}

/// <summary>One table row. The header flag marks a repeating header row in Word.</summary>
public sealed record TableRow
{
    /// <summary>Cells in column order. Required, at least one.</summary>
    public required IReadOnlyList<TableCell> Cells { get; init; }

    /// <summary>True when this row is a repeating header row.</summary>
    public bool IsHeader { get; init; }
}

/// <summary>One table cell: text content plus optional fill and alignment.</summary>
public sealed record TableCell
{
    /// <summary>Cell text content (may be empty for a blank cell).</summary>
    public TextModel? Content { get; init; }

    /// <summary>Cell fill: palette token name or #RRGGBB literal.</summary>
    public string? Fill { get; init; }

    /// <summary>Cell content alignment. Null = inherited.</summary>
    public TextAlignment? Alignment { get; init; }
}
