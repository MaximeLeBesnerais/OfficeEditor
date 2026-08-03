namespace XlsxEditor.Core.Instructions;

/// <summary>
/// The outcome of preflighting an instruction set: the immutable, fully-resolved plan
/// (null when validation or resolution found errors) plus every diagnostic collected
/// along the way.
/// </summary>
public sealed record XlsxPlanResult(XlsxPlan? Plan, XlsxValidationResult Validation)
{
    public bool IsValid => Validation.IsValid;
}

/// <summary>
/// Immutable preflight/execution plan for an <see cref="XlsxInstructionSet"/>. Produced by
/// <see cref="XlsxPlanner"/> before any mutation: every variable is resolved
/// deterministically (with cycle detection), every cell address/range is normalized,
/// every cell type is resolved, and every style reference is validated. The plan is pure
/// data — it depends on no builder implementation, so the executor (which maps it onto
/// workbook primitives) can land independently.
/// </summary>
public sealed record XlsxPlan
{
    public string Version { get; init; } = string.Empty;

    public string? Description { get; init; }

    public WorkbookMetadata? Metadata { get; init; }

    /// <summary>Validated named styles, in declaration order.</summary>
    public IReadOnlyList<NamedStyle> Styles { get; init; } = Array.Empty<NamedStyle>();

    public IReadOnlyList<XlsxPlanWorksheet> Worksheets { get; init; } = Array.Empty<XlsxPlanWorksheet>();

    /// <summary>Fully-resolved variables (no {{…}} placeholders remain).</summary>
    public IReadOnlyDictionary<string, string> Variables { get; init; } = new Dictionary<string, string>();
}

public sealed record XlsxPlanWorksheet
{
    public string Name { get; init; } = string.Empty;

    /// <summary>1-based row the header row (if any) is written to.</summary>
    public int StartRow { get; init; } = 1;

    /// <summary>Named style (or legacy numeric style id) applied to the header row.</summary>
    public string? HeaderStyle { get; init; }

    public IReadOnlyList<XlsxPlanColumn> Columns { get; init; } = Array.Empty<XlsxPlanColumn>();

    /// <summary>Resolved header cells at <see cref="StartRow"/>.</summary>
    public IReadOnlyList<XlsxPlanCell> Headers { get; init; } = Array.Empty<XlsxPlanCell>();

    /// <summary>Resolved data rows; the row that follows the header row is index 0.</summary>
    public IReadOnlyList<XlsxPlanRow> Rows { get; init; } = Array.Empty<XlsxPlanRow>();

    /// <summary>Resolved discrete cells (address-pinned, unaffected by <see cref="StartRow"/>).</summary>
    public IReadOnlyList<XlsxPlanCell> Cells { get; init; } = Array.Empty<XlsxPlanCell>();

    /// <summary>Explicit row heights keyed by 1-based row index (points).</summary>
    public IReadOnlyDictionary<int, double> RowHeights { get; init; } = new Dictionary<int, double>();

    /// <summary>Normalized merged ranges in A1 notation (e.g. "A1:C3").</summary>
    public IReadOnlyList<string> Merges { get; init; } = Array.Empty<string>();

    public XlsxPlanFreezePanes? FreezePanes { get; init; }

    /// <summary>Standalone autofilter range in A1 notation, or null.</summary>
    public string? AutoFilterRange { get; init; }

    public IReadOnlyList<XlsxPlanTable> Tables { get; init; } = Array.Empty<XlsxPlanTable>();
}

/// <summary>A positional typed column definition; column A is index 0.</summary>
public sealed record XlsxPlanColumn
{
    /// <summary>0-based index; column A = 0.</summary>
    public int Index { get; init; }

    /// <summary>1-based column number; column A = 1.</summary>
    public int ColumnNumber => Index + 1;

    /// <summary>Column letter ("A", "AB", "XFD").</summary>
    public string ColumnLetter { get; init; } = string.Empty;

    public string? Name { get; init; }

    public double? Width { get; init; }

    public string? StyleRef { get; init; }

    /// <summary>Default cell type for cells in this column.</summary>
    public XlsxCellType Type { get; init; } = XlsxCellType.Auto;
}

public sealed record XlsxPlanRow
{
    /// <summary>1-based row index in the worksheet.</summary>
    public int RowIndex { get; init; }

    public IReadOnlyList<XlsxPlanCell> Cells { get; init; } = Array.Empty<XlsxPlanCell>();
}

public sealed record XlsxPlanCell
{
    /// <summary>Normalized A1 reference (e.g. "A1").</summary>
    public string Reference { get; init; } = string.Empty;

    public int Row { get; init; }

    public int Column { get; init; }

    /// <summary>Resolved literal value, or null for formula-only cells.</summary>
    public string? Value { get; init; }

    /// <summary>Resolved formula in display form (leading '='), or null.</summary>
    public string? Formula { get; init; }

    public XlsxCellType Type { get; init; } = XlsxCellType.Auto;

    /// <summary>Named style reference or legacy numeric style-id string.</summary>
    public string? StyleRef { get; init; }

    public string? NumberFormat { get; init; }

    /// <summary>JSON-style path of the source instruction (for diagnostics).</summary>
    public string Source { get; init; } = string.Empty;
}

public sealed record XlsxPlanFreezePanes
{
    /// <summary>Rows 1..Row-1 are frozen. 1-based.</summary>
    public int Row { get; init; } = 1;

    /// <summary>Columns 1..Column-1 are frozen. 1-based.</summary>
    public int Column { get; init; } = 1;
}

public sealed record XlsxPlanTable
{
    public string Name { get; init; } = string.Empty;

    /// <summary>Normalized range in A1 notation (e.g. "A1:D20").</summary>
    public string Range { get; init; } = string.Empty;
}
