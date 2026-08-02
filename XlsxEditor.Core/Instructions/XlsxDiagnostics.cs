namespace XlsxEditor.Core.Instructions;

public enum XlsxDiagnosticSeverity
{
    Info,
    Warning,
    Error
}

/// <summary>
/// Stable identifiers for every diagnostic the validation engine and planner can emit.
/// The <see cref="XlsxDiagnostic.Code"/> is the enum member name, so consumers can switch
/// on it without string literals.
/// </summary>
public enum XlsxDiagnosticCode
{
    // Root / version / variables
    InvalidJson,
    InstructionSetNull,
    VersionRequired,
    UnsupportedVersion,
    MissingWorksheets,
    EmptyVariablesBlock,
    NullVariableName,
    NullVariableValue,
    UndefinedVariableReference,
    VariableCycle,

    // Worksheet
    SheetNameRequired,
    SheetNameTooLong,
    SheetNameIllegalChars,
    SheetNameApostrophe,
    SheetNameDuplicate,
    SheetEmpty,
    SheetStartRowOutOfBounds,
    SheetHeaderStyleUnknown,
    SheetHeadersExceedColumns,

    // Headers / rows / cells
    HeaderNull,
    RowNull,
    RowCellNull,
    CellNull,
    CellAddressRequired,
    CellAddressInvalid,
    CellAddressOutOfBounds,
    CellValueAndFormula,
    CellNoValueOrFormula,
    FormulaMustStartWithEquals,
    CellTypeInvalid,
    CellStyleUnknown,
    CellNumberFormatInvalid,
    DuplicateRowHeight,
    TextTooLong,
    FormulaTooLong,

    // Columns
    ColumnCountTooLarge,
    ColumnWidthInvalid,
    ColumnStyleUnknown,
    ColumnTypeInvalid,

    // Row heights
    RowHeightRowOutOfBounds,
    RowHeightInvalid,

    // Merges / freeze panes / autofilter / tables
    MergeRangeInvalid,
    MergeSingleCell,
    MergeReversed,
    MergeOverlap,
    FreezePanesInvalid,
    FreezePanesOutOfBounds,
    AutoFilterInvalid,
    TableNameRequired,
    TableNameInvalid,
    TableRangeInvalid,
    TableRangeReversed,
    TableDuplicateName,
    TableOverlap,
    TableMergeOverlap,

    // Planner
    DuplicateWrite,
    UnresolvedVariable
}

/// <summary>
/// A single structured validation/planning finding. <see cref="Path"/> is a stable,
/// JSON-style path into the instruction set (e.g. <c>worksheets[1].cells[0].address</c>)
/// so callers can highlight exactly where the problem is.
/// </summary>
public sealed record XlsxDiagnostic(
    XlsxDiagnosticCode Code,
    XlsxDiagnosticSeverity Severity,
    string Path,
    string Message)
{
    /// <summary>Stable machine-readable code, e.g. <c>SheetNameTooLong</c>.</summary>
    public string CodeName => Code.ToString();

    public bool IsError => Severity == XlsxDiagnosticSeverity.Error;

    public bool IsWarning => Severity == XlsxDiagnosticSeverity.Warning;
}

/// <summary>
/// The aggregate outcome of validation/planning: the full, ordered diagnostic list plus
/// convenience views. Never throws for bad input — all findings are collected.
/// </summary>
public sealed record XlsxValidationResult
{
    public static XlsxValidationResult Empty { get; } = new()
    {
        Diagnostics = Array.Empty<XlsxDiagnostic>()
    };

    public IReadOnlyList<XlsxDiagnostic> Diagnostics { get; init; } = Array.Empty<XlsxDiagnostic>();

    public bool IsValid => Diagnostics.All(d => !d.IsError);

    public bool HasErrors => Diagnostics.Any(d => d.IsError);

    public bool HasWarnings => Diagnostics.Any(d => d.IsWarning);

    public IEnumerable<XlsxDiagnostic> Errors => Diagnostics.Where(d => d.IsError);

    public IEnumerable<XlsxDiagnostic> Warnings => Diagnostics.Where(d => d.IsWarning);
}
