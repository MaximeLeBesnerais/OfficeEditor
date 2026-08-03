using XlsxEditor.Core.Builders;
using XlsxEditor.Core.Exceptions;
using XlsxEditor.Core.Styles;

namespace XlsxEditor.Core.Instructions;

/// <summary>
/// Structured validation of an <see cref="XlsxInstructionSet"/>. Unlike the throwing
/// <see cref="XlsxInstructionValidator"/>, this engine never throws for bad input: it
/// aggregates every finding (including warnings such as duplicate row-height entries)
/// into a <see cref="XlsxValidationResult"/> with path-qualified diagnostics. The
/// throwing validator/parser are thin adapters over this engine.
/// </summary>
public static class XlsxValidationEngine
{
    private static readonly char[] IllegalSheetNameChars = { ':', '\\', '/', '?', '*', '[', ']' };

    private static readonly System.Text.RegularExpressions.Regex CellAddressSyntax =
        new(@"^[A-Za-z]{1,3}[1-9][0-9]*$", System.Text.RegularExpressions.RegexOptions.Compiled);

    public static XlsxValidationResult Validate(XlsxInstructionSet? set)
    {
        var diagnostics = new List<XlsxDiagnostic>();

        if (set is null)
        {
            return Error(new XlsxDiagnostic(
                XlsxDiagnosticCode.InstructionSetNull,
                XlsxDiagnosticSeverity.Error,
                string.Empty,
                "Instruction set must not be null."));
        }

        ValidateRoot(set, diagnostics);

        if (set.Worksheets is { Count: > 0 })
        {
            var seenNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < set.Worksheets.Count; i++)
            {
                ValidateWorksheet(set, set.Worksheets[i], i, seenNames, diagnostics);
            }

            ValidateTableNames(set, diagnostics);
            ValidateMergeOverlaps(set, diagnostics);
            ValidateTableOverlaps(set, diagnostics);
            ValidateNamedStyles(set, diagnostics);
        }

        return new XlsxValidationResult { Diagnostics = diagnostics };
    }

    private static XlsxValidationResult Error(XlsxDiagnostic diagnostic) =>
        new() { Diagnostics = new[] { diagnostic } };

    private static void Add(
        List<XlsxDiagnostic> diagnostics, XlsxDiagnosticCode code, XlsxDiagnosticSeverity severity,
        string path, string message) =>
        diagnostics.Add(new XlsxDiagnostic(code, severity, path, message));

    private static void Error(List<XlsxDiagnostic> diagnostics, XlsxDiagnosticCode code, string path, string message) =>
        Add(diagnostics, code, XlsxDiagnosticSeverity.Error, path, message);

    private static void Warning(List<XlsxDiagnostic> diagnostics, XlsxDiagnosticCode code, string path, string message) =>
        Add(diagnostics, code, XlsxDiagnosticSeverity.Warning, path, message);

    // ─── Root ─────────────────────────────────────────────────────

    private static void ValidateRoot(XlsxInstructionSet set, List<XlsxDiagnostic> diagnostics)
    {
        if (string.IsNullOrWhiteSpace(set.Version))
        {
            Error(diagnostics, XlsxDiagnosticCode.VersionRequired, "version",
                "Instruction set must declare a 'version'.");
        }
        else if (set.Version != "1.0")
        {
            Error(diagnostics, XlsxDiagnosticCode.UnsupportedVersion, "version",
                $"Unsupported instruction version '{set.Version}'. Only version '1.0' is supported.");
        }

        if (set.Worksheets is null || set.Worksheets.Count == 0)
        {
            Error(diagnostics, XlsxDiagnosticCode.MissingWorksheets, "worksheets",
                "Instruction set must contain at least one worksheet in 'worksheets'.");
        }

        if (set.Variables is { Count: 0 })
        {
            Error(diagnostics, XlsxDiagnosticCode.EmptyVariablesBlock, "variables",
                "'variables' block is present but empty. Either remove it or provide values.");
        }

        if (set.Variables != null)
        {
            foreach (var pair in set.Variables)
            {
                if (pair.Key is null)
                {
                    Error(diagnostics, XlsxDiagnosticCode.NullVariableName, "variables",
                        "'variables' block contains a null variable name; variable names must be non-null strings.");
                }
                else if (pair.Value is null)
                {
                    Error(diagnostics, XlsxDiagnosticCode.NullVariableValue, $"variables.{pair.Key}",
                        $"Variable '{pair.Key}' has a null value; variable values must be strings. " +
                        "Use an empty string to clear a value.");
                }
            }
        }
    }

    // ─── Named styles ─────────────────────────────────────────────

    private static void ValidateNamedStyles(XlsxInstructionSet set, List<XlsxDiagnostic> diagnostics)
    {
        if (set.Styles is null)
        {
            return;
        }

        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < set.Styles.Count; i++)
        {
            var style = set.Styles[i];
            var path = $"styles[{i}]";
            if (style is null)
            {
                Error(diagnostics, XlsxDiagnosticCode.CellStyleUnknown, path,
                    "Every entry in 'styles' must be a style object.");
                continue;
            }

            if (string.IsNullOrWhiteSpace(style.Name))
            {
                Error(diagnostics, XlsxDiagnosticCode.CellStyleUnknown, $"{path}.name",
                    $"Style {i} has no 'name'; named styles must declare a unique name.");
            }
            else if (!names.Add(style.Name))
            {
                Error(diagnostics, XlsxDiagnosticCode.CellStyleUnknown, $"{path}.name",
                    $"Duplicate style name '{style.Name}'. Style names must be unique (case-insensitive).");
            }

            ValidateFont(style.Font, $"{path}.font", diagnostics);
            ValidateFill(style.Fill, $"{path}.fill", diagnostics);
            ValidateBorder(style.Border, $"{path}.border", diagnostics);
            ValidateAlignment(style.Alignment, $"{path}.alignment", diagnostics);
            ValidateNumberFormat(style.NumberFormat, $"{path}.numberFormat", diagnostics);
        }
    }

    private static void ValidateFont(FontStyleInstruction? font, string path, List<XlsxDiagnostic> diagnostics)
    {
        if (font?.Size is { } size && (!double.IsFinite(size) || size <= 0 || size > 409))
        {
            Error(diagnostics, XlsxDiagnosticCode.CellStyleUnknown, $"{path}.size",
                $"Font size {size} is outside Excel's supported range (greater than 0, up to 409 points).");
        }

        if (font?.Color is not null && !ExcelColor.IsValid(font.Color))
        {
            Error(diagnostics, XlsxDiagnosticCode.CellStyleUnknown, $"{path}.color",
                $"Invalid font color '{font.Color}'. Colors must be a 6-digit RGB (e.g. 'FF0000'), " +
                "an 8-digit ARGB (e.g. 'FFFF0000') hex string, or a common color name " +
                "(e.g. 'red', 'white', 'darkgray').");
        }
    }

    private static void ValidateFill(FillStyleInstruction? fill, string path, List<XlsxDiagnostic> diagnostics)
    {
        if (fill?.Pattern is { } pattern && !IsKnownFillPattern(pattern))
        {
            Error(diagnostics, XlsxDiagnosticCode.CellStyleUnknown, $"{path}.pattern",
                $"Unknown fill pattern '{pattern}'. Real Excel pattern types include 'none', " +
                "'solid', 'gray125', 'gray0625', 'darkgray', 'mediumgray', 'lightgray', " +
                "'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', " +
                "'darkTrellis', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', " +
                "'lightGrid', 'lightTrellis'.");
        }

        if (fill is { Pattern: not null } && !IsKnownFillPattern(fill.Pattern))
        {
            return;
        }

        var patternType = fill?.Pattern?.Trim().ToLowerInvariant();
        if (patternType == "solid" && string.IsNullOrWhiteSpace(fill!.Color))
        {
            Error(diagnostics, XlsxDiagnosticCode.CellStyleUnknown, $"{path}.color",
                "A 'solid' fill requires a 'color'; without one there is nothing to fill the " +
                "cell with. Remove the fill or provide a color.");
        }

        if (fill?.Color is not null && !ExcelColor.IsValid(fill.Color))
        {
            Error(diagnostics, XlsxDiagnosticCode.CellStyleUnknown, $"{path}.color",
                $"Invalid fill color '{fill.Color}'. Colors must be a 6-digit RGB (e.g. 'FF0000'), " +
                "an 8-digit ARGB (e.g. 'FFFF0000') hex string, or a common color name " +
                "(e.g. 'red', 'white', 'darkgray').");
        }
    }

    private static bool IsKnownFillPattern(string pattern)
    {
        switch (pattern.Trim().ToLowerInvariant())
        {
            case "none":
            case "solid":
            case "darkgray":
            case "mediumgray":
            case "lightgray":
            case "gray125":
            case "gray0625":
            case "darkvertical":
            case "darkhorizontal":
            case "darkdown":
            case "darkup":
            case "darkgrid":
            case "darktrellis":
            case "lightvertical":
            case "lighthorizontal":
            case "lightdown":
            case "lightup":
            case "lightgrid":
            case "lighttrellis":
                return true;
            default:
                return false;
        }
    }

    private static void ValidateBorder(BorderStyleInstruction? border, string path, List<XlsxDiagnostic> diagnostics)
    {
        ValidateBorderEdge(border?.Left, $"{path}.left", diagnostics);
        ValidateBorderEdge(border?.Right, $"{path}.right", diagnostics);
        ValidateBorderEdge(border?.Top, $"{path}.top", diagnostics);
        ValidateBorderEdge(border?.Bottom, $"{path}.bottom", diagnostics);
    }

    private static void ValidateBorderEdge(BorderEdgeInstruction? edge, string path, List<XlsxDiagnostic> diagnostics)
    {
        if (edge?.Style is { } style && !IsKnownBorderStyle(style))
        {
            Error(diagnostics, XlsxDiagnosticCode.CellStyleUnknown, $"{path}.style",
                $"Unknown border style '{style}'. Known styles: none, thin, medium, thick, double, " +
                "dashed, dotted, dashDot, dashDotDot, hair, mediumDashed, mediumDashDot, " +
                "mediumDashDotDot, slantDashDot.");
        }

        if (edge?.Color is not null
            && (edge.Style is null || edge.Style.Trim().ToLowerInvariant() is "" or "none"))
        {
            // A color on an edgeless side would be silently dropped at generation; reject it
            // so an accepted border always produces the edges it declares.
            Error(diagnostics, XlsxDiagnosticCode.CellStyleUnknown, $"{path}.color",
                $"Border edge has a color '{edge.Color}' but no border style; a border edge " +
                "needs a style (e.g. 'thin', 'medium', 'dashed', 'dotted') to render. Remove " +
                "the color or add a style.");
        }

        if (edge?.Color is not null && !ExcelColor.IsValid(edge.Color))
        {
            Error(diagnostics, XlsxDiagnosticCode.CellStyleUnknown, $"{path}.color",
                $"Invalid border color '{edge.Color}'. Colors must be a 6-digit RGB (e.g. " +
                "'FF0000'), an 8-digit ARGB (e.g. 'FFFF0000') hex string, or a common color " +
                "name (e.g. 'red', 'white', 'darkgray').");
        }
    }

    private static bool IsKnownBorderStyle(string style)
    {
        switch (style.Trim().ToLowerInvariant())
        {
            case "none":
            case "thin":
            case "medium":
            case "thick":
            case "double":
            case "dashed":
            case "dotted":
            case "dashdot":
            case "dashdotdot":
            case "hair":
            case "mediumdashed":
            case "mediumdashdot":
            case "mediumdashdotdot":
            case "slantdashdot":
                return true;
            default:
                return false;
        }
    }

    private static void ValidateAlignment(AlignmentStyleInstruction? alignment, string path, List<XlsxDiagnostic> diagnostics)
    {
        if (alignment?.Horizontal is { } horizontal && !IsKnownHorizontalAlignment(horizontal))
        {
            Error(diagnostics, XlsxDiagnosticCode.CellStyleUnknown, $"{path}.horizontal",
                $"Unknown horizontal alignment '{horizontal}'. Known values: general, left, center, right, fill, justify, centerContinuous.");
        }

        if (alignment?.Vertical is { } vertical && !IsKnownVerticalAlignment(vertical))
        {
            Error(diagnostics, XlsxDiagnosticCode.CellStyleUnknown, $"{path}.vertical",
                $"Unknown vertical alignment '{vertical}'. Known values: top, center, bottom, justify, distributed.");
        }
    }

    private static bool IsKnownHorizontalAlignment(string value) =>
        value.Trim().ToLowerInvariant() is "general" or "left" or "center" or "right" or "fill" or "justify" or "centercontinuous";

    private static bool IsKnownVerticalAlignment(string value) =>
        value.Trim().ToLowerInvariant() is "top" or "center" or "bottom" or "justify" or "distributed";

    // ─── Worksheets ───────────────────────────────────────────────

    private static void ValidateWorksheet(
        XlsxInstructionSet set, WorksheetInstruction ws, int sheetIndex, HashSet<string> seenNames,
        List<XlsxDiagnostic> diagnostics)
    {
        var path = $"worksheets[{sheetIndex}]";
        if (ws is null)
        {
            Error(diagnostics, XlsxDiagnosticCode.SheetNameRequired, path,
                "Every entry in 'worksheets' must be a worksheet object.");
            return;
        }

        var sheetPath = $"{path}.name";
        if (string.IsNullOrWhiteSpace(ws.Name))
        {
            Error(diagnostics, XlsxDiagnosticCode.SheetNameRequired, sheetPath,
                "Every worksheet must have a non-empty 'name'.");
        }
        else
        {
            if (ws.Name.Length > 31)
            {
                Error(diagnostics, XlsxDiagnosticCode.SheetNameTooLong, sheetPath,
                    $"Worksheet name '{ws.Name}' exceeds Excel's 31-character limit.");
            }

            if (ws.Name.IndexOfAny(IllegalSheetNameChars) >= 0)
            {
                Error(diagnostics, XlsxDiagnosticCode.SheetNameIllegalChars, sheetPath,
                    $"Worksheet name '{ws.Name}' contains illegal characters (: \\ / ? * [ ]).");
            }

            if (ws.Name[0] == '\'' || ws.Name[^1] == '\'')
            {
                Error(diagnostics, XlsxDiagnosticCode.SheetNameApostrophe, sheetPath,
                    $"Worksheet name '{ws.Name}' begins or ends with an apostrophe ('), " +
                    "which Excel does not allow. Remove the leading or trailing apostrophe.");
            }

            if (!seenNames.Add(ws.Name))
            {
                Error(diagnostics, XlsxDiagnosticCode.SheetNameDuplicate, sheetPath,
                    $"Duplicate worksheet name '{ws.Name}'. Worksheet names must be unique (case-insensitive).");
            }
        }

        ValidateStartRow(ws, sheetIndex, diagnostics);
        ValidateHeaderStyle(set, ws, sheetIndex, diagnostics);
        ValidateHeaders(set, ws, sheetIndex, diagnostics);
        ValidateColumns(set, ws, sheetIndex, diagnostics);
        ValidateRows(ws, sheetIndex, diagnostics);
        ValidateRowHeights(ws, sheetIndex, diagnostics);
        ValidateCells(set, ws, sheetIndex, diagnostics);
        ValidateMerges(ws, sheetIndex, diagnostics);
        ValidateFreezePanes(ws, sheetIndex, diagnostics);
        ValidateAutoFilter(ws, sheetIndex, diagnostics);
        ValidateTables(set, ws, sheetIndex, diagnostics);
        ValidateSheetNotEmpty(ws, sheetIndex, diagnostics);
    }

    private static void ValidateStartRow(WorksheetInstruction ws, int sheetIndex, List<XlsxDiagnostic> diagnostics)
    {
        if (ws.StartRow is not { } startRow)
        {
            return;
        }

        if (startRow < 1 || startRow > XlsxRangeUtilities.MaxRows)
        {
            Error(diagnostics, XlsxDiagnosticCode.SheetStartRowOutOfBounds, $"worksheets[{sheetIndex}].startRow",
                $"Worksheet '{ws.Name}' has 'startRow' {startRow}, which is outside Excel's valid " +
                $"row range 1-{XlsxRangeUtilities.MaxRows:N0}.");
        }
    }

    private static void ValidateSheetNotEmpty(WorksheetInstruction ws, int sheetIndex, List<XlsxDiagnostic> diagnostics)
    {
        var hasContent = ws.Headers is { Count: > 0 }
            || ws.Rows is { Count: > 0 }
            || ws.Cells is { Count: > 0 }
            || ws.Columns is { Count: > 0 }
            || ws.RowHeights is { Count: > 0 }
            || ws.Merges is { Count: > 0 }
            || ws.FreezePanes is not null
            || ws.AutoFilter is not null
            || ws.Tables is { Count: > 0 };

        if (!hasContent)
        {
            Error(diagnostics, XlsxDiagnosticCode.SheetEmpty, $"worksheets[{sheetIndex}]",
                $"Worksheet '{ws.Name}' must have at least one of 'headers', 'rows', 'cells', " +
                "'columns', 'rowHeights', 'merges', 'freezePanes', 'autoFilter', or 'tables'.");
        }
    }

    private static void ValidateHeaderStyle(XlsxInstructionSet set, WorksheetInstruction ws, int sheetIndex, List<XlsxDiagnostic> diagnostics)
    {
        if (ws.HeaderStyle is null)
        {
            return;
        }

        if (!IsValidStyleReference(set, ws.HeaderStyle))
        {
            Error(diagnostics, XlsxDiagnosticCode.SheetHeaderStyleUnknown, $"worksheets[{sheetIndex}].headerStyle",
                $"Worksheet '{ws.Name}' references unknown style '{ws.HeaderStyle}'. Style must be a numeric " +
                "style id or the name of a defined style in 'styles'.");
        }
    }

    private static void ValidateHeaders(XlsxInstructionSet set, WorksheetInstruction ws, int sheetIndex, List<XlsxDiagnostic> diagnostics)
    {
        if (ws.Headers is null)
        {
            return;
        }

        var count = ws.Headers.Count;
        if (count > XlsxRangeUtilities.MaxColumns)
        {
            Error(diagnostics, XlsxDiagnosticCode.SheetHeadersExceedColumns, $"worksheets[{sheetIndex}].headers",
                $"Worksheet '{ws.Name}' declares {count} headers, but Excel only has " +
                $"{XlsxRangeUtilities.MaxColumns} columns (A-XFD).");
        }

        for (var i = 0; i < count; i++)
        {
            if (ws.Headers[i] is null)
            {
                Error(diagnostics, XlsxDiagnosticCode.HeaderNull, $"worksheets[{sheetIndex}].headers[{i}]",
                    $"Header {i + 1} of worksheet '{ws.Name}' is null; headers must be strings.");
            }
        }
    }

    private static void ValidateColumns(XlsxInstructionSet set, WorksheetInstruction ws, int sheetIndex, List<XlsxDiagnostic> diagnostics)
    {
        if (ws.Columns is null)
        {
            return;
        }

        if (ws.Columns.Count > XlsxRangeUtilities.MaxColumns)
        {
            Error(diagnostics, XlsxDiagnosticCode.ColumnCountTooLarge, $"worksheets[{sheetIndex}].columns",
                $"Worksheet '{ws.Name}' declares {ws.Columns.Count} columns, but Excel only has " +
                $"{XlsxRangeUtilities.MaxColumns} columns (A-XFD).");
        }

        for (var i = 0; i < ws.Columns.Count; i++)
        {
            var column = ws.Columns[i];
            var path = $"worksheets[{sheetIndex}].columns[{i}]";
            if (column is null)
            {
                Error(diagnostics, XlsxDiagnosticCode.ColumnStyleUnknown, path,
                    $"Column {i + 1} of worksheet '{ws.Name}' must be a column object.");
                continue;
            }

            if (column.Width is { } width)
            {
                if (!double.IsFinite(width))
                {
                    Error(diagnostics, XlsxDiagnosticCode.ColumnWidthInvalid, $"{path}.width",
                        $"Column width for column {i + 1} of sheet '{ws.Name}' must be a finite number; " +
                        "NaN and infinity are not valid Excel column widths.");
                }
                else if (width <= 0)
                {
                    Error(diagnostics, XlsxDiagnosticCode.ColumnWidthInvalid, $"{path}.width",
                        $"Column width for column {i + 1} of sheet '{ws.Name}' must be greater than zero " +
                        $"(Excel column-width units). Got {width}.");
                }
                else if (width > 255)
                {
                    Error(diagnostics, XlsxDiagnosticCode.ColumnWidthInvalid, $"{path}.width",
                        $"Column width for column {i + 1} of sheet '{ws.Name}' must not exceed 255 " +
                        "(Excel's maximum column width). Got {width}.");
                }
            }

            if (column.Style is not null && !IsValidStyleReference(set, column.Style))
            {
                Error(diagnostics, XlsxDiagnosticCode.ColumnStyleUnknown, $"{path}.style",
                    $"Column {i + 1} of sheet '{ws.Name}' references unknown style '{column.Style}'. " +
                    "Style must be a numeric style id or the name of a defined style in 'styles'.");
            }

            if (column.Type is not null && !XlsxCellTypeParser.TryParse(column.Type, out _))
            {
                Error(diagnostics, XlsxDiagnosticCode.ColumnTypeInvalid, $"{path}.type",
                    $"Column {i + 1} of sheet '{ws.Name}' has unknown 'type' '{column.Type}'. " +
                    "Expected one of: auto, string, number, boolean, date, datetime.");
            }
        }
    }

    private static void ValidateRows(WorksheetInstruction ws, int sheetIndex, List<XlsxDiagnostic> diagnostics)
    {
        if (ws.Rows is null)
        {
            return;
        }

        var startRow = ws.StartRow ?? 1;
        var headerOffset = ws.Headers is { Count: > 0 } ? 1 : 0;
        for (var i = 0; i < ws.Rows.Count; i++)
        {
            var row = ws.Rows[i];
            var path = $"worksheets[{sheetIndex}].rows[{i}]";
            if (row is null)
            {
                Error(diagnostics, XlsxDiagnosticCode.RowNull, path,
                    $"Row {i + 1} of worksheet '{ws.Name}' is null; each row must be an array of cell values.");
                continue;
            }

            var rowIndex = startRow + headerOffset + i;
            if (rowIndex > XlsxRangeUtilities.MaxRows)
            {
                Error(diagnostics, XlsxDiagnosticCode.SheetStartRowOutOfBounds, path,
                    $"Row {i + 1} of worksheet '{ws.Name}' would land at row {rowIndex:N0}, beyond Excel's " +
                    $"maximum row {XlsxRangeUtilities.MaxRows:N0}.");
            }

            // A row wider than Excel's sheet is rejected here (semantic validation) so the
            // planner never reaches the address math that would throw past column XFD.
            if (row.Count > XlsxRangeUtilities.MaxColumns)
            {
                Error(diagnostics, XlsxDiagnosticCode.RowTooManyCells, path,
                    $"Row {i + 1} of worksheet '{ws.Name}' has {row.Count} cells, but Excel only has " +
                    $"{XlsxRangeUtilities.MaxColumns} columns (A-XFD).");
            }

            for (var j = 0; j < row.Count; j++)
            {
                if (row[j] is null)
                {
                    Error(diagnostics, XlsxDiagnosticCode.RowCellNull, $"{path}[{j}]",
                        $"Cell {j + 1} of row {i + 1} in worksheet '{ws.Name}' is null; cell values must be strings.");
                }
            }
        }
    }

    private static void ValidateRowHeights(WorksheetInstruction ws, int sheetIndex, List<XlsxDiagnostic> diagnostics)
    {
        if (ws.RowHeights is null)
        {
            return;
        }

        var seenRows = new HashSet<int>();
        for (var i = 0; i < ws.RowHeights.Count; i++)
        {
            var entry = ws.RowHeights[i];
            var path = $"worksheets[{sheetIndex}].rowHeights[{i}]";
            if (entry is null)
            {
                Error(diagnostics, XlsxDiagnosticCode.RowHeightInvalid, path,
                    $"Row height entry {i + 1} of sheet '{ws.Name}' must be an object with 'row' and 'height'.");
                continue;
            }

            var rowValue = entry.Row;
            if (rowValue is null)
            {
                Error(diagnostics, XlsxDiagnosticCode.RowHeightRowOutOfBounds, $"{path}.row",
                    $"Row height entry {i + 1} of sheet '{ws.Name}' is missing a 'row' index.");
            }
            else
            {
                if (rowValue < 1 || rowValue > XlsxRangeUtilities.MaxRows)
                {
                    Error(diagnostics, XlsxDiagnosticCode.RowHeightRowOutOfBounds, $"{path}.row",
                        $"Row height for row {rowValue} of sheet '{ws.Name}' is outside Excel's valid row " +
                        $"range 1-{XlsxRangeUtilities.MaxRows:N0}.");
                }

                if (!seenRows.Add(rowValue.Value))
                {
                    Warning(diagnostics, XlsxDiagnosticCode.DuplicateRowHeight, $"{path}.row",
                        $"Row {rowValue} of sheet '{ws.Name}' has more than one explicit height; the last one wins.");
                }
            }

            var rowDisplay = rowValue?.ToString() ?? "<unspecified>";
            if (entry.Height is { } height)
            {
                if (!double.IsFinite(height))
                {
                    Error(diagnostics, XlsxDiagnosticCode.RowHeightInvalid, $"{path}.height",
                        $"Row height for row {rowDisplay} of sheet '{ws.Name}' must be a finite number; NaN and " +
                        "infinity are not valid Excel row heights.");
                }
                else if (height <= 0)
                {
                    Error(diagnostics, XlsxDiagnosticCode.RowHeightInvalid, $"{path}.height",
                        $"Row height for row {rowDisplay} of sheet '{ws.Name}' must be greater than zero (points). " +
                        $"Got {height}.");
                }
                else if (height > 409.5)
                {
                    Error(diagnostics, XlsxDiagnosticCode.RowHeightInvalid, $"{path}.height",
                        $"Row height for row {rowDisplay} of sheet '{ws.Name}' must not exceed 409.5 points " +
                        "(Excel's maximum row height). Got {height}.");
                }
            }
        }
    }

    private static void ValidateCells(XlsxInstructionSet set, WorksheetInstruction ws, int sheetIndex, List<XlsxDiagnostic> diagnostics)
    {
        if (ws.Cells is null)
        {
            return;
        }

        for (var i = 0; i < ws.Cells.Count; i++)
        {
            var cell = ws.Cells[i];
            var path = $"worksheets[{sheetIndex}].cells[{i}]";
            if (cell is null)
            {
                Error(diagnostics, XlsxDiagnosticCode.CellNull, path,
                    $"Cell instruction in sheet '{ws.Name}' must be a cell object.");
                continue;
            }

            ValidateCellInstruction(set, ws, cell, sheetIndex, i, diagnostics);
        }
    }

    private static void ValidateCellInstruction(
        XlsxInstructionSet set, WorksheetInstruction ws, CellInstruction cell, int sheetIndex, int cellIndex,
        List<XlsxDiagnostic> diagnostics)
    {
        var path = $"worksheets[{sheetIndex}].cells[{cellIndex}]";

        if (string.IsNullOrWhiteSpace(cell.Address))
        {
            Error(diagnostics, XlsxDiagnosticCode.CellAddressRequired, $"{path}.address",
                $"Cell instruction in sheet '{ws.Name}' is missing 'address'.");
        }
        else
        {
            // Distinguish malformed references ("A", "1A") from syntactically valid but
            // out-of-bounds ones ("XFE1", "A1048577"): both are errors, but with distinct
            // codes and messages.
            if (!CellAddressSyntax.IsMatch(cell.Address))
            {
                Error(diagnostics, XlsxDiagnosticCode.CellAddressInvalid, $"{path}.address",
                    $"Invalid cell address '{cell.Address}' in sheet '{ws.Name}'. " +
                    "Expected a valid Excel reference like 'A1', 'AA10', etc.");
            }
            else
            {
                try
                {
                    WorksheetBuilder.NormalizeCellReference(cell.Address);
                }
                catch (XlsxException)
                {
                    Error(diagnostics, XlsxDiagnosticCode.CellAddressOutOfBounds, $"{path}.address",
                        $"Cell address '{cell.Address}' in sheet '{ws.Name}' is outside Excel's real " +
                        "sheet bounds (columns A-XFD, rows 1-1,048,576).");
                }
            }
        }

        var hasValue = cell.Value is not null;
        var hasFormula = cell.Formula is not null;

        if (hasValue && hasFormula)
        {
            Error(diagnostics, XlsxDiagnosticCode.CellValueAndFormula, path,
                $"Cell '{cell.Address}' in sheet '{ws.Name}' has both 'value' and 'formula'. " +
                "A cell must use one or the other, not both.");
        }
        else if (!hasValue && !hasFormula)
        {
            Error(diagnostics, XlsxDiagnosticCode.CellNoValueOrFormula, path,
                $"Cell '{cell.Address}' in sheet '{ws.Name}' has neither 'value' nor 'formula'. " +
                "Set one of them.");
        }

        if (hasFormula && !cell.Formula!.StartsWith("="))
        {
            Error(diagnostics, XlsxDiagnosticCode.FormulaMustStartWithEquals, $"{path}.formula",
                $"Formula in cell '{cell.Address}' sheet '{ws.Name}' must start with '='. " +
                $"Got: '{cell.Formula}'");
        }

        if (cell.Type is not null && !XlsxCellTypeParser.TryParse(cell.Type, out _))
        {
            Error(diagnostics, XlsxDiagnosticCode.CellTypeInvalid, $"{path}.type",
                $"Cell '{cell.Address}' in sheet '{ws.Name}' has unknown 'type' '{cell.Type}'. " +
                "Expected one of: auto, string, number, boolean, date, datetime.");
        }

        if (cell.Style is not null && !IsValidStyleReference(set, cell.Style))
        {
            Error(diagnostics, XlsxDiagnosticCode.CellStyleUnknown, $"{path}.style",
                $"Cell '{cell.Address}' in sheet '{ws.Name}' references unknown style '{cell.Style}'. " +
                "Style must be a numeric style id or the name of a defined style in 'styles'.");
        }

        ValidateNumberFormat(cell.NumberFormat, $"{path}.numberFormat", diagnostics);
    }

    private static void ValidateNumberFormat(string? numberFormat, string path, List<XlsxDiagnostic> diagnostics)
    {
        if (numberFormat is null)
        {
            return;
        }

        var trimmed = numberFormat.Trim();
        if (trimmed.Length == 0)
        {
            Error(diagnostics, XlsxDiagnosticCode.CellNumberFormatInvalid, path,
                "'numberFormat' must not be empty or whitespace; remove the field to inherit the default format.");
        }
        else if (trimmed.Length > 250)
        {
            Error(diagnostics, XlsxDiagnosticCode.CellNumberFormatInvalid, path,
                "'numberFormat' must not exceed 250 characters (Excel's number-format code limit).");
        }
    }

    /// <summary>
    /// A style reference is valid when it is either a legacy numeric style id (an
    /// unsigned integer string) or the name of a style declared in the set's
    /// <c>styles</c> block. Numeric ids are not range-checked here — only the workbook's
    /// stylesheet knows how many cell formats exist, which is an executor concern.
    /// </summary>
    private static bool IsValidStyleReference(XlsxInstructionSet set, string style) =>
        uint.TryParse(style, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out _)
        || (set.Styles?.Any(s => s is not null && string.Equals(s.Name, style, StringComparison.OrdinalIgnoreCase)) ?? false);

    // ─── Merges / freeze / autofilter / tables ────────────────────

    private static void ValidateMerges(WorksheetInstruction ws, int sheetIndex, List<XlsxDiagnostic> diagnostics)
    {
        if (ws.Merges is null)
        {
            return;
        }

        for (var i = 0; i < ws.Merges.Count; i++)
        {
            var range = ws.Merges[i];
            var path = $"worksheets[{sheetIndex}].merges[{i}]";
            if (range is null)
            {
                Error(diagnostics, XlsxDiagnosticCode.MergeRangeInvalid, path,
                    $"Merge {i + 1} of sheet '{ws.Name}' must be a range string.");
                continue;
            }

            var parsed = TryNormalizeRange(range);
            if (parsed.error is not null)
            {
                Error(diagnostics, XlsxDiagnosticCode.MergeRangeInvalid, path,
                    parsed.error.Message);
                continue;
            }

            var (start, end, _) = parsed;
            var startRow = WorksheetBuilder.GetRowIndex(start);
            var startCol = WorksheetBuilder.GetColumnIndex(start);
            var endRow = WorksheetBuilder.GetRowIndex(end);
            var endCol = WorksheetBuilder.GetColumnIndex(end);

            if (startRow == endRow && startCol == endCol)
            {
                Error(diagnostics, XlsxDiagnosticCode.MergeSingleCell, path,
                    $"Merge range '{range}' in sheet '{ws.Name}' is a single cell. " +
                    "Excel only merges a range of two or more cells.");
            }
            else if (endRow < startRow || endCol < startCol)
            {
                Error(diagnostics, XlsxDiagnosticCode.MergeReversed, path,
                    $"Merge range '{range}' in sheet '{ws.Name}' is reversed: the start cell must be the " +
                    "top-left corner of the merge and the end cell the bottom-right corner.");
            }
        }
    }

    private static void ValidateFreezePanes(WorksheetInstruction ws, int sheetIndex, List<XlsxDiagnostic> diagnostics)
    {
        var freeze = ws.FreezePanes;
        if (freeze is null)
        {
            return;
        }

        var path = $"worksheets[{sheetIndex}].freezePanes";
        if (freeze.Cell is not null)
        {
            if (freeze.Row is not null || freeze.Column is not null)
            {
                Error(diagnostics, XlsxDiagnosticCode.FreezePanesInvalid, path,
                    $"Worksheet '{ws.Name}' sets both 'cell' and 'row'/'column' in 'freezePanes'. " +
                    "Use one form or the other.");
                return;
            }

            var parsed = TryNormalizeCell(freeze.Cell);
            if (parsed.error is not null)
            {
                Error(diagnostics, XlsxDiagnosticCode.FreezePanesOutOfBounds, $"{path}.cell",
                    $"Worksheet '{ws.Name}' has an invalid freeze 'cell' '{freeze.Cell}'. " +
                    "Expected an A1-style reference such as 'A2'.");
            }

            return;
        }

        var row = freeze.Row ?? 1;
        var column = freeze.Column ?? 1;
        if (row < 1 || row > XlsxRangeUtilities.MaxRows)
        {
            Error(diagnostics, XlsxDiagnosticCode.FreezePanesOutOfBounds, $"{path}.row",
                $"Worksheet '{ws.Name}' freeze 'row' {row} is outside Excel's valid row range " +
                $"1-{XlsxRangeUtilities.MaxRows:N0}.");
        }

        if (column < 1 || column > XlsxRangeUtilities.MaxColumns)
        {
            Error(diagnostics, XlsxDiagnosticCode.FreezePanesOutOfBounds, $"{path}.column",
                $"Worksheet '{ws.Name}' freeze 'column' {column} is outside Excel's valid column " +
                $"range 1-{XlsxRangeUtilities.MaxColumns} (A-XFD).");
        }
    }

    private static void ValidateAutoFilter(WorksheetInstruction ws, int sheetIndex, List<XlsxDiagnostic> diagnostics)
    {
        if (ws.AutoFilter is null)
        {
            return;
        }

        var parsed = TryNormalizeRange(ws.AutoFilter);
        if (parsed.error is not null)
        {
            Error(diagnostics, XlsxDiagnosticCode.AutoFilterInvalid, $"worksheets[{sheetIndex}].autoFilter",
                $"Worksheet '{ws.Name}' has an invalid 'autoFilter' range '{ws.AutoFilter}'. " +
                "Expected an A1-style range such as 'A1:D20'.");
        }
    }

    private static void ValidateTables(XlsxInstructionSet set, WorksheetInstruction ws, int sheetIndex, List<XlsxDiagnostic> diagnostics)
    {
        if (ws.Tables is null)
        {
            return;
        }

        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < ws.Tables.Count; i++)
        {
            var table = ws.Tables[i];
            var path = $"worksheets[{sheetIndex}].tables[{i}]";
            if (table is null)
            {
                Error(diagnostics, XlsxDiagnosticCode.TableNameRequired, path,
                    $"Table {i + 1} of sheet '{ws.Name}' must be a table object with 'name' and 'range'.");
                continue;
            }

            if (string.IsNullOrWhiteSpace(table.Name))
            {
                Error(diagnostics, XlsxDiagnosticCode.TableNameRequired, $"{path}.name",
                    $"Table {i + 1} of sheet '{ws.Name}' has no 'name'; tables must have a display name.");
            }
            else
            {
                if (!XlsxRangeUtilities.IsValidTableName(table.Name))
                {
                    Error(diagnostics, XlsxDiagnosticCode.TableNameInvalid, $"{path}.name",
                        $"Invalid table name '{table.Name}' in sheet '{ws.Name}'. Excel table names must " +
                        "start with a letter, underscore or backslash and contain only letters, digits, " +
                        "periods and underscores (no spaces or other special characters).");
                }
                else if (XlsxRangeUtilities.IsCellReferenceName(table.Name))
                {
                    Error(diagnostics, XlsxDiagnosticCode.TableNameInvalid, $"{path}.name",
                        $"Invalid table name '{table.Name}' in sheet '{ws.Name}'. Excel rejects table " +
                        "names that look like a cell reference (for example A1, BC12, XFD1048576 or R1C1).");
                }

                if (!names.Add(table.Name))
                {
                    Error(diagnostics, XlsxDiagnosticCode.TableDuplicateName, $"{path}.name",
                        $"Duplicate table name '{table.Name}' in sheet '{ws.Name}'. Table names must be " +
                        "unique within a worksheet (case-insensitive).");
                }
            }

            var parsed = TryNormalizeRange(table.Range);
            if (parsed.error is not null)
            {
                Error(diagnostics, XlsxDiagnosticCode.TableRangeInvalid, $"{path}.range",
                    $"Table '{table.Name}' in sheet '{ws.Name}' has an invalid range '{table.Range}'. " +
                    "Expected an A1-style range such as 'A1:D20'.");
                continue;
            }

            var (start, end, _) = parsed;
            if (WorksheetBuilder.GetRowIndex(end) < WorksheetBuilder.GetRowIndex(start)
                || WorksheetBuilder.GetColumnIndex(end) < WorksheetBuilder.GetColumnIndex(start))
            {
                Error(diagnostics, XlsxDiagnosticCode.TableRangeReversed, $"{path}.range",
                    $"Table '{table.Name}' in sheet '{ws.Name}' has a reversed range '{table.Range}': " +
                    "the start cell must be the top-left corner and the end cell the bottom-right corner.");
            }
        }
    }

    private static void ValidateTableNames(XlsxInstructionSet set, List<XlsxDiagnostic> diagnostics)
    {
        // Table display names are unique workbook-wide (case-insensitive), matching the
        // builder's RegisterTableName contract.
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var s = 0; s < set.Worksheets.Count; s++)
        {
            var ws = set.Worksheets[s];
            if (ws?.Tables is null)
            {
                continue;
            }

            for (var t = 0; t < ws.Tables.Count; t++)
            {
                var name = ws.Tables[t]?.Name;
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                if (!seen.Add(name))
                {
                    Error(diagnostics, XlsxDiagnosticCode.TableDuplicateName,
                        $"worksheets[{s}].tables[{t}].name",
                        $"Duplicate table name '{name}'. Table names must be unique workbook-wide " +
                        "(case-insensitive).");
                }
            }
        }
    }

    private static void ValidateMergeOverlaps(XlsxInstructionSet set, List<XlsxDiagnostic> diagnostics)
    {
        for (var s = 0; s < set.Worksheets.Count; s++)
        {
            var ws = set.Worksheets[s];
            if (ws?.Merges is not { Count: > 0 })
            {
                continue;
            }

            for (var i = 0; i < ws.Merges.Count; i++)
            {
                var parsed = TryNormalizeRange(ws.Merges[i]);
                if (parsed.error is not null)
                {
                    continue;
                }

                var (start, end, _) = parsed;
                var aStartRow = WorksheetBuilder.GetRowIndex(start);
                var aStartCol = WorksheetBuilder.GetColumnIndex(start);
                var aEndRow = WorksheetBuilder.GetRowIndex(end);
                var aEndCol = WorksheetBuilder.GetColumnIndex(end);

                for (var j = i + 1; j < ws.Merges.Count; j++)
                {
                    var other = TryNormalizeRange(ws.Merges[j]);
                    if (other.error is not null)
                    {
                        continue;
                    }

                    var (oStart, oEnd, _) = other;
                    var oStartRow = WorksheetBuilder.GetRowIndex(oStart);
                    var oStartCol = WorksheetBuilder.GetColumnIndex(oStart);
                    var oEndRow = WorksheetBuilder.GetRowIndex(oEnd);
                    var oEndCol = WorksheetBuilder.GetColumnIndex(oEnd);

                    if (aStartRow <= oEndRow && oStartRow <= aEndRow && aStartCol <= oEndCol && oStartCol <= aEndCol)
                    {
                        Error(diagnostics, XlsxDiagnosticCode.MergeOverlap,
                            $"worksheets[{s}].merges[{i}]",
                            $"Merge '{ws.Merges[i]}' in sheet '{ws.Name}' overlaps merge '{ws.Merges[j]}'. " +
                            "Excel does not allow overlapping merged ranges on the same worksheet.");
                    }
                }
            }
        }
    }

    private static void ValidateTableOverlaps(XlsxInstructionSet set, List<XlsxDiagnostic> diagnostics)
    {
        for (var s = 0; s < set.Worksheets.Count; s++)
        {
            var ws = set.Worksheets[s];
            if (ws?.Tables is not { Count: > 0 })
            {
                continue;
            }

            for (var i = 0; i < ws.Tables.Count; i++)
            {
                var table = ws.Tables[i];
                var tableParsed = TryNormalizeRange(table?.Range);
                if (tableParsed.error is not null)
                {
                    continue;
                }

                var (tStart, tEnd, _) = tableParsed;
                var tStartRow = WorksheetBuilder.GetRowIndex(tStart);
                var tStartCol = WorksheetBuilder.GetColumnIndex(tStart);
                var tEndRow = WorksheetBuilder.GetRowIndex(tEnd);
                var tEndCol = WorksheetBuilder.GetColumnIndex(tEnd);

                // Table vs table overlap.
                for (var j = i + 1; j < ws.Tables.Count; j++)
                {
                    var other = ws.Tables[j];
                    var otherParsed = TryNormalizeRange(other?.Range);
                    if (otherParsed.error is not null)
                    {
                        continue;
                    }

                    var (oStart, oEnd, _) = otherParsed;
                    var oStartRow = WorksheetBuilder.GetRowIndex(oStart);
                    var oStartCol = WorksheetBuilder.GetColumnIndex(oStart);
                    var oEndRow = WorksheetBuilder.GetRowIndex(oEnd);
                    var oEndCol = WorksheetBuilder.GetColumnIndex(oEnd);

                    if (tStartRow <= oEndRow && oStartRow <= tEndRow && tStartCol <= oEndCol && oStartCol <= tEndCol)
                    {
                        Error(diagnostics, XlsxDiagnosticCode.TableOverlap,
                            $"worksheets[{s}].tables[{i}]",
                            $"Table '{table!.Name}' in sheet '{ws.Name}' overlaps table '{other!.Name}'. " +
                            "Excel does not allow overlapping tables on the same worksheet.");
                    }
                }

                // Table vs merge overlap.
                if (ws.Merges is not null)
                {
                    for (var m = 0; m < ws.Merges.Count; m++)
                    {
                        var mergeParsed = TryNormalizeRange(ws.Merges[m]);
                        if (mergeParsed.error is not null)
                        {
                            continue;
                        }

                        var (mStart, mEnd, _) = mergeParsed;
                        var mStartRow = WorksheetBuilder.GetRowIndex(mStart);
                        var mStartCol = WorksheetBuilder.GetColumnIndex(mStart);
                        var mEndRow = WorksheetBuilder.GetRowIndex(mEnd);
                        var mEndCol = WorksheetBuilder.GetColumnIndex(mEnd);

                        if (tStartRow <= mEndRow && mStartRow <= tEndRow && tStartCol <= mEndCol && mStartCol <= tEndCol)
                        {
                            Error(diagnostics, XlsxDiagnosticCode.TableMergeOverlap,
                                $"worksheets[{s}].tables[{i}]",
                                $"Table '{table!.Name}' in sheet '{ws.Name}' overlaps merge '{ws.Merges[m]}'. " +
                                "Excel does not allow a table to overlap a merged range on the same worksheet.");
                        }
                    }
                }
            }
        }
    }

    // ─── Helpers ──────────────────────────────────────────────────

    private static (string start, string end, Exception? error) TryNormalizeRange(string? range)
    {
        if (range is null)
        {
            return (string.Empty, string.Empty, new XlsxException("Range must not be null."));
        }

        try
        {
            var (start, end) = XlsxRangeUtilities.NormalizeRange(range);
            return (start, end, null);
        }
        catch (Exception ex)
        {
            return (string.Empty, string.Empty, ex);
        }
    }

    private static (string cell, Exception? error) TryNormalizeCell(string? cell)
    {
        if (cell is null)
        {
            return (string.Empty, new XlsxException("Cell reference must not be null."));
        }

        try
        {
            return (WorksheetBuilder.NormalizeCellReference(cell), null);
        }
        catch (Exception ex)
        {
            return (string.Empty, ex);
        }
    }
}
