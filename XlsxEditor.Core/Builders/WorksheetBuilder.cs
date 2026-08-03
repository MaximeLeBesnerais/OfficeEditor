using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxEditor.Core.Exceptions;

namespace XlsxEditor.Core.Builders;

public class WorksheetBuilder : IWorksheetBuilder
{
    private readonly WorksheetPart _worksheetPart;
    private readonly Worksheet _worksheet;
    private readonly WorkbookBuilder _workbookBuilder;
    private readonly SheetData _sheetData;

    private static readonly Regex CellReferencePattern =
        new(@"^[A-Za-z]{1,3}[1-9][0-9]*$", RegexOptions.Compiled);

    // Excel's real sheet limits: columns A-XFD (1-based column 16384) and rows
    // 1-1,048,576. References past these corrupt nothing but are impossible to open.
    private const int MaxColumnNumber = 16384;
    private const int MaxRowNumber = 1_048_576;

    // Excel's UI limits for explicit sizes: a column width is measured in characters of
    // the default font and caps at 255; a row height is measured in points and caps at
    // 409.5. Larger values open but are silently clamped by Excel, so the builder rejects
    // them as non-Excel-compatible instead of writing a value Excel will not honor.
    private const double MaxColumnWidth = 255;
    private const double MaxRowHeight = 409.5;

    // Excel's 1900 date system counts days since 1899-12-30 (serial 1 = 1900-01-01).
    // Dates before 1900-01-01 have no serial and are rejected up front.
    private static readonly DateOnly ExcelDateEpoch = new(1899, 12, 30);
    private static readonly DateTime ExcelDateTimeEpoch = new(1899, 12, 30);

    // Default number formats so ISO date/datetime cells render as dates in Excel.
    internal const string DefaultDateNumberFormat = "yyyy-mm-dd";
    internal const string DefaultDateTimeNumberFormat = "yyyy-mm-dd h:mm:ss";

    public WorksheetPart WorksheetPart => _worksheetPart;

    public WorksheetBuilder(WorksheetPart worksheetPart, Worksheet worksheet, WorkbookBuilder workbookBuilder)
    {
        _worksheetPart = worksheetPart;
        _worksheet = worksheet;
        _workbookBuilder = workbookBuilder;
        _sheetData = worksheet.GetFirstChild<SheetData>() ?? new SheetData();
        if (worksheet.GetFirstChild<SheetData>() == null)
        {
            worksheet.Append(_sheetData);
        }
    }

    public IWorksheetBuilder AddCell(string cellReference, string value)
    {
        var cell = GetOrCreateCell(cellReference);

        // Overwriting replaces the cell's content entirely: a previous formula
        // (and its cached value) must not survive a value write.
        cell.CellFormula = null;

        // InvariantCulture: number detection must not depend on the machine's
        // locale (e.g. ',' as decimal separator). NaN/Infinity are rejected —
        // Excel cannot store them as numbers. NumberStyles.Float excludes
        // thousands separators, so "1,000" stays a string rather than losing
        // its formatting.
        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var numericValue)
            && !double.IsNaN(numericValue)
            && !double.IsInfinity(numericValue))
        {
            cell.CellValue = new CellValue(numericValue);
            cell.DataType = CellValues.Number;
        }
        else
        {
            var sharedStringIndex = _workbookBuilder.GetSharedStringIndex(value);
            cell.CellValue = new CellValue(sharedStringIndex.ToString());
            cell.DataType = CellValues.SharedString;
        }

        return this;
    }

    public IWorksheetBuilder AddCell(string cellReference, string formula, bool isFormula)
    {
        var cell = GetOrCreateCell(cellReference);

        if (isFormula)
        {
            // SpreadsheetML stores formula text WITHOUT the leading '='; '=' is
            // Excel's UI/input syntax only. Storing it in <f> triggers Excel's
            // repair prompt. Callers may pass either form.
            var formulaText = formula.StartsWith('=') ? formula[1..] : formula;
            cell.CellFormula = new CellFormula(formulaText);

            // A formula replaces any previous literal content; the old cached
            // value is stale until Excel recalculates, so drop it (and the old
            // data type) instead of letting GetCellValue return it.
            cell.CellValue = null;
            cell.DataType = null;
        }
        else
        {
            // Literal text goes through the shared string table: t="str" is
            // reserved for formula string results, not literal values.
            cell.CellFormula = null;
            var sharedStringIndex = _workbookBuilder.GetSharedStringIndex(formula);
            cell.CellValue = new CellValue(sharedStringIndex.ToString());
            cell.DataType = CellValues.SharedString;
        }

        return this;
    }

    public IWorksheetBuilder AddCell(string cellReference, string value, string styleId)
    {
        if (!uint.TryParse(styleId, NumberStyles.None, CultureInfo.InvariantCulture, out var parsedStyleId))
        {
            throw new XlsxException(
                $"Invalid styleId '{styleId}' for cell {cellReference}. " +
                "Style identifiers must be unsigned integers. " +
                "Use the Phase 3 style builder for named styles.");
        }

        // Validate BEFORE writing anything: an s= attribute pointing past the
        // stylesheet's cellXfs entries triggers Excel's repair prompt.
        var normalized = NormalizeCellReference(cellReference);
        _workbookBuilder.EnsureStyleIndexExists(parsedStyleId, normalized);

        AddCell(normalized, value);
        var cell = GetOrCreateCell(normalized);
        cell.StyleIndex = parsedStyleId;
        return this;
    }

    public IWorksheetBuilder AddHeaderRow(List<string> values, int rowIndex = 1)
    {
        // Validate the whole collection (null list, null elements) BEFORE any worksheet
        // or stylesheet mutation, so a bad header row can never leave a partially
        // written row behind.
        ValidateCellValues(values, nameof(values));

        var row = GetOrCreateRow(rowIndex);
        var headerStyleIndex = _workbookBuilder.EnsureHeaderStyleIndex();

        for (int i = 0; i < values.Count; i++)
        {
            var cellReference = GetCellReference(i, rowIndex);
            var cell = GetOrCreateCell(cellReference);
            cell.CellFormula = null;
            var sharedStringIndex = _workbookBuilder.GetSharedStringIndex(values[i]);
            cell.CellValue = new CellValue(sharedStringIndex.ToString());
            cell.DataType = CellValues.SharedString;

            cell.StyleIndex = headerStyleIndex;
        }

        return this;
    }

    public IWorksheetBuilder AddDataRow(List<string> values, int rowIndex)
    {
        ValidateCellValues(values, nameof(values));
        var row = GetOrCreateRow(rowIndex);

        for (int i = 0; i < values.Count; i++)
        {
            var cellReference = GetCellReference(i, rowIndex);
            AddCell(cellReference, values[i]);
        }

        return this;
    }

    public IWorksheetBuilder AddFormulaRow(List<string> formulas, int rowIndex)
    {
        // A null formula element is a programming error and must fail before the row
        // is created. An empty string is the documented "no formula in this column"
        // skip and stays supported.
        ValidateCellValues(formulas, nameof(formulas));
        var row = GetOrCreateRow(rowIndex);

        for (int i = 0; i < formulas.Count; i++)
        {
            var cellReference = GetCellReference(i, rowIndex);
            if (!string.IsNullOrEmpty(formulas[i]))
            {
                AddCell(cellReference, formulas[i], true);
            }
        }

        return this;
    }

    // ─── Typed cell writes ─────────────────────────────────────────

    /// <summary>
    /// Applies a raw cellXf style index to the cell at <paramref name="cellReference"/>,
    /// creating the cell if needed. Used by the instruction executor to preserve legacy
    /// numeric style ids that reference existing cell formats directly (the caller must
    /// have verified the index via <see cref="WorkbookBuilder.EnsureStyleIndexExists"/>).
    /// </summary>
    internal IWorksheetBuilder ApplyStyleIndex(string cellReference, uint styleIndex)
    {
        var normalized = NormalizeCellReference(cellReference);
        var cell = GetOrCreateCell(normalized);
        cell.StyleIndex = styleIndex;
        return this;
    }

    public IWorksheetBuilder AddCellString(string cellReference, string value, string? styleName = null)
    {
        ArgumentNullException.ThrowIfNull(value);

        var normalized = NormalizeCellReference(cellReference);
        var styleIndex = ResolveStyle(styleName, null);
        var cell = GetOrCreateCell(normalized);

        // Overwriting replaces the cell's content entirely: a previous formula (and its
        // cached value) must not survive a typed write.
        cell.CellFormula = null;
        var sharedStringIndex = _workbookBuilder.GetSharedStringIndex(value);
        cell.CellValue = new CellValue(sharedStringIndex.ToString());
        cell.DataType = CellValues.SharedString;
        ApplyStyle(cell, styleIndex);
        return this;
    }

    public IWorksheetBuilder AddCellNumber(string cellReference, double value, string? numberFormat = null, string? styleName = null)
    {
        // Excel stores IEEE doubles; NaN and infinities are not representable and would
        // corrupt the value. Reject them before any cell or stylesheet mutation.
        if (!double.IsFinite(value))
        {
            throw new XlsxException(
                $"Cell {cellReference} must hold a finite number; NaN and positive or " +
                "negative infinity are not valid Excel cell values.");
        }

        var normalized = NormalizeCellReference(cellReference);
        var styleIndex = ResolveStyle(styleName, numberFormat);
        var cell = GetOrCreateCell(normalized);
        cell.CellFormula = null;
        cell.CellValue = new CellValue(value);
        cell.DataType = CellValues.Number;
        ApplyStyle(cell, styleIndex);
        return this;
    }

    public IWorksheetBuilder AddCellBoolean(string cellReference, bool value, string? styleName = null)
    {
        var normalized = NormalizeCellReference(cellReference);
        var styleIndex = ResolveStyle(styleName, null);
        var cell = GetOrCreateCell(normalized);
        cell.CellFormula = null;
        cell.CellValue = new CellValue(value ? "1" : "0");
        cell.DataType = CellValues.Boolean;
        ApplyStyle(cell, styleIndex);
        return this;
    }

    public IWorksheetBuilder AddCellDate(string cellReference, string isoDate, string? numberFormat = null, string? styleName = null)
    {
        ArgumentNullException.ThrowIfNull(isoDate);

        if (!DateOnly.TryParseExact(
                isoDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
        {
            throw new XlsxException(
                $"Cell {cellReference} has an invalid ISO date '{isoDate}'. " +
                "Expected an ISO-8601 date in the form yyyy-MM-dd (for example 2024-01-15).");
        }

        if (date < new DateOnly(1900, 1, 1))
        {
            throw new XlsxException(
                $"Cell {cellReference} has a date before 1900-01-01 ('{isoDate}'), which " +
                "Excel's 1900 date system cannot represent.");
        }

        var serial = date.DayNumber - ExcelDateEpoch.DayNumber;
        return WriteDateCell(cellReference, serial, numberFormat ?? DefaultDateNumberFormat, styleName);
    }

    public IWorksheetBuilder AddCellDateTime(string cellReference, string isoDateTime, string? numberFormat = null, string? styleName = null)
    {
        ArgumentNullException.ThrowIfNull(isoDateTime);

        if (!DateTime.TryParseExact(
                isoDateTime,
                new[] { "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-ddTHH:mm:ss.fff" },
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out var dateTime))
        {
            throw new XlsxException(
                $"Cell {cellReference} has an invalid ISO datetime '{isoDateTime}'. " +
                "Expected an ISO-8601 datetime in the form yyyy-MM-ddTHH:mm:ss " +
                "(for example 2024-01-15T14:30:00, optional fractional seconds).");
        }

        if (dateTime < new DateTime(1900, 1, 1))
        {
            throw new XlsxException(
                $"Cell {cellReference} has a datetime before 1900-01-01 ('{isoDateTime}'), " +
                "which Excel's 1900 date system cannot represent.");
        }

        var serial = (dateTime - ExcelDateTimeEpoch).TotalDays;
        return WriteDateCell(cellReference, serial, numberFormat ?? DefaultDateTimeNumberFormat, styleName);
    }

    public IWorksheetBuilder AddFormula(string cellReference, string formula, string? styleName = null, string? numberFormat = null)
    {
        ArgumentNullException.ThrowIfNull(formula);

        var normalized = NormalizeCellReference(cellReference);

        // SpreadsheetML stores formula text WITHOUT the leading '='; '=' is Excel's
        // UI/input syntax only. Callers may pass either form.
        var formulaText = formula.StartsWith('=') ? formula[1..] : formula;
        if (string.IsNullOrWhiteSpace(formulaText))
        {
            throw new XlsxException(
                $"Cell {normalized} has an empty formula. " +
                "A formula must contain at least one token.");
        }

        // All validation runs before any mutation, so a rejected formula or style can
        // never leave a partially written (empty) cell behind.
        var styleIndex = ResolveStyle(styleName, numberFormat);
        var cell = GetOrCreateCell(normalized);

        cell.CellFormula = new CellFormula(formulaText);

        // A formula replaces any previous literal content; the old cached value is stale
        // until Excel recalculates, so drop it (and the old data type) instead of letting
        // GetCellValue return it.
        cell.CellValue = null;
        cell.DataType = null;
        ApplyStyle(cell, styleIndex);
        return this;
    }

    /// <summary>
    /// Resolves a named style plus optional number-format override to a cellXf index, or
    /// null when neither is given (the cell keeps its current style). All validation runs
    /// before any cell mutation so a bad style name fails loudly and atomically.
    /// </summary>
    private uint? ResolveStyle(string? styleName, string? numberFormat)
    {
        if (string.IsNullOrWhiteSpace(styleName) && string.IsNullOrWhiteSpace(numberFormat))
        {
            return null;
        }

        return _workbookBuilder.ResolveStyleIndex(styleName, numberFormat);
    }

    private static void ApplyStyle(Cell cell, uint? styleIndex)
    {
        if (styleIndex is { } index)
        {
            cell.StyleIndex = index;
        }
    }

    private IWorksheetBuilder WriteDateCell(string cellReference, double serial, string numberFormat, string? styleName)
    {
        var normalized = NormalizeCellReference(cellReference);
        var styleIndex = ResolveStyle(styleName, numberFormat);
        var cell = GetOrCreateCell(normalized);
        cell.CellFormula = null;
        cell.CellValue = new CellValue(serial);
        cell.DataType = CellValues.Number;
        ApplyStyle(cell, styleIndex);
        return this;
    }

    /// <summary>
    /// Atomic input validation shared by the row-writing APIs: throws
    /// <see cref="ArgumentNullException"/> when <paramref name="values"/> is null or
    /// contains a null element. It must run BEFORE any worksheet mutation so a bad row
    /// never leaves a partially written row (or stylesheet) behind.
    /// </summary>
    private static void ValidateCellValues(IList<string> values, string paramName)
    {
        ArgumentNullException.ThrowIfNull(values, paramName);

        for (var i = 0; i < values.Count; i++)
        {
            if (values[i] is null)
            {
                throw new ArgumentNullException(
                    $"{paramName}[{i}]",
                    $"Element {i} of '{paramName}' must not be null; row cell values must be non-null strings.");
            }
        }
    }

    // Excel table display names: must start with a letter, underscore or backslash,
    // then only letters, digits, periods and underscores — no spaces or other specials.
    private static readonly Regex TableDisplayNamePattern =
        new(@"^[A-Za-z_\\][A-Za-z0-9._]*$", RegexOptions.Compiled);

    // An absolute R1C1-style cell reference (e.g. 'R1C1', 'R1048576C16384'). The
    // bracket-relative forms ('R[1]C[2]') can never be table names: '[' and ']' are
    // already rejected by TableDisplayNamePattern.
    private static readonly Regex R1C1ReferencePattern =
        new(@"^[Rr][1-9][0-9]*[Cc][1-9][0-9]*$", RegexOptions.Compiled);

    public IWorksheetBuilder AddTable(string startCell, string endCell, string tableName)
    {
        var start = NormalizeCellReference(startCell);
        var end = NormalizeCellReference(endCell);
        ValidateTableDisplayName(tableName);

        var startColumn = GetColumnIndex(start);
        var startRow = GetRowIndex(start);
        var endColumn = GetColumnIndex(end);
        var endRow = GetRowIndex(end);

        if (endColumn < startColumn || endRow < startRow)
        {
            throw new XlsxException(
                $"Table range '{start}:{end}' is reversed: the start cell must be the " +
                "top-left corner of the table and the end cell the bottom-right corner.");
        }

        EnsureNoTableOverlap(startRow, endRow, startColumn, endColumn, tableName);
        _workbookBuilder.RegisterTableName(tableName);

        var columnCount = endColumn - startColumn + 1;
        var columnNames = ResolveTableColumnNames(startRow, startColumn, columnCount);

        var tablePart = _worksheetPart.AddNewPart<TableDefinitionPart>();
        var relationshipId = _worksheetPart.GetIdOfPart(tablePart);
        var table = new Table
        {
            Id = _workbookBuilder.NextTableId(),
            Name = tableName,
            DisplayName = tableName,
            Reference = $"{start}:{end}"
        };

        table.Append(new AutoFilter { Reference = $"{start}:{end}" });
        var tableColumns = new TableColumns { Count = (uint)columnCount };
        for (uint i = 0; i < columnCount; i++)
        {
            tableColumns.Append(new TableColumn { Id = i + 1, Name = columnNames[(int)i] });
        }
        table.Append(tableColumns);
        table.Append(new TableStyleInfo
        {
            Name = "TableStyleMedium2",
            ShowFirstColumn = false,
            ShowLastColumn = false,
            ShowRowStripes = true,
            ShowColumnStripes = false
        });
        tablePart.Table = table;

        var tableParts = _worksheet.Elements<TableParts>().FirstOrDefault();
        if (tableParts == null)
        {
            tableParts = new TableParts();
            _worksheet.Append(tableParts);
        }
        tableParts.Append(new TablePart { Id = relationshipId });
        tableParts.Count = (uint)tableParts.Elements<TablePart>().Count();

        return this;
    }

    private static void ValidateTableDisplayName(string tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName) || !TableDisplayNamePattern.IsMatch(tableName))
        {
            throw new XlsxException(
                $"Invalid table name '{tableName}'. Excel table names must start with a letter, " +
                "underscore or backslash and contain only letters, digits, periods and " +
                "underscores (no spaces or other special characters).");
        }

        if (IsCellReferenceName(tableName))
        {
            throw new XlsxException(
                $"Invalid table name '{tableName}'. Excel rejects table names that look like a " +
                "cell reference (for example A1, BC12, XFD1048576 or R1C1) because it would " +
                "interpret the name as a reference instead of a name.");
        }
    }

    /// <summary>
    /// True when <paramref name="name"/> matches the grammar Excel parses as a cell
    /// reference — A1 style ('A1', 'BC12', 'XFD1048576') or R1C1 style ('R1C1',
    /// 'R2C3') — and the referenced cell actually exists within Excel's real sheet
    /// bounds (columns A-XFD, rows 1-1,048,576). Names that merely resemble an
    /// out-of-bounds reference (e.g. 'XFE1') do not refer to any cell, so they stay
    /// usable as ordinary table names. Table names are case-insensitive in Excel, so
    /// both 'R1C1' and 'r1c1' are caught.
    /// </summary>
    private static bool IsCellReferenceName(string name)
    {
        if (CellReferencePattern.IsMatch(name))
        {
            try
            {
                NormalizeCellReference(name);
                return true;
            }
            catch (XlsxException)
            {
                // Column beyond XFD or row beyond 1,048,576 — not a real cell reference.
                return false;
            }
        }

        if (!R1C1ReferencePattern.IsMatch(name))
        {
            return false;
        }

        var columnLetterIndex = name.IndexOfAny(new[] { 'C', 'c' });
        var rowText = name.AsSpan(1, columnLetterIndex - 1);
        var columnText = name.AsSpan(columnLetterIndex + 1);
        return long.TryParse(rowText, NumberStyles.None, CultureInfo.InvariantCulture, out var row)
            && long.TryParse(columnText, NumberStyles.None, CultureInfo.InvariantCulture, out var column)
            && row >= 1
            && row <= MaxRowNumber
            && column >= 1
            && column <= MaxColumnNumber;
    }

    private void EnsureNoTableOverlap(int startRow, int endRow, int startColumn, int endColumn, string tableName)
    {
        foreach (var existingPart in _worksheetPart.TableDefinitionParts)
        {
            var reference = existingPart.Table?.Reference?.Value;
            if (string.IsNullOrEmpty(reference))
            {
                continue;
            }

            var bounds = reference.Split(':');
            if (bounds.Length != 2)
            {
                continue;
            }

            var existingStartColumn = GetColumnIndex(bounds[0]);
            var existingStartRow = GetRowIndex(bounds[0]);
            var existingEndColumn = GetColumnIndex(bounds[1]);
            var existingEndRow = GetRowIndex(bounds[1]);

            var overlaps = startRow <= existingEndRow && existingStartRow <= endRow
                && startColumn <= existingEndColumn && existingStartColumn <= endColumn;
            if (overlaps)
            {
                throw new XlsxException(
                    $"Table '{tableName}' overlaps existing table " +
                    $"'{existingPart.Table?.Name?.Value}' ({reference}). " +
                    "Excel does not allow overlapping tables on the same worksheet.");
            }
        }
    }

    private List<string> ResolveTableColumnNames(int headerRow, int startColumn, int columnCount)
    {
        // Excel requires table column names to exactly match the header-row cell
        // text, be non-empty, and be unique within the table. Hardcoded "Column{i}"
        // names that disagree with the header cells trigger a repair prompt.
        var names = new List<string>(columnCount);
        for (var offset = 0; offset < columnCount; offset++)
        {
            var headerReference = GetCellReference(startColumn - 1 + offset, headerRow);
            var headerText = GetCellValue(headerReference)?.Trim();
            var baseName = string.IsNullOrEmpty(headerText) ? $"Column{offset + 1}" : headerText;
            names.Add(MakeUniqueColumnName(baseName, names));
        }

        return names;
    }

    private static string MakeUniqueColumnName(string baseName, List<string> existingNames)
    {
        var candidate = baseName;
        var suffix = 2;
        while (existingNames.Contains(candidate, StringComparer.OrdinalIgnoreCase))
        {
            candidate = baseName + suffix;
            suffix++;
        }

        return candidate;
    }

    public IWorksheetBuilder AddChart(ChartType type, string dataRange)
    {
        throw new XlsxException(
            $"Chart creation is not implemented yet. " +
            $"Charts ({type}) are planned for Phase 4 of the XLSX roadmap. " +
            "Use fluent-API or JSON instructions to build worksheet content without charts for now.");
    }

    public IWorksheetBuilder SetColumnWidth(string column, double width)
    {
        var columnNumber = GetColumnNumber(column);
        ValidateColumnWidth(width, column);

        // All validation runs before any mutation, so a rejected width can never leave a
        // partial <col> behind. Idempotency is handled below: repeated calls for the same
        // column update the existing definition instead of appending an overlapping one.
        var cols = GetOrCreateCols();
        var existing = cols.Elements<Column>().FirstOrDefault(c =>
            c.Min?.Value is { } min && c.Max?.Value is { } max
            && min <= columnNumber && columnNumber <= max);

        if (existing != null)
        {
            if (existing.Min?.Value == columnNumber && existing.Max?.Value == columnNumber)
            {
                existing.Width = width;
                existing.CustomWidth = true;
                // bestFit signals an auto-calculated width (auto-width); it must not
                // survive an explicit custom width or Excel may ignore the value we set.
                existing.BestFit = false;
                return this;
            }

            // The target column sits inside a multi-column definition (e.g. min=1 max=5
            // authored by Excel or another tool). Splitting it keeps every other column's
            // width/style intact while inserting a single-column override — an overlapping
            // <col> would corrupt the worksheet's column model.
            SplitColumnRange(existing, columnNumber, width);
            return this;
        }

        InsertColumnAtSortedPosition(cols, CreateWidthCol(columnNumber, width));
        return this;
    }

    public IWorksheetBuilder SetRowHeight(int rowIndex, double height)
    {
        ValidateRowIndex(rowIndex);
        ValidateRowHeight(height, rowIndex);

        // GetOrCreateRow reuses the existing row element (cells and styles untouched) or
        // creates an empty one, so repeated calls update ht/customHeight in place and can
        // never write duplicate <row> definitions.
        var row = GetOrCreateRow(rowIndex);
        row.Height = height;
        row.CustomHeight = true;
        return this;
    }

    public IWorksheetBuilder MergeCells(string range)
    {
        var (start, end) = NormalizeMergeRange(range);

        // All validation runs BEFORE any mutation (no <mergeCell> is appended and the
        // <mergeCells> container is not created) so a rejected range can never leave a
        // partially merged worksheet behind.
        var startRow = GetRowIndex(start);
        var startColumn = GetColumnIndex(start);
        var endRow = GetRowIndex(end);
        var endColumn = GetColumnIndex(end);

        if (startRow == endRow && startColumn == endColumn)
        {
            throw new XlsxException(
                $"Merge range '{range}' is a single cell. " +
                "Excel only merges a range of two or more cells; " +
                "a single cell needs no merge.");
        }

        if (endRow < startRow || endColumn < startColumn)
        {
            throw new XlsxException(
                $"Merge range '{range}' is reversed: the start cell must be the top-left " +
                "corner of the merge and the end cell the bottom-right corner.");
        }

        foreach (var existing in GetMergeCellEntries())
        {
            var (exStartRow, exStartColumn, exEndRow, exEndColumn) = ParseMergeReference(existing, range);

            if (exStartRow == startRow && exStartColumn == startColumn
                && exEndRow == endRow && exEndColumn == endColumn)
            {
                throw new XlsxException(
                    $"Merge range '{range}' is already merged in this worksheet " +
                    $"(existing merge '{existing.Reference?.Value}'). " +
                    "Remove the existing merge first or unmerge the range.");
            }

            var overlaps = startRow <= exEndRow && exStartRow <= endRow
                && startColumn <= exEndColumn && exStartColumn <= endColumn;
            if (overlaps)
            {
                throw new XlsxException(
                    $"Merge range '{range}' overlaps existing merge '{existing.Reference?.Value}'. " +
                    "Excel does not allow overlapping merged ranges on the same worksheet.");
            }
        }

        var mergeCells = GetOrCreateMergeCells();
        mergeCells.Append(new MergeCell { Reference = $"{start}:{end}" });
        mergeCells.Count = (uint)mergeCells.Elements<MergeCell>().Count();
        return this;
    }

    public IWorksheetBuilder UnmergeCells(string range)
    {
        var (start, end) = NormalizeMergeRange(range);
        var startRow = GetRowIndex(start);
        var startColumn = GetColumnIndex(start);
        var endRow = GetRowIndex(end);
        var endColumn = GetColumnIndex(end);

        if (startRow == endRow && startColumn == endColumn)
        {
            throw new XlsxException(
                $"Merge range '{range}' is a single cell. " +
                "A merge spans two or more cells, so there is nothing to unmerge.");
        }

        if (endRow < startRow || endColumn < startColumn)
        {
            throw new XlsxException(
                $"Merge range '{range}' is reversed: the start cell must be the top-left " +
                "corner of the merge and the end cell the bottom-right corner.");
        }

        // Unmerging a range that is not currently merged is a no-op — matching Excel's
        // behaviour and keeping the operation safe to call unconditionally.
        var mergeCells = _worksheet.GetFirstChild<MergeCells>();
        if (mergeCells == null)
        {
            return this;
        }

        foreach (var existing in mergeCells.Elements<MergeCell>().ToList())
        {
            var (exStartRow, exStartColumn, exEndRow, exEndColumn) = ParseMergeReference(existing, range);
            if (exStartRow == startRow && exStartColumn == startColumn
                && exEndRow == endRow && exEndColumn == endColumn)
            {
                existing.Remove();
                mergeCells.Count = (uint)mergeCells.Elements<MergeCell>().Count();
                if (mergeCells.Count?.Value == 0)
                {
                    // An empty <mergeCells> container is unnecessary; drop it so
                    // GetOrCreateMergeCells can rebuild it cleanly on the next merge.
                    mergeCells.Remove();
                }
                return this;
            }
        }

        return this;
    }

    // ─── Freeze panes ──────────────────────────────────────────────

    public IWorksheetBuilder FreezePanes(int frozenRows, int frozenColumns)
    {
        // Validation before any mutation: a frozen pane needs at least one frozen row or
        // column, and the pane's top-left cell must exist within Excel's real grid.
        if (frozenRows < 0 || frozenColumns < 0)
        {
            throw new XlsxException(
                $"Freeze pane dimensions must be non-negative; got rows={frozenRows}, " +
                $"columns={frozenColumns}.");
        }

        if (frozenRows == 0 && frozenColumns == 0)
        {
            throw new XlsxException(
                "Freeze pane dimensions cannot both be zero: a frozen pane must freeze at " +
                "least one row or one column.");
        }

        if (frozenRows > MaxRowNumber - 1 || frozenColumns > MaxColumnNumber - 1)
        {
            throw new XlsxException(
                $"Freeze pane dimensions are out of range: at most {MaxRowNumber - 1:N0} " +
                $"rows and {MaxColumnNumber - 1} columns can be frozen. Got rows={frozenRows}, " +
                $"columns={frozenColumns}.");
        }

        var sheetViews = GetOrCreateSheetViews();
        var sheetView = sheetViews.Elements<SheetView>().FirstOrDefault();
        if (sheetView is null)
        {
            sheetView = new SheetView();
            sheetViews.Append(sheetView);
        }

        if (sheetView.WorkbookViewId is null)
        {
            sheetView.WorkbookViewId = 0;
        }

        var pane = sheetView.Pane ?? new Pane();
        pane.HorizontalSplit = (double)frozenColumns;
        pane.VerticalSplit = (double)frozenRows;
        pane.State = PaneStateValues.Frozen;
        pane.TopLeftCell = GetCellReference(frozenColumns, frozenRows + 1);
        pane.ActivePane = frozenRows > 0 && frozenColumns > 0
            ? PaneValues.BottomRight
            : frozenRows > 0 ? PaneValues.BottomLeft : PaneValues.TopRight;

        if (sheetView.Pane is null)
        {
            sheetView.Append(pane);
        }

        // Excel pairs a frozen pane with a selection naming the active pane, so the user
        // never lands in a frozen (non-scrollable) region.
        if (!sheetView.Elements<Selection>().Any())
        {
            sheetView.Append(new Selection
            {
                Pane = pane.ActivePane,
                ActiveCell = pane.TopLeftCell
            });
        }

        return this;
    }

    public (int FrozenRows, int FrozenColumns)? GetFreezePanes()
    {
        var pane = _worksheet.GetFirstChild<SheetViews>()
            ?.Elements<SheetView>().FirstOrDefault()?.Pane;
        if (pane?.State?.Value != PaneStateValues.Frozen)
        {
            return null;
        }

        return ((int)(pane.VerticalSplit?.Value ?? 0), (int)(pane.HorizontalSplit?.Value ?? 0));
    }

    /// <summary>
    /// Returns the worksheet's &lt;sheetViews&gt; container, creating and inserting it in
    /// schema position (immediately after sheetPr/dimension, before sheetFormatPr) on
    /// first use. The element is reused afterwards, so every sheet view lives in one list
    /// with a Count that stays in sync with its &lt;sheetView&gt; children.
    /// </summary>
    private SheetViews GetOrCreateSheetViews()
    {
        var existing = _worksheet.GetFirstChild<SheetViews>();
        if (existing is not null)
        {
            return existing;
        }

        var sheetViews = new SheetViews();
        InsertWorksheetElementAtSchemaPosition(sheetViews);
        return sheetViews;
    }

    // ─── Standalone autofilter ─────────────────────────────────────

    public IWorksheetBuilder SetAutoFilter(string range)
    {
        var (start, end) = NormalizeMergeRange(range);
        var startRow = GetRowIndex(start);
        var startColumn = GetColumnIndex(start);
        var endRow = GetRowIndex(end);
        var endColumn = GetColumnIndex(end);

        if (endRow < startRow || endColumn < startColumn)
        {
            throw new XlsxException(
                $"AutoFilter range '{range}' is reversed: the start cell must be the " +
                "top-left corner of the range and the end cell the bottom-right corner.");
        }

        var autoFilter = _worksheet.GetFirstChild<AutoFilter>();
        if (autoFilter is null)
        {
            autoFilter = new AutoFilter();
            InsertWorksheetElementAtSchemaPosition(autoFilter);
        }

        autoFilter.Reference = $"{start}:{end}";
        return this;
    }

    public string? GetAutoFilterRange()
    {
        return _worksheet.GetFirstChild<AutoFilter>()?.Reference?.Value;
    }

    public IWorksheetBuilder RemoveAutoFilter()
    {
        _worksheet.GetFirstChild<AutoFilter>()?.Remove();
        return this;
    }

    /// <summary>
    /// Inserts a new worksheet child at its correct position in the CT_Worksheet child
    /// sequence (see <see cref="WorksheetChildOrder"/>). Existing children of a reopened
    /// worksheet are already in schema order, so inserting before the first child that
    /// must follow is correct whether the worksheet was just created or carries elements
    /// such as sheetFormatPr, autoFilter, mergeCells or tableParts.
    /// </summary>
    private void InsertWorksheetElementAtSchemaPosition(OpenXmlElement element)
    {
        var order = GetWorksheetChildOrder(element.LocalName);
        foreach (var child in _worksheet.ChildElements)
        {
            if (GetWorksheetChildOrder(child.LocalName) > order)
            {
                _worksheet.InsertBefore(element, child);
                return;
            }
        }

        _worksheet.Append(element);
    }

    public List<string> GetMergeRanges()
    {
        var mergeCells = _worksheet.GetFirstChild<MergeCells>();
        if (mergeCells == null)
        {
            return new List<string>();
        }

        return mergeCells.Elements<MergeCell>()
            .Select(c => c.Reference?.Value)
            .Where(r => r != null)
            .Cast<string>()
            .ToList();
    }

    public string? GetCellValue(string cellReference)
    {
        var cell = FindCell(cellReference);
        if (cell == null) return null;

        if (cell.CellFormula != null && cell.CellValue != null)
        {
            return cell.CellValue.Text;
        }

        if (cell.DataType?.Value == CellValues.SharedString)
        {
            if (int.TryParse(cell.CellValue?.Text, out var idx))
            {
                return _workbookBuilder.GetSharedStringByIndex(idx);
            }
            return null;
        }

        if (cell.DataType?.Value == CellValues.InlineString || cell.DataType?.Value == CellValues.String)
        {
            return cell.InnerText;
        }

        if (cell.CellValue != null)
        {
            return cell.CellValue.Text;
        }

        return null;
    }

    /// <summary>
    /// Returns the cell's formula in Excel display syntax (with a leading '=').
    /// The stored &lt;f&gt; text never contains '=' (SpreadsheetML requirement);
    /// it is re-prepended here so callers always see the familiar form.
    /// </summary>
    public string? GetCellFormula(string cellReference)
    {
        var cell = FindCell(cellReference);
        return ToDisplayFormula(cell?.CellFormula);
    }

    public bool CellExists(string cellReference)
    {
        return FindCell(cellReference) != null;
    }

    public CellInfo? GetCellInfo(string cellReference)
    {
        var cell = FindCell(cellReference);
        if (cell == null) return null;

        return new CellInfo
        {
            Reference = cell.CellReference?.Value ?? cellReference,
            Value = GetCellValue(cellReference),
            Formula = ToDisplayFormula(cell.CellFormula),
            DataType = cell.DataType?.Value
        };
    }

    public List<CellInfo> GetRange(string start, string end)
    {
        var startCol = GetColumnIndex(start);
        var startRow = GetRowIndex(start);
        var endCol = GetColumnIndex(end);
        var endRow = GetRowIndex(end);

        if (endRow < startRow || endCol < startCol)
        {
            throw new XlsxException(
                $"Range '{start}:{end}' is reversed: the start cell must be the top-left " +
                "corner and the end cell the bottom-right corner of the range.");
        }

        var results = new List<CellInfo>();
        for (int r = startRow; r <= endRow; r++)
        {
            for (int c = startCol; c <= endCol; c++)
            {
                var ref_ = GetCellReference(c - 1, r);
                var info = GetCellInfo(ref_);
                results.Add(info ?? new CellInfo { Reference = ref_ });
            }
        }
        return results;
    }

    public List<RowInfo> GetRows()
    {
        return _sheetData.Elements<Row>()
            .OrderBy(r => r.RowIndex?.Value)
            .Select(row => new RowInfo
            {
                RowIndex = (int?)row.RowIndex?.Value ?? 0,
                Cells = row.Elements<Cell>()
                    .Select(c => new CellInfo
                    {
                        Reference = c.CellReference?.Value ?? string.Empty,
                        Value = ResolveCellValue(c),
                        Formula = ToDisplayFormula(c.CellFormula),
                        DataType = c.DataType?.Value
                    })
                    .ToList()
            })
            .ToList();
    }

    public RowInfo? GetRow(int rowIndex)
    {
        ValidateRowIndex(rowIndex);

        var row = _sheetData.Elements<Row>()
            .FirstOrDefault(r => r.RowIndex?.Value == (uint)rowIndex);
        if (row == null) return null;

        return new RowInfo
        {
            RowIndex = rowIndex,
            Cells = row.Elements<Cell>()
                .Select(c => new CellInfo
                {
                    Reference = c.CellReference?.Value ?? string.Empty,
                    Value = ResolveCellValue(c),
                    Formula = ToDisplayFormula(c.CellFormula),
                    DataType = c.DataType?.Value
                })
                .ToList()
        };
    }

    public (int firstRow, int lastRow, int firstCol, int lastCol) GetDimensions()
    {
        var rows = _sheetData.Elements<Row>().ToList();
        if (rows.Count == 0)
        {
            return (0, 0, 0, 0);
        }

        var rowIndices = rows.Select(r => (int)r.RowIndex!.Value).OrderBy(i => i).ToList();
        var firstRow = rowIndices.First();
        var lastRow = rowIndices.Last();

        int firstCol = int.MaxValue;
        int lastCol = 0;
        foreach (var row_ in rows)
        {
            foreach (var c in row_.Elements<Cell>())
            {
                if (c.CellReference?.Value is not string ref_) continue;
                var col = GetColumnIndex(ref_);
                if (col < firstCol) firstCol = col;
                if (col > lastCol) lastCol = col;
            }
        }

        if (firstCol == int.MaxValue) firstCol = 0;

        return (firstRow, lastRow, firstCol, lastCol);
    }

    public double? GetColumnWidth(string column)
    {
        var columnNumber = GetColumnNumber(column);
        var cols = _worksheet.GetFirstChild<Columns>();
        var col = cols?.Elements<Column>().FirstOrDefault(c =>
            c.Min?.Value is { } min && c.Max?.Value is { } max
            && min <= columnNumber && columnNumber <= max);
        return col?.Width?.Value;
    }

    public double? GetRowHeight(int rowIndex)
    {
        ValidateRowIndex(rowIndex);
        var row = _sheetData.Elements<Row>()
            .FirstOrDefault(r => r.RowIndex?.Value == (uint)rowIndex);
        return row?.Height?.Value;
    }

    public IWorksheetBuilder DeleteCell(string cellReference)
    {
        var cell = FindCell(cellReference);
        cell?.Remove();
        return this;
    }

    public IWorksheetBuilder DeleteRow(int rowIndex)
    {
        ValidateRowIndex(rowIndex);

        var row = _sheetData.Elements<Row>()
            .FirstOrDefault(r => r.RowIndex?.Value == (uint)rowIndex);
        row?.Remove();
        return this;
    }

    public IWorksheetBuilder ClearRange(string start, string end)
    {
        var startCol = GetColumnIndex(start);
        var startRow = GetRowIndex(start);
        var endCol = GetColumnIndex(end);
        var endRow = GetRowIndex(end);

        if (endRow < startRow || endCol < startCol)
        {
            throw new XlsxException(
                $"Range '{start}:{end}' is reversed: the start cell must be the top-left " +
                "corner and the end cell the bottom-right corner of the range.");
        }

        for (int r = startRow; r <= endRow; r++)
        {
            for (int c = startCol; c <= endCol; c++)
            {
                var ref_ = GetCellReference(c - 1, r);
                DeleteCell(ref_);
            }
        }
        return this;
    }

    /// <summary>
    /// Converts stored formula text (no '=') to Excel display syntax (leading '=').
    /// </summary>
    private static string? ToDisplayFormula(CellFormula? formula)
    {
        return formula?.Text is { Length: > 0 } text ? "=" + text : null;
    }
    private Cell? FindCell(string cellReference)
    {
        cellReference = NormalizeCellReference(cellReference);
        var rowIndex = GetRowIndex(cellReference);
        var row = _sheetData.Elements<Row>()
            .FirstOrDefault(r => r.RowIndex?.Value == (uint)rowIndex);
        if (row == null) return null;

        return row.Elements<Cell>()
            .FirstOrDefault(c => c.CellReference?.Value == cellReference);
    }

    private string? ResolveCellValue(Cell cell)
    {
        if (cell.CellFormula != null && cell.CellValue != null)
        {
            return cell.CellValue.Text;
        }

        if (cell.DataType?.Value == CellValues.SharedString)
        {
            if (int.TryParse(cell.CellValue?.Text, out var idx))
            {
                return _workbookBuilder.GetSharedStringByIndex(idx);
            }
            return null;
        }

        if (cell.DataType?.Value == CellValues.InlineString || cell.DataType?.Value == CellValues.String)
        {
            return cell.InnerText;
        }

        return cell.CellValue?.Text;
    }

    private Cell GetOrCreateCell(string cellReference)
    {
        // Normalize at the entry point: Excel references are case-insensitive, so
        // 'a1' and 'A1' must resolve to the same cell rather than duplicate cells.
        cellReference = NormalizeCellReference(cellReference);
        var rowIndex = GetRowIndex(cellReference);
        var row = GetOrCreateRow(rowIndex);

        var cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == cellReference);
        if (cell == null)
        {
            cell = new Cell { CellReference = cellReference };
            InsertCellAtSortedPosition(row, cell);
        }

        return cell;
    }

    private static void InsertCellAtSortedPosition(Row row, Cell newCell)
    {
        var newRef = newCell.CellReference?.Value;
        if (string.IsNullOrEmpty(newRef))
        {
            row.Append(newCell);
            return;
        }

        foreach (var existing in row.Elements<Cell>())
        {
            var existingRef = existing.CellReference?.Value;
            if (string.IsNullOrEmpty(existingRef))
            {
                continue;
            }

            if (CompareCellReferences(newRef, existingRef) < 0)
            {
                row.InsertBefore(newCell, existing);
                return;
            }
        }

        row.Append(newCell);
    }

    private static int CompareCellReferences(string a, string b)
    {
        var colA = GetColumnIndex(a);
        var colB = GetColumnIndex(b);
        if (colA != colB)
        {
            return colA.CompareTo(colB);
        }

        var rowA = GetRowIndex(a);
        var rowB = GetRowIndex(b);
        return rowA.CompareTo(rowB);
    }

    private Row GetOrCreateRow(int rowIndex)
    {
        ValidateRowIndex(rowIndex);

        var targetIndex = (uint)rowIndex;
        var row = _sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == targetIndex);
        if (row == null)
        {
            row = new Row { RowIndex = targetIndex };
            InsertRowAtSortedPosition(row);
        }

        return row;
    }

    private void InsertRowAtSortedPosition(Row newRow)
    {
        var targetIndex = newRow.RowIndex?.Value;
        if (targetIndex == null)
        {
            _sheetData.Append(newRow);
            return;
        }

        foreach (var existing in _sheetData.Elements<Row>())
        {
            var existingIndex = existing.RowIndex?.Value;
            if (existingIndex == null)
            {
                continue;
            }

            if (targetIndex < existingIndex)
            {
                _sheetData.InsertBefore(newRow, existing);
                return;
            }
        }

        _sheetData.Append(newRow);
    }

    /// <summary>
    /// Row indexes are 1-based and bounded by Excel's real sheet limit (1-1,048,576).
    /// Shared by every row-facing entry point so GetRow, DeleteRow, GetOrCreateRow and
    /// GetCellReference can never disagree about what a legal row index is.
    /// </summary>
    private static void ValidateRowIndex(int rowIndex)
    {
        if (rowIndex < 1)
        {
            throw new XlsxException(
                $"Row index must be >= 1; Excel rows are 1-based. Got {rowIndex}.");
        }

        if (rowIndex > MaxRowNumber)
        {
            throw new XlsxException(
                $"Row index must be between 1 and {MaxRowNumber:N0}; Excel rows are 1-based. " +
                $"Got {rowIndex:N0}.");
        }
    }

    internal static string GetCellReference(int columnIndex, int rowIndex)
    {
        if (columnIndex < 0 || columnIndex >= MaxColumnNumber)
        {
            throw new XlsxException(
                $"Column index must be between 0 and {MaxColumnNumber - 1} " +
                $"(0 = column A, {MaxColumnNumber - 1} = column XFD). Got {columnIndex}.");
        }

        ValidateRowIndex(rowIndex);

        var columnName = GetColumnName(columnIndex);
        return $"{columnName}{rowIndex}";
    }

    internal static string GetColumnName(int index)
    {
        var name = string.Empty;
        var n = index + 1;

        while (n > 0)
        {
            n--;
            name = Convert.ToChar('A' + (n % 26)) + name;
            n /= 26;
        }

        return name;
    }

    /// <summary>
    /// Validates an A1-style cell reference and returns its canonical upper-case form.
    /// Excel cell references are case-insensitive; normalizing prevents 'a1' and 'A1'
    /// from becoming two distinct cells in the same row. Malformed references throw
    /// an <see cref="XlsxException"/> naming the offending reference (instead of a
    /// raw <see cref="FormatException"/> from deep in the parser). References are also
    /// bounded to Excel's real limits (columns A-XFD, rows 1-1,048,576); the row is
    /// parsed as <see cref="long"/> so absurdly large row strings fail with a clear
    /// <see cref="XlsxException"/> instead of an integer <see cref="OverflowException"/>.
    /// </summary>
    internal static string NormalizeCellReference(string cellReference)
    {
        if (string.IsNullOrWhiteSpace(cellReference) || !CellReferencePattern.IsMatch(cellReference))
        {
            throw new XlsxException(
                $"Invalid cell reference '{cellReference}'. " +
                "Expected an Excel A1-style reference such as 'A1' or 'BC12' " +
                "(1-3 letters, then a row number >= 1).");
        }

        var normalized = cellReference.ToUpperInvariant();
        var letterCount = CountLeadingLetters(normalized);

        var columnIndex = ColumnIndexFromLetters(normalized.AsSpan(0, letterCount));
        if (columnIndex > MaxColumnNumber)
        {
            throw new XlsxException(
                $"Cell reference '{cellReference}' uses column '{normalized[..letterCount]}', " +
                "which is beyond Excel's maximum column 'XFD'.");
        }

        var rowText = normalized.AsSpan(letterCount);
        if (!long.TryParse(rowText, NumberStyles.None, CultureInfo.InvariantCulture, out var rowIndex)
            || rowIndex < 1
            || rowIndex > MaxRowNumber)
        {
            throw new XlsxException(
                $"Cell reference '{cellReference}' uses row '{rowText}', which is outside " +
                $"Excel's valid row range 1-{MaxRowNumber:N0}.");
        }

        return normalized;
    }

    internal static int GetRowIndex(string cellReference)
    {
        var normalized = NormalizeCellReference(cellReference);
        var letterCount = CountLeadingLetters(normalized);
        return int.Parse(normalized.AsSpan(letterCount), NumberStyles.None, CultureInfo.InvariantCulture);
    }

    internal static int GetColumnIndex(string cellReference)
    {
        var normalized = NormalizeCellReference(cellReference);
        return ColumnIndexFromLetters(normalized.AsSpan(0, CountLeadingLetters(normalized)));
    }

    private static int ColumnIndexFromLetters(ReadOnlySpan<char> letters)
    {
        var result = 0;
        foreach (var c in letters)
        {
            result = result * 26 + c - 'A' + 1;
        }

        return result;
    }

    private static int CountLeadingLetters(string reference)
    {
        var count = 0;
        while (count < reference.Length && char.IsLetter(reference[count]))
        {
            count++;
        }

        return count;
    }

    // ─── Merged cells helpers ─────────────────────────────────────

    /// <summary>
    /// Splits an A1-style range ("A1:C3") into its two canonical upper-case endpoint cell
    /// references. The endpoints are normalized and bounded to Excel's real sheet limits
    /// (columns A-XFD, rows 1-1,048,576) by <see cref="NormalizeCellReference"/>; a range
    /// without a colon, or with a malformed/out-of-bounds endpoint, throws an
    /// <see cref="XlsxException"/> before any mutation.
    /// </summary>
    private static (string start, string end) NormalizeMergeRange(string range)
    {
        ArgumentNullException.ThrowIfNull(range);

        var bounds = range.Split(':');
        if (bounds.Length != 2)
        {
            throw new XlsxException(
                $"Invalid merge range '{range}'. Expected an A1-style range such as 'A1:C3' " +
                "(two cell references separated by a colon).");
        }

        return (NormalizeCellReference(bounds[0]), NormalizeCellReference(bounds[1]));
    }

    /// <summary>
    /// Parses an existing &lt;mergeCell&gt; reference into its bounding rows/columns. A
    /// merge entry that is missing, malformed or reversed means the workbook's merge data is
    /// already corrupt; that is surfaced as an <see cref="XlsxException"/> naming both the
    /// offending entry and the incoming range it was being checked against, instead of being
    /// silently skipped.
    /// </summary>
    private static (int startRow, int startColumn, int endRow, int endColumn) ParseMergeReference(
        MergeCell mergeCell, string incomingRange)
    {
        var reference = mergeCell.Reference?.Value;
        if (string.IsNullOrEmpty(reference))
        {
            throw new XlsxException(
                $"This worksheet contains a merge entry with no reference, so it cannot be " +
                $"checked against '{incomingRange}'. The workbook's merge data is corrupt; " +
                "repair or remove the invalid entry.");
        }

        var bounds = reference.Split(':');
        if (bounds.Length != 2)
        {
            throw new XlsxException(
                $"This worksheet contains an invalid merge reference '{reference}', so it " +
                $"cannot be checked against '{incomingRange}'. Expected an A1-style range " +
                $"such as 'A1:C3'. The workbook's merge data is corrupt; repair or remove " +
                "the invalid entry.");
        }

        var startRow = GetRowIndex(bounds[0]);
        var startColumn = GetColumnIndex(bounds[0]);
        var endRow = GetRowIndex(bounds[1]);
        var endColumn = GetColumnIndex(bounds[1]);

        if (endRow < startRow || endColumn < startColumn)
        {
            throw new XlsxException(
                $"This worksheet contains a reversed merge reference '{reference}', so it " +
                $"cannot be checked against '{incomingRange}'. Expected a range whose start " +
                "is the top-left corner and whose end is the bottom-right corner. The " +
                "workbook's merge data is corrupt; repair or remove the invalid entry.");
        }

        return (startRow, startColumn, endRow, endColumn);
    }

    private List<MergeCell> GetMergeCellEntries()
    {
        var mergeCells = _worksheet.GetFirstChild<MergeCells>();
        if (mergeCells == null)
        {
            return new List<MergeCell>();
        }

        return mergeCells.Elements<MergeCell>().ToList();
    }

    /// <summary>
    /// Returns the worksheet's &lt;mergeCells&gt; container, creating and inserting it in
    /// schema position on first use. The element is reused afterwards, so every merge lives
    /// in one list with a Count that stays in sync with its &lt;mergeCell&gt; children.
    /// </summary>
    private MergeCells GetOrCreateMergeCells()
    {
        var existing = _worksheet.GetFirstChild<MergeCells>();
        if (existing != null)
        {
            return existing;
        }

        var mergeCells = new MergeCells();
        InsertMergeCellsAtSchemaPosition(mergeCells);
        return mergeCells;
    }

    /// <summary>
    /// Inserts a new &lt;mergeCells&gt; container at its correct position in the
    /// CT_Worksheet child sequence (immediately after customSheetViews, before phoneticPr).
    /// Existing children of a reopened worksheet are already in schema order, so inserting
    /// before the first child that follows mergeCells is correct whether the worksheet was
    /// just created or carries elements such as autoFilter, hyperlinks or tableParts.
    /// </summary>
    private void InsertMergeCellsAtSchemaPosition(MergeCells mergeCells)
    {
        foreach (var child in _worksheet.ChildElements)
        {
            if (GetWorksheetChildOrder(child.LocalName) > MergeCellsOrder)
            {
                _worksheet.InsertBefore(mergeCells, child);
                return;
            }
        }

        _worksheet.Append(mergeCells);
    }

    private const int MergeCellsOrder = 15;

    /// <summary>
    /// Ordinal positions of the worksheet child elements in the CT_Worksheet sequence
    /// (ECMA-376 §18.3.1.99). mergeCells sits at position 15, between customSheetViews and
    /// phoneticPr. Elements not listed (non-standard children) are treated as following
    /// mergeCells — the safe side for an element that sits two-thirds into the sequence.
    /// </summary>
    private static readonly Dictionary<string, int> WorksheetChildOrder = new()
    {
        ["sheetPr"] = 1,
        ["dimension"] = 2,
        ["sheetViews"] = 3,
        ["sheetFormatPr"] = 4,
        ["cols"] = 5,
        ["sheetData"] = 6,
        ["sheetCalcPr"] = 7,
        ["sheetProtection"] = 8,
        ["protectedRanges"] = 9,
        ["scenarios"] = 10,
        ["autoFilter"] = 11,
        ["sortState"] = 12,
        ["dataConsolidate"] = 13,
        ["customSheetViews"] = 14,
        ["mergeCells"] = 15,
        ["phoneticPr"] = 16,
        ["conditionalFormatting"] = 17,
        ["dataValidations"] = 18,
        ["hyperlinks"] = 19,
        ["printOptions"] = 20,
        ["pageMargins"] = 21,
        ["pageSetup"] = 22,
        ["headerFooter"] = 23,
        ["rowBreaks"] = 24,
        ["colBreaks"] = 25,
        ["customProperties"] = 26,
        ["cellWatches"] = 27,
        ["ignoredErrors"] = 28,
        ["smartTags"] = 29,
        ["drawing"] = 30,
        ["legacyDrawing"] = 31,
        ["legacyDrawingHF"] = 32,
        ["picture"] = 33,
        ["oleObjects"] = 34,
        ["controls"] = 35,
        ["webPublishItems"] = 36,
        ["tableParts"] = 37,
        ["extLst"] = 38
    };

    private static int GetWorksheetChildOrder(string localName)
    {
        return WorksheetChildOrder.TryGetValue(localName, out var order) ? order : int.MaxValue;
    }

    // ─── Column width / row height helpers ────────────────────────

    /// <summary>
    /// Validates an A1-style column letter ('A'..'XFD', case-insensitive) and returns its
    /// 1-based column number (1 = A, 16384 = XFD). A row-less column letter — not a cell
    /// reference — is the natural Excel vocabulary for addressing a column; passing a cell
    /// reference or a row number is a misuse and is rejected with a clear message.
    /// </summary>
    private static int GetColumnNumber(string column)
    {
        if (string.IsNullOrWhiteSpace(column))
        {
            throw new XlsxException(
                "Column must not be empty or whitespace. " +
                "Use an A1-style column letter such as 'A' or 'AB'.");
        }

        if (column.Length > 3 || !column.All(char.IsAsciiLetter))
        {
            throw new XlsxException(
                $"Invalid column '{column}'. Expected an A1-style column letter such as " +
                "'A' or 'AB' (1-3 letters, no row number).");
        }

        var columnNumber = ColumnIndexFromLetters(column.ToUpperInvariant());
        if (columnNumber > MaxColumnNumber)
        {
            throw new XlsxException(
                $"Column '{column}' is beyond Excel's maximum column 'XFD'.");
        }

        return columnNumber;
    }

    private static void ValidateColumnWidth(double width, string column)
    {
        if (!double.IsFinite(width))
        {
            throw new XlsxException(
                $"Column width for '{column}' must be a finite number; NaN and " +
                "infinity are not valid Excel column widths.");
        }

        if (width <= 0)
        {
            throw new XlsxException(
                $"Column width for '{column}' must be greater than zero " +
                $"(Excel column-width units). Got {width}.");
        }

        if (width > MaxColumnWidth)
        {
            throw new XlsxException(
                $"Column width for '{column}' must not exceed {MaxColumnWidth} " +
                $"(Excel's maximum column width). Got {width}.");
        }
    }

    private static void ValidateRowHeight(double height, int rowIndex)
    {
        if (!double.IsFinite(height))
        {
            throw new XlsxException(
                $"Row height for row {rowIndex} must be a finite number; NaN and " +
                "infinity are not valid Excel row heights.");
        }

        if (height <= 0)
        {
            throw new XlsxException(
                $"Row height for row {rowIndex} must be greater than zero (points). " +
                $"Got {height}.");
        }

        if (height > MaxRowHeight)
        {
            throw new XlsxException(
                $"Row height for row {rowIndex} must not exceed {MaxRowHeight} points " +
                $"(Excel's maximum row height). Got {height}.");
        }
    }

    /// <summary>
    /// Returns the worksheet's &lt;cols&gt; container, creating and inserting it in schema
    /// position (immediately before &lt;sheetData&gt;) on first use. The element is reused
    /// afterwards, so every column definition lives in one sorted, non-overlapping list.
    /// </summary>
    private Columns GetOrCreateCols()
    {
        var cols = _worksheet.GetFirstChild<Columns>();
        if (cols != null)
        {
            return cols;
        }

        cols = new Columns();
        _worksheet.InsertBefore(cols, _sheetData);
        return cols;
    }

    private static void InsertColumnAtSortedPosition(Columns cols, Column newCol)
    {
        // Column ranges are non-overlapping, so ordering by Min alone is sufficient: a new
        // single-column definition belongs just before the first range that starts past it.
        var min = newCol.Min?.Value ?? 0;
        foreach (var existing in cols.Elements<Column>())
        {
            if ((existing.Min?.Value ?? 0) > min)
            {
                cols.InsertBefore(newCol, existing);
                return;
            }
        }

        cols.Append(newCol);
    }

    private static Column CreateWidthCol(int columnNumber, double width)
    {
        return new Column
        {
            Min = (uint)columnNumber,
            Max = (uint)columnNumber,
            Width = width,
            CustomWidth = true
        };
    }

    /// <summary>
    /// Splits a multi-column definition that covers <paramref name="columnNumber"/> into up
    /// to three non-overlapping parts (left, single-column override, right), preserving the
    /// original definition's width and formatting attributes on the untouched parts.
    /// </summary>
    private static void SplitColumnRange(Column existing, int columnNumber, double width)
    {
        var container = (Columns)existing.Parent!;
        var min = existing.Min!.Value;
        var max = existing.Max!.Value;

        var parts = new List<Column>(3);
        if (min < columnNumber)
        {
            parts.Add(CopyColumnRange(existing, min, (uint)(columnNumber - 1)));
        }

        // The override is a copy of the original range too, so it keeps every unrelated
        // attribute (style, hidden, outline level, collapsed, phonetic…) of the definition
        // being split — only the width aspects are replaced and bestFit is cleared, since an
        // explicit width can never be auto-fitted.
        var target = CopyColumnRange(existing, (uint)columnNumber, (uint)columnNumber);
        target.Width = width;
        target.CustomWidth = true;
        target.BestFit = false;
        parts.Add(target);

        if (columnNumber < max)
        {
            parts.Add(CopyColumnRange(existing, (uint)(columnNumber + 1), max));
        }

        container.InsertBefore(parts[0], existing);
        for (var i = 1; i < parts.Count; i++)
        {
            container.InsertAfter(parts[i], parts[i - 1]);
        }

        container.RemoveChild(existing);
    }

    private static Column CopyColumnRange(Column source, uint min, uint max)
    {
        return new Column
        {
            Min = min,
            Max = max,
            Width = source.Width,
            CustomWidth = source.CustomWidth,
            BestFit = source.BestFit,
            Hidden = source.Hidden,
            Style = source.Style,
            Phonetic = source.Phonetic,
            OutlineLevel = source.OutlineLevel,
            Collapsed = source.Collapsed
        };
    }
}
