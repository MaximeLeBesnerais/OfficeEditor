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

    // Excel table display names: must start with a letter, underscore or backslash,
    // then only letters, digits, periods and underscores — no spaces or other specials.
    private static readonly Regex TableDisplayNamePattern =
        new(@"^[A-Za-z_\\][A-Za-z0-9._]*$", RegexOptions.Compiled);

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
}
