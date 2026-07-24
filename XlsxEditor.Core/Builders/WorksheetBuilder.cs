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

        if (double.TryParse(value, out var numericValue))
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
        }
        else
        {
            cell.CellValue = new CellValue(formula);
            cell.DataType = CellValues.String;
        }

        return this;
    }

    public IWorksheetBuilder AddCell(string cellReference, string value, string styleId)
    {
        if (!uint.TryParse(styleId, out var parsedStyleId))
        {
            throw new XlsxException(
                $"Invalid styleId '{styleId}' for cell {cellReference}. " +
                "Style identifiers must be unsigned integers. " +
                "Use the Phase 3 style builder for named styles.");
        }

        AddCell(cellReference, value);
        var cell = GetOrCreateCell(cellReference);
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

    public IWorksheetBuilder AddTable(string startCell, string endCell, string tableName)
    {
        var tablePart = _worksheetPart.AddNewPart<TableDefinitionPart>();
        var relationshipId = _worksheetPart.GetIdOfPart(tablePart);
        var columnCount = GetColumnIndex(endCell) - GetColumnIndex(startCell) + 1;
        var table = new Table
        {
            Id = _workbookBuilder.NextTableId(),
            Name = tableName,
            DisplayName = tableName,
            Reference = $"{startCell}:{endCell}"
        };

        table.Append(new AutoFilter { Reference = $"{startCell}:{endCell}" });
        var tableColumns = new TableColumns { Count = (uint)columnCount };
        for (uint i = 1; i <= columnCount; i++)
        {
            tableColumns.Append(new TableColumn { Id = i, Name = $"Column{i}" });
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
    {        var rowIndex = GetRowIndex(cellReference);
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
        var colA = GetColumnIndexStatic(a);
        var colB = GetColumnIndexStatic(b);
        if (colA != colB)
        {
            return colA.CompareTo(colB);
        }

        var rowA = GetRowIndexStatic(a);
        var rowB = GetRowIndexStatic(b);
        return rowA.CompareTo(rowB);
    }

    private Row GetOrCreateRow(int rowIndex)
    {
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

    internal static string GetCellReference(int columnIndex, int rowIndex)
    {
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

    internal static int GetRowIndex(string cellReference)
    {
        var rowPart = string.Empty;
        foreach (var c in cellReference)
        {
            if (char.IsDigit(c))
            {
                rowPart += c;
            }
        }

        return int.Parse(rowPart);
    }

    internal static int GetColumnIndex(string cellReference)
    {
        var result = 0;
        foreach (var c in cellReference)
        {
            if (!char.IsLetter(c))
            {
                break;
            }

            result = result * 26 + char.ToUpperInvariant(c) - 'A' + 1;
        }

        return result;
    }

    private static int GetColumnIndexStatic(string cellReference)
    {
        var result = 0;
        foreach (var c in cellReference)
        {
            if (!char.IsLetter(c))
            {
                break;
            }

            result = result * 26 + char.ToUpperInvariant(c) - 'A' + 1;
        }

        return result;
    }

    private static int GetRowIndexStatic(string cellReference)
    {
        var rowPart = string.Empty;
        foreach (var c in cellReference)
        {
            if (char.IsDigit(c))
            {
                rowPart += c;
            }
        }

        return int.Parse(rowPart);
    }
}
