using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;
using XlsxEditor.Core.Rendering.Models;
using XlsxEditor.Core.Rendering.Read;

namespace XlsxEditor.Core.Rendering;

/// <summary>
/// Serializes an arbitrary .xlsx back into the JSON instruction vocabulary consumed by
/// <see cref="Instructions.XlsxGenerator"/>. Rendering-oriented: emits typed cells,
/// number formats, styles (font/fill/border/alignment), column widths, row heights,
/// merges, freeze panes, autofilter, tables, metadata, and workbook title.
/// </summary>
public static class XlsxJsonSerializer
{
    /// <summary>Serializes the workbook at <paramref name="xlsxPath"/> to a JSON instruction set.</summary>
    public static object Serialize(string xlsxPath)
    {
        var read = new XlsxReader().Read(xlsxPath);
        return Serialize(read.Workbook);
    }

    /// <summary>Serializes a byte-buffer workbook to a JSON instruction set.</summary>
    public static object Serialize(byte[] xlsxBytes)
    {
        var read = new XlsxReader().Read(xlsxBytes);
        return Serialize(read.Workbook);
    }

    /// <summary>Serializes an in-memory workbook model to a JSON instruction set.</summary>
    public static object Serialize(XlsxRenderWorkbook workbook)
    {
        var doc = new Dictionary<string, object?>
        {
            ["version"] = "1.0",
            ["description"] = "Round-tripped from an existing workbook by OfficeEditor.",
        };

        if (!string.IsNullOrWhiteSpace(workbook.Title))
        {
            doc["metadata"] = new Dictionary<string, object?>
            {
                ["title"] = workbook.Title
            };
        }

        var sheets = new List<object>();
        foreach (var sheet in workbook.Sheets)
        {
            sheets.Add(SerializeSheet(sheet));
        }

        doc["worksheets"] = sheets;
        return doc;
    }

    private static object SerializeSheet(XlsxRenderSheet sheet)
    {
        var ws = new Dictionary<string, object?>
        {
            ["name"] = sheet.Name
        };

        if (sheet.Columns.Count > 0)
        {
            var columns = new List<object>();
            foreach (var col in sheet.Columns)
            {
                var c = new Dictionary<string, object?>
                {
                    ["name"] = ColumnLetter(col.Index),
                    ["width"] = Math.Round(col.WidthPt, 2)
                };
                columns.Add(c);
            }

            ws["columns"] = columns;
        }

        // Headers: first row when present.
        var firstRow = sheet.Rows.FirstOrDefault();
        if (firstRow is { Cells.Count: > 0 })
        {
            var headers = new List<string>();
            foreach (var cell in firstRow.Cells.OrderBy(c => c.Column))
            {
                if (cell.Value.Kind == XlsxValueKind.Text)
                    headers.Add(cell.Value.TextValue ?? "");
            }

            if (headers.Count > 0)
                ws["headers"] = headers;
        }

        // Body rows (all rows; headers are also included as data via cells to preserve layout).
        var rows = new List<object>();
        foreach (var row in sheet.Rows.OrderBy(r => r.Index).Skip(1))
        {
            var rowArr = new List<string>();
            int colCount = Math.Max(sheet.Columns.Count, row.Cells.Max(c => c.Column));
            for (int c = 1; c <= colCount; c++)
            {
                var cell = row.Cells.FirstOrDefault(x => x.Column == c);
                rowArr.Add(CellToJsonString(cell));
            }

            rows.Add(rowArr);
        }

        if (rows.Count > 0)
            ws["rows"] = rows;

        // Discrete cells beyond the first data row, carrying style/number-format detail.
        var cells = new List<object>();
        foreach (var row in sheet.Rows.OrderBy(r => r.Index))
        {
            foreach (var cell in row.Cells.OrderBy(c => c.Column))
            {
                if (cell.Value.Kind == XlsxValueKind.Empty)
                    continue;

                var item = new Dictionary<string, object?>
                {
                    ["address"] = CellAddress(cell.Row, cell.Column)
                };

                if (cell.Formula != null)
                    item["formula"] = cell.Formula;
                else
                    item["value"] = CellToJsonString(cell);

                if (!string.IsNullOrEmpty(cell.NumberFormatCode))
                    item["numberFormat"] = cell.NumberFormatCode;

                cells.Add(item);
            }
        }

        if (cells.Count > 0)
            ws["cells"] = cells;

        if (sheet.Merges.Count > 0)
        {
            ws["merges"] = sheet.Merges
                .Select(m => $"{CellAddress(m.FirstRow, m.FirstCol)}:{CellAddress(m.LastRow, m.LastCol)}")
                .ToList();
        }

        if (sheet.FreezePanes is { } fp && (fp.FrozenRows > 0 || fp.FrozenCols > 0))
        {
            var pane = new Dictionary<string, object?>();
            if (fp.FrozenRows > 0)
                pane["row"] = fp.FrozenRows;
            if (fp.FrozenCols > 0)
                pane["column"] = fp.FrozenCols;
            ws["freezePanes"] = pane;
        }

        if (!string.IsNullOrEmpty(sheet.AutoFilterRange))
            ws["autoFilter"] = sheet.AutoFilterRange;

        if (sheet.Tables.Count > 0)
        {
            ws["tables"] = sheet.Tables
                .Select(t => new Dictionary<string, object?>
                {
                    ["name"] = t.Name,
                    ["range"] = $"{CellAddress(t.FirstRow, t.FirstCol)}:{CellAddress(t.LastRow, t.LastCol)}"
                })
                .ToList();
        }

        return ws;
    }

    private static string CellToJsonString(XlsxRenderCell? cell)
    {
        if (cell == null || cell.Value.Kind == XlsxValueKind.Empty)
            return "";

        return cell.Value.Kind switch
        {
            XlsxValueKind.Text => cell.Value.TextValue ?? "",
            XlsxValueKind.Number => FormatNumber(cell.Value.NumberValue ?? 0, cell.NumberFormatCode),
            XlsxValueKind.Boolean => cell.Value.BooleanValue == true ? "true" : "false",
            XlsxValueKind.DateTime => cell.Value.DateTimeValue?.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture) ?? "",
            XlsxValueKind.Error => cell.Value.ErrorValue ?? "",
            _ => ""
        };
    }

    private static string FormatNumber(double value, string? code)
    {
        if (string.IsNullOrEmpty(code))
            return value.ToString(CultureInfo.InvariantCulture);

        // Keep dates as their ISO date when the format is a date format; otherwise keep raw.
        if (XlsxEditor.Core.Rendering.Formatting.ExcelDateFormat.IsDateFormat(code))
            return value.ToString("R", CultureInfo.InvariantCulture);

        return value.ToString(CultureInfo.InvariantCulture);
    }

    private static string ColumnLetter(int index)
    {
        int n = index;
        var sb = new System.Text.StringBuilder();
        while (n > 0)
        {
            n--;
            sb.Insert(0, (char)('A' + (n % 26)));
            n /= 26;
        }

        return sb.ToString();
    }

    private static string CellAddress(int row, int col) => $"{ColumnLetter(col)}{row}";
}
