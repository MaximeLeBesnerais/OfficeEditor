using System.Text;
using XlsxEditor.Core.Rendering.Formatting;
using XlsxEditor.Core.Rendering.Models;

namespace XlsxEditor.Core.Rendering.Emit;

/// <summary>
/// Converts an immutable <see cref="XlsxRenderWorkbook"/> into Typst source that renders
/// the sheets as native Typst tables (fixed column widths, row heights, merged cells,
/// fills, per-edge borders, alignment, repeating header rows, and pagination).
///
/// This is the repository's own spreadsheet renderer — no external office suite is used.
/// </summary>
public sealed class XlsxToTypstConverter
{
    private const double DefaultColumnWidthPt = 64;
    private const double DefaultRowHeightPt = 16;
    private const double PageMarginPt = 36;

    private readonly XlsxRenderWorkbook _workbook;

    public XlsxToTypstConverter(XlsxRenderWorkbook workbook)
    {
        _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
    }

    /// <summary>Renders the whole workbook to a single Typst source document.</summary>
    public string GenerateTypstSource()
    {
        var sb = new StringBuilder();

        sb.AppendLine("#set text(font: \"Arial\", size: 9pt)");
        sb.AppendLine("#set page(numbering: \"1 / 1\")");
        sb.AppendLine();

        var sheets = _workbook.Sheets;
        for (int i = 0; i < sheets.Count; i++)
        {
            if (i > 0)
                sb.AppendLine("#pagebreak()");

            EmitSheet(sb, sheets[i]);
            sb.AppendLine();
        }

        return sb.ToString();
    }

    private void EmitSheet(StringBuilder sb, XlsxRenderSheet sheet)
    {
        double pageWidth = sheet.PageWidthPt > 0 ? sheet.PageWidthPt : 595;
        double pageHeight = sheet.PageHeightPt > 0 ? sheet.PageHeightPt : 842;

        if (sheet.Landscape && pageHeight > pageWidth)
            (pageWidth, pageHeight) = (pageHeight, pageWidth);

        // Sheet title as a small bold heading.
        if (!string.IsNullOrEmpty(sheet.Name))
        {
            sb.Append("#text(size: 12pt, weight: \"bold\")[");
            sb.Append(EscapeTypst(sheet.Name));
            sb.AppendLine("]");
            sb.AppendLine();
        }

        sb.Append("#set page(width: ");
        sb.Append(FormatPt(pageWidth));
        sb.Append(", height: ");
        sb.Append(FormatPt(pageHeight));
        sb.Append(", margin: ");
        sb.Append(FormatPt(PageMarginPt));
        sb.AppendLine(")");
        sb.AppendLine();

        EmitTable(sb, sheet);
    }

    private void EmitTable(StringBuilder sb, XlsxRenderSheet sheet)
    {
        if (sheet.Columns.Count == 0 && sheet.Rows.Count == 0)
        {
            sb.AppendLine("#table(columns: 1, [ ])");
            return;
        }

        int columnCount = Math.Max(sheet.Columns.Count, ComputeUsedColumns(sheet));

        // Column widths (points) and per-column alignment defaults.
        var widths = new List<string>(columnCount);
        var aligns = new List<string>(columnCount);
        var hiddenCols = new HashSet<int>();

        for (int c = 0; c < columnCount; c++)
        {
            var col = sheet.Columns.FirstOrDefault(x => x.Index == c + 1);
            double width = col?.WidthPt > 0 ? col.WidthPt : DefaultColumnWidthPt;
            if (col is { Hidden: true })
                hiddenCols.Add(c);
            widths.Add(FormatPt(width));
            aligns.Add(DefaultColumnAlignment(sheet, c + 1));
        }

        sb.Append("#table(");
        sb.Append("columns: (");
        sb.Append(string.Join(", ", widths));
        sb.Append("), ");

        // Explicit row heights: Typst's `rows:` array needs one entry per row, so emit a
        // complete array using `auto` for rows that carry the default height.
        bool anyExplicitHeight = sheet.Rows.Any(r => r.HeightPt > 0 && r.HeightPt != DefaultRowHeightPt);
        if (anyExplicitHeight && sheet.Rows.Count > 0)
        {
            sb.Append("rows: (");
            sb.Append(string.Join(", ", sheet.Rows.Select(r =>
                r.HeightPt > 0 && r.HeightPt != DefaultRowHeightPt ? FormatPt(r.HeightPt) : "auto")));
            sb.Append("), ");
        }

        // Per-column alignment: numbers right, text left (overridden per cell when styled).
        sb.Append("align: (");
        sb.Append(string.Join(", ", aligns));
        sb.Append("), ");
        sb.AppendLine("stroke: 0.25pt + rgb(\"C8C8C8\"),");

        int headerRows = HeaderRowCount(sheet);
        // (Row, Col) -> remaining rows occupied by a rowspan that started above.
        var occupiedByRowspan = new Dictionary<(int Row, int Col), int>();

        int rowIndex = 0;
        foreach (var row in sheet.Rows)
        {
            if (row.Hidden)
                continue;

            bool inHeader = headerRows > 0 && rowIndex < headerRows;
            if (inHeader && rowIndex == 0)
                sb.Append("  table.header(repeat: true,");

            var rowCells = new List<string>();
            // Columns already consumed by a colspan on this row.
            var coveredByColspan = new HashSet<int>();

            foreach (var col in Enumerable.Range(0, columnCount))
            {
                if (hiddenCols.Contains(col))
                    continue;

                if (coveredByColspan.Contains(col))
                    continue;

                if (occupiedByRowspan.TryGetValue((rowIndex, col), out int remaining))
                {
                    if (remaining > 1)
                        occupiedByRowspan[(rowIndex, col)] = remaining - 1;
                    else
                        occupiedByRowspan.Remove((rowIndex, col));
                    continue; // covered by a rowspan from an earlier row
                }

                // Is this the top-left of a merge?
                var merge = FindMerge(sheet, rowIndex, col);
                if (merge != null)
                {
                    int colspan = merge.LastCol - merge.FirstCol + 1;
                    int rowspan = merge.LastRow - merge.FirstRow + 1;

                    for (int rr = 1; rr < rowspan; rr++)
                        occupiedByRowspan[(rowIndex + rr, merge.FirstCol)] = rowspan - rr;
                    for (int cc = 1; cc < colspan; cc++)
                        coveredByColspan.Add(merge.FirstCol + cc - 1);

                    var cell = FindCell(sheet, merge.FirstRow - 1, merge.FirstCol - 1);
                    rowCells.Add(BuildCell(cell, colspan, rowspan, sheet));
                }
                else
                {
                    var cell = FindCell(sheet, rowIndex, col);
                    rowCells.Add(BuildCell(cell, 1, 1, sheet));
                }
            }

            if (inHeader)
            {
                if (rowIndex < headerRows - 1)
                {
                    sb.Append(' ');
                    sb.Append(string.Join(", ", rowCells));
                    sb.Append(", ");
                }
                else
                {
                    sb.Append(' ');
                    sb.Append(string.Join(", ", rowCells));
                    sb.AppendLine("),");
                }
            }
            else
            {
                sb.Append("  ");
                sb.Append(string.Join(", ", rowCells));
                sb.AppendLine(",");
            }

            rowIndex++;
        }

        sb.AppendLine(")");
    }

    private static string BuildCell(XlsxRenderCell? cell, int colspan, int rowspan,
        XlsxRenderSheet sheet)
    {
        var parts = new List<string>();
        if (colspan > 1)
            parts.Add($"colspan: {colspan}");
        if (rowspan > 1)
            parts.Add($"rowspan: {rowspan}");

        var style = cell?.Style;
        if (style != null)
        {
            var stroke = BuildStroke(style);
            if (!string.IsNullOrEmpty(stroke))
                parts.Add($"stroke: {stroke}");

            if (!string.IsNullOrEmpty(style.FillArgb))
                parts.Add($"fill: rgb(\"{ToHex(style.FillArgb)}\")");
        }

        string content = cell == null
            ? ""
            : RenderCellContent(cell);

        string? align = cell?.Style?.HorizontalAlignment switch
        {
            "center" => "center + horizon",
            "right" => "right + horizon",
            _ => null
        } ?? (cell != null && IsNumeric(cell) ? "right + horizon" : null);

        if (align != null)
            parts.Add($"align: {align}");

        if (parts.Count == 0)
            return $"[{content}]";

        return $"table.cell({string.Join(", ", parts)})[{content}]";
    }

    private static string BuildStroke(XlsxRenderStyle style)
    {
        var edges = new List<string>();

        void AddEdge(string name, XlsxRenderBorderEdge? edge)
        {
            if (edge == null)
                return;
            double thickness = BorderThickness(edge.Style);
            string color = edge.ColorArgb != null ? ToHex(edge.ColorArgb) : "000000";
            edges.Add($"{name}: {FormatPt(thickness)} + rgb(\"{color}\")");
        }

        AddEdge("top", style.BorderTop);
        AddEdge("bottom", style.BorderBottom);
        AddEdge("left", style.BorderLeft);
        AddEdge("right", style.BorderRight);

        return edges.Count == 0 ? "" : $"({string.Join(", ", edges)})";
    }

    private static double BorderThickness(string? style) => style?.ToLowerInvariant() switch
    {
        "hair" or "thin" or "dotted" or "dashed" => 0.25,
        "medium" => 1.0,
        "thick" => 1.5,
        "double" => 1.5,
        _ => 0.5
    };

    private static string RenderCellContent(XlsxRenderCell cell)
    {
        var value = cell.Value;
        if (value == null)
            return "";

        switch (value.Kind)
        {
            case XlsxValueKind.Empty:
                return "";

            case XlsxValueKind.Number:
            {
                double n = value.NumberValue ?? 0;
                string? code = cell.NumberFormatCode ?? cell.Style?.NumberFormatCode;
                if (!string.IsNullOrEmpty(code))
                {
                    var formatted = ExcelNumberFormatFormatter.FormatNumber(n, code);
                    if (formatted.ColorArgb != null)
                        return $"text(fill: rgb(\"{ToHex(formatted.ColorArgb)}\"))[{EscapeTypst(formatted.Text)}]";
                    return EscapeTypst(formatted.Text);
                }

                return EscapeTypst(ExcelNumberFormatFormatter.FormatNumber(n, "General").Text);
            }

            case XlsxValueKind.Boolean:
                return value.BooleanValue == true ? "TRUE" : "FALSE";

            case XlsxValueKind.DateTime:
            {
                var dt = value.DateTimeValue ?? DateTime.MinValue;
                string? code = cell.NumberFormatCode ?? cell.Style?.NumberFormatCode;
                if (!string.IsNullOrEmpty(code) && ExcelDateFormat.IsDateFormat(code))
                    return EscapeTypst(ExcelDateFormat.FormatDate(dt, code));
                return EscapeTypst(dt.ToString("yyyy-MM-dd"));
            }

            case XlsxValueKind.Text:
                return EscapeTypst(value.TextValue ?? "");

            case XlsxValueKind.Error:
                return EscapeTypst(value.ErrorValue ?? "");

            default:
                return "";
        }
    }

    private static bool IsNumeric(XlsxRenderCell cell) =>
        cell.Value.Kind is XlsxValueKind.Number or XlsxValueKind.DateTime;

    private static string DefaultColumnAlignment(XlsxRenderSheet sheet, int column)
    {
        // Heuristic: if the majority of populated cells in this column are numeric, right-align.
        int numeric = 0;
        int populated = 0;
        foreach (var row in sheet.Rows)
        {
            var cell = FindCell(sheet, row.Index - 1, column - 1);
            if (cell == null || cell.Value.Kind == XlsxValueKind.Empty)
                continue;
            populated++;
            if (cell.Value.Kind is XlsxValueKind.Number or XlsxValueKind.DateTime)
                numeric++;
        }

        return populated > 0 && numeric >= populated / 2.0 ? "right + horizon" : "left + horizon";
    }

    private static int ComputeUsedColumns(XlsxRenderSheet sheet)
    {
        int max = 0;
        foreach (var row in sheet.Rows)
        {
            foreach (var cell in row.Cells)
                max = Math.Max(max, cell.Column);
        }

        foreach (var merge in sheet.Merges)
            max = Math.Max(max, merge.LastCol);

        return max;
    }

    private static int HeaderRowCount(XlsxRenderSheet sheet)
    {
        if (sheet.FreezePanes is { FrozenRows: > 0 } fp)
            return fp.FrozenRows;

        // Fall back to repeating the first row when the sheet has a table with a header,
        // or when the first row looks like a header (bold). Prefer simplicity: only freeze.
        return 0;
    }

    private static XlsxRenderMerge? FindMerge(XlsxRenderSheet sheet, int row0, int col0)
    {
        foreach (var m in sheet.Merges)
        {
            if (m.FirstRow == row0 + 1 && m.FirstCol == col0 + 1)
                return m;
        }

        return null;
    }

    private static XlsxRenderCell? FindCell(XlsxRenderSheet sheet, int row0, int col0)
    {
        foreach (var row in sheet.Rows)
        {
            if (row.Index != row0 + 1)
                continue;
            foreach (var cell in row.Cells)
            {
                if (cell.Column == col0 + 1)
                    return cell;
            }
        }

        return null;
    }

    private static string FormatPt(double pt) => pt.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture) + "pt";

    private static string ToHex(string argb)
    {
        // ARGB (AARRGGBB) -> RRGGBB for Typst rgb("#RRGGBB").
        var s = argb.TrimStart('#');
        if (s.Length == 8)
            s = s[2..];
        return s;
    }

    /// <summary>Escapes Typst markup special characters for literal content blocks.</summary>
    private static string EscapeTypst(string text)
    {
        var sb = new StringBuilder(text.Length + 8);
        foreach (char c in text)
        {
            switch (c)
            {
                case '\\': sb.Append('\\'); sb.Append('\\'); break;
                case '#': sb.Append("\\#"); break;
                case '$': sb.Append("\\$"); break;
                case '%': sb.Append("\\%"); break;
                case '&': sb.Append("\\&"); break;
                case '@': sb.Append("\\@"); break;
                case '*': sb.Append("\\*"); break;
                case '_': sb.Append("\\_"); break;
                case '^': sb.Append("\\^"); break;
                case '~': sb.Append("\\~"); break;
                case '<': sb.Append("\\<"); break;
                case '>': sb.Append("\\>"); break;
                case '[': sb.Append("\\["); break;
                case ']': sb.Append("\\]"); break;
                default: sb.Append(c); break;
            }
        }

        return sb.ToString();
    }
}
