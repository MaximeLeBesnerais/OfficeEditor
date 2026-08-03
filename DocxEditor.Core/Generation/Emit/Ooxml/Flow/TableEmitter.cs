using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Generation.Design;
using DocxEditor.Core.Generation.Emit.Ooxml.Design;
using DocxEditor.Core.Generation.Model;
using Model = DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Flow;

/// <summary>
/// Emits rectangular-grid tables (no merged/nested cells): valid CT_Tbl ordering
/// (tblPr → tblGrid → rows), theme-aware light borders and cell padding, per-column widths when
/// declared, table alignment, automatic header-row treatment (theme fill + header style),
/// subtle row banding, repeating headers, row-split prevention and an oversized-table guardrail.
/// Explicit cell fills and content overrides always win over the theme treatment.
/// </summary>
internal static class TableEmitter
{
    public static void EmitTable(
        OoxmlEmitContext context,
        OpenXmlCompositeElement container,
        TableBlock table,
        string path,
        ResolvedPageFormat? pageFormat = null)
    {
        if (table.Rows.Count == 0)
        {
            context.Warn(path, "the table has no rows and was skipped.");
            return;
        }

        var design = context.DesignResolver.ResolveAll();
        var columnCount = table.Rows[0].Cells.Count;
        var widths = table.ColumnWidthsPt;

        WarnOversizedTable(context, table, widths, pageFormat, path);

        var headerFill = ThemeColor(design, "primary");
        var bandFill = ThemeColor(design, "pale");
        var borderColor = ThemeColor(design, "border") ?? "D9D9D9";

        var tableProperties = new TableProperties();

        // Border order follows CT_TblBorders: top, left, bottom, right, insideH, insideV.
        tableProperties.TableBorders = new TableBorders(
            new TopBorder { Val = BorderValues.Single, Size = 4, Color = borderColor },
            new LeftBorder { Val = BorderValues.Single, Size = 4, Color = borderColor },
            new BottomBorder { Val = BorderValues.Single, Size = 4, Color = borderColor },
            new RightBorder { Val = BorderValues.Single, Size = 4, Color = borderColor },
            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = borderColor },
            new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = borderColor });

        tableProperties.TableCellMarginDefault = new TableCellMarginDefault(
            new TopMargin { Width = "60", Type = TableWidthUnitValues.Dxa },
            new LeftMargin { Width = "100", Type = TableWidthUnitValues.Dxa },
            new BottomMargin { Width = "60", Type = TableWidthUnitValues.Dxa },
            new RightMargin { Width = "100", Type = TableWidthUnitValues.Dxa });

        if (context.ResolveStyle(table.Style, StyleValues.Table, path) is { } tableStyleId)
        {
            tableProperties.TableStyle = new TableStyle { Val = tableStyleId };
        }

        tableProperties.TableWidth = widths is not null
            ? new TableWidth { Width = FormattingHelpers.Twips(widths.Sum()), Type = TableWidthUnitValues.Dxa }
            : new TableWidth { Width = "0", Type = TableWidthUnitValues.Auto };

        if (FormattingHelpers.TableJustification(table.Alignment) is { } justification)
        {
            tableProperties.TableJustification = new TableJustification { Val = justification };
        }

        var docTable = new Table();
        docTable.Append(tableProperties);

        var grid = new TableGrid();
        for (var c = 0; c < columnCount; c++)
        {
            if (widths is null)
            {
                grid.Append(new GridColumn());
                continue;
            }
            grid.Append(new GridColumn { Width = FormattingHelpers.Twips(widths[c]) });
        }
        docTable.Append(grid);

        var bodyRowIndex = 0;
        for (var r = 0; r < table.Rows.Count; r++)
        {
            var row = table.Rows[r];
            var tableRow = new DocumentFormat.OpenXml.Wordprocessing.TableRow();

            var rowProperties = new TableRowProperties();
            // CT_TrPr sequence: cantSplit precedes tblHeader — match the Markdown renderer.
            rowProperties.Append(new CantSplit());
            if (row.IsHeader)
            {
                rowProperties.Append(new TableHeader());
            }
            tableRow.TableRowProperties = rowProperties;

            var banded = !row.IsHeader && (bodyRowIndex % 2) == 0;
            for (var c = 0; c < columnCount; c++)
            {
                var cell = row.Cells[c];
                tableRow.Append(BuildCell(context, cell, widths?[c], r, c, row.IsHeader, banded, headerFill, bandFill, path));
            }

            if (!row.IsHeader)
            {
                bodyRowIndex++;
            }

            docTable.Append(tableRow);
        }

        container.Append(docTable);
    }

    private static DocumentFormat.OpenXml.Wordprocessing.TableCell BuildCell(
        OoxmlEmitContext context,
        Model.TableCell cell,
        double? widthPt,
        int row,
        int column,
        bool isHeader,
        bool banded,
        string? headerFill,
        string? bandFill,
        string tablePath)
    {
        var tableCell = new DocumentFormat.OpenXml.Wordprocessing.TableCell();
        var cellProperties = new TableCellProperties();

        if (widthPt is { } width)
        {
            cellProperties.TableCellWidth = new TableCellWidth { Width = FormattingHelpers.Twips(width), Type = TableWidthUnitValues.Dxa };
        }

        var cellPath = $"{tablePath}.rows[{row}].cells[{column}]";
        var fillHex = ResolveCellFill(context, cell.Fill, isHeader, banded, headerFill, bandFill, cellPath);
        if (fillHex is { } fill)
        {
            cellProperties.Shading = new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = fill };
        }

        if (cellProperties.HasChildren)
        {
            tableCell.Append(cellProperties);
        }

        var cellStyleId = cell.Content?.Role is { } role
            ? context.StyleManager.GetOrCreateStyle(role.ToBaselineStyleKind())
            : isHeader
                ? context.StyleManager.GetOrCreateTableHeaderStyle()
                : context.StyleManager.GetOrCreateStyle(BaselineStyleKind.TableBody);

        if (cell.Content is { } content)
        {
            var cellOptions = new ParagraphBuildOptions
            {
                DefaultStyleKind = isHeader ? BaselineStyleKind.TableHeader : BaselineStyleKind.TableBody,
                ApplyBodySpacingDefaults = false
            };
            var paragraph = ParagraphEmitter.BuildParagraph(context, content, cellStyleId, null, cellPath, cellOptions);
            if (cell.Alignment is { } alignment && FormattingHelpers.Justification(alignment) is { } justification)
            {
                paragraph.ParagraphProperties ??= new ParagraphProperties();
                paragraph.ParagraphProperties.Justification = new Justification { Val = justification };
            }
            tableCell.Append(paragraph);
        }
        else
        {
            // A blank cell still needs at least one paragraph to be valid.
            tableCell.Append(new Paragraph());
        }

        return tableCell;
    }

    /// <summary>
    /// Resolves a cell's fill with the theme treatment as default: explicit fills win, header rows
    /// fall back to the primary fill, and every other body row bands to the pale surface.
    /// </summary>
    private static string? ResolveCellFill(
        OoxmlEmitContext context,
        string? explicitFill,
        bool isHeader,
        bool banded,
        string? headerFill,
        string? bandFill,
        string cellPath)
    {
        if (explicitFill is not null)
        {
            return FormattingHelpers.ResolveColor(context, explicitFill, cellPath);
        }
        if (isHeader)
        {
            return headerFill;
        }
        return banded ? bandFill : null;
    }

    /// <summary>
    /// Guardrail: warns when the declared table width exceeds the layout's maximum (or the section's
    /// text width when no explicit maximum is set). Advisory only — no pagination prediction.
    /// </summary>
    private static void WarnOversizedTable(
        OoxmlEmitContext context,
        TableBlock table,
        IReadOnlyList<double>? widths,
        ResolvedPageFormat? pageFormat,
        string path)
    {
        if (widths is null)
        {
            return;
        }
        var tableWidth = widths.Sum();
        var layout = context.DesignResolver.ResolveAll().Layout;
        var textWidth = pageFormat?.TextWidthPt ?? context.DesignResolver.ResolvePage(null, path).TextWidthPt;
        var threshold = layout.MaxTableWidthPt ?? textWidth;
        if (tableWidth > threshold)
        {
            context.Warn(
                path,
                $"table width {FormatNumber(tableWidth)}pt exceeds the recommended maximum of {FormatNumber(threshold)}pt; the table may overflow the text area.");
        }
    }

    /// <summary>A palette token's OOXML color (no leading '#'), when the design defines it.</summary>
    private static string? ThemeColor(ResolvedDesign design, string token) =>
        design.Palette.TryGetValue(token, out var hex) ? DocxFormattingHelpers.ToOoxmlColor(hex) : null;

    private static string FormatNumber(double value) =>
        value.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);
}
