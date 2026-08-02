using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Generation.Model;
using Model = DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Flow;

/// <summary>
/// Emits rectangular-grid tables: valid CT_Tbl ordering (tblPr → tblGrid → rows), per-column
/// widths when declared, table alignment, repeating header rows and cell fills/alignment.
/// </summary>
internal static class TableEmitter
{
    public static void EmitTable(OoxmlEmitContext context, OpenXmlCompositeElement container, TableBlock table, string path)
    {
        if (table.Rows.Count == 0)
        {
            context.Warn(path, "the table has no rows and was skipped.");
            return;
        }

        var columnCount = table.Rows[0].Cells.Count;
        var widths = table.ColumnWidthsPt;

        var tableProperties = new TableProperties();

        // Border order follows CT_TblBorders: top, left, bottom, right, insideH, insideV.
        tableProperties.TableBorders = new TableBorders(
            new TopBorder { Val = BorderValues.Single, Size = 4 },
            new LeftBorder { Val = BorderValues.Single, Size = 4 },
            new BottomBorder { Val = BorderValues.Single, Size = 4 },
            new RightBorder { Val = BorderValues.Single, Size = 4 },
            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
            new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 });

        if (context.ResolveStyle(table.Style, StyleValues.Table, path) is { } tableStyleId)
        {
            tableProperties.TableStyle = new TableStyle { Val = tableStyleId };
        }

        tableProperties.TableWidth = widths is not null
            ? new TableWidth { Width = FormattingHelpers.Twips(widths.Sum()), Type = TableWidthUnitValues.Dxa }
            : new TableWidth { Width = "auto", Type = TableWidthUnitValues.Auto };

        if (FormattingHelpers.TableJustification(table.Alignment) is { } justification)
        {
            tableProperties.TableJustification = new TableJustification { Val = justification };
        }

        var docTable = new Table();
        docTable.Append(tableProperties);

        var grid = new TableGrid();
        for (var c = 0; c < columnCount; c++)
        {
            grid.Append(new GridColumn
            {
                Width = widths is not null ? FormattingHelpers.Twips(widths[c]) : null
            });
        }
        docTable.Append(grid);

        for (var r = 0; r < table.Rows.Count; r++)
        {
            var row = table.Rows[r];
            var tableRow = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
            if (row.IsHeader)
            {
                tableRow.TableRowProperties = new TableRowProperties(new TableHeader());
            }

            for (var c = 0; c < columnCount; c++)
            {
                var cell = row.Cells[c];
                tableRow.Append(BuildCell(context, cell, widths?[c], r, c, path));
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
        string tablePath)
    {
        var tableCell = new DocumentFormat.OpenXml.Wordprocessing.TableCell();
        var cellProperties = new TableCellProperties();

        if (widthPt is { } width)
        {
            cellProperties.TableCellWidth = new TableCellWidth { Width = FormattingHelpers.Twips(width), Type = TableWidthUnitValues.Dxa };
        }

        if (cell.Fill is { } fill && FormattingHelpers.ResolveColor(context, fill, $"{tablePath}.rows[{row}].cells[{column}].fill") is { } hex)
        {
            cellProperties.Shading = new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = hex };
        }

        if (cellProperties.HasChildren)
        {
            tableCell.Append(cellProperties);
        }

        var cellPath = $"{tablePath}.rows[{row}].cells[{column}]";
        if (cell.Content is { } content)
        {
            var paragraph = ParagraphEmitter.BuildParagraph(context, content, null, null, cellPath);
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
}
