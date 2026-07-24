using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Models;
using OfficeEditor.Core.Models;

namespace DocxEditor.Core.Content;

public class ContentBlockRenderer
{
    private readonly StyleMapping _styleMapping;
    private readonly Dictionary<string, Style> _cachedStyles;
    private readonly Action<string>? _ensureStyle;
    private readonly Func<bool, int>? _allocateNumberingId;

    /// <param name="allocateNumberingId">
    /// Allocates a fresh, collision-free numbering instance id for one list (argument: ordered?).
    /// Each list must get its own instance so counters restart and lists never bind to
    /// numbering definitions that already exist in the document. When null (standalone
    /// rendering against a detached body), legacy ids 1 (ordered) / 2 (bullet) are used.
    /// </param>
    public ContentBlockRenderer(
        StyleMapping? styleMapping = null,
        Dictionary<string, Style>? cachedStyles = null,
        Action<string>? ensureStyle = null,
        Func<bool, int>? allocateNumberingId = null)
    {
        _styleMapping = styleMapping ?? StyleMapping.Default;
        _cachedStyles = cachedStyles ?? new Dictionary<string, Style>();
        _ensureStyle = ensureStyle;
        _allocateNumberingId = allocateNumberingId;
    }

    public void Render(Body body, List<ContentBlock> blocks)
    {
        foreach (var block in blocks)
        {
            RenderBlock(body, block);
        }
    }

    private void RenderBlock(Body body, ContentBlock block)
    {
        switch (block)
        {
            case ParagraphBlock paragraph:
                RenderParagraph(body, paragraph);
                break;
            case HeadingBlock heading:
                RenderHeading(body, heading);
                break;
            case ListBlock list:
                RenderList(body, list);
                break;
            case TableBlock table:
                RenderTable(body, table);
                break;
            case BlockquoteBlock blockquote:
                RenderBlockquote(body, blockquote);
                break;
            case CodeBlock code:
                RenderCode(body, code);
                break;
            case HorizontalRuleBlock:
                RenderHorizontalRule(body);
                break;
            case CustomBlock custom:
                RenderCustom(body, custom);
                break;
        }
    }

    private void RenderParagraph(Body body, ParagraphBlock block)
    {
        var paragraph = new Paragraph();
        
        if (block.InlineFormats != null && block.InlineFormats.Count > 0)
        {
            foreach (var inline in block.InlineFormats)
            {
                var run = CreateRun(inline.Text, inline.Type);
                paragraph.Append(run);
            }
        }
        else
        {
            var run = new Run(new Text(block.Text));
            paragraph.Append(run);
        }

        var style = block.Style ?? _styleMapping.GetStyle("paragraph");
        if (!string.IsNullOrEmpty(style))
        {
            _ensureStyle?.Invoke(style);
            paragraph.ParagraphProperties = new ParagraphProperties(
                new ParagraphStyleId { Val = style }
            );
        }

        body.Append(paragraph);
    }

    private void RenderHeading(Body body, HeadingBlock block)
    {
        var paragraph = new Paragraph();
        var run = new Run(new Text(block.Text));
        paragraph.Append(run);

        var style = block.Style ?? _styleMapping.GetStyle($"heading{block.Level}");
        if (!string.IsNullOrEmpty(style))
        {
            _ensureStyle?.Invoke(style);
            paragraph.ParagraphProperties = new ParagraphProperties(
                new ParagraphStyleId { Val = style }
            );
        }

        body.Append(paragraph);
    }

    private void RenderList(Body body, ListBlock block)
    {
        // One numbering instance per list: ids 1/2 can already be taken (or mean a
        // different format) in documents that contain lists, and reusing an instance
        // would continue its counter instead of restarting at 1.
        var numberingId = _allocateNumberingId?.Invoke(block.Ordered) ?? (block.Ordered ? 1 : 2);

        for (int i = 0; i < block.Items.Count; i++)
        {
            var paragraph = new Paragraph();
            var run = new Run(new Text(block.Items[i]));
            paragraph.Append(run);

            // Add list properties
            var numberingProperties = new NumberingProperties(
                new NumberingLevelReference { Val = 0 },
                new NumberingId { Val = numberingId }
            );
            
            paragraph.ParagraphProperties = new ParagraphProperties(numberingProperties);

            var style = block.Style ?? _styleMapping.GetStyle("paragraph");
            if (!string.IsNullOrEmpty(style))
            {
                _ensureStyle?.Invoke(style);
                paragraph.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId { Val = style };
            }

            body.Append(paragraph);
        }
    }

    private void RenderTable(Body body, TableBlock block)
    {
        // Empty-table semantics: a table with no rows (or no cells at all) renders nothing.
        // Emitting it would produce a corrupt document because CT_Tbl requires tblGrid
        // (and at least one row) after tblPr.
        var columnCount = block.Rows.Count == 0 ? 0 : block.Rows.Max(r => r.Cells.Count);
        if (columnCount == 0)
        {
            return;
        }

        var table = new Table();

        // Add table properties (CT_TblBorders sequence: top, left, bottom, right, insideH, insideV)
        var tableProperties = new TableProperties(
            new TableBorders(
                new TopBorder { Val = new DocumentFormat.OpenXml.EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new LeftBorder { Val = new DocumentFormat.OpenXml.EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new BottomBorder { Val = new DocumentFormat.OpenXml.EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new RightBorder { Val = new DocumentFormat.OpenXml.EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new InsideHorizontalBorder { Val = new DocumentFormat.OpenXml.EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new InsideVerticalBorder { Val = new DocumentFormat.OpenXml.EnumValue<BorderValues>(BorderValues.Single), Size = 4 }
            )
        );
        table.Append(tableProperties);

        // CT_Tbl requires tblGrid immediately after tblPr; one gridCol per column.
        var tableGrid = new TableGrid();
        for (int i = 0; i < columnCount; i++)
        {
            tableGrid.Append(new GridColumn());
        }
        table.Append(tableGrid);

        foreach (var row in block.Rows)
        {
            var tableRow = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
            foreach (var cell in row.Cells)
            {
                var tableCell = new DocumentFormat.OpenXml.Wordprocessing.TableCell();
                var paragraph = new Paragraph(new Run(new Text(cell.Text)));
                tableCell.Append(paragraph);
                tableRow.Append(tableCell);
            }
            table.Append(tableRow);
        }

        body.Append(table);
    }

    private void RenderBlockquote(Body body, BlockquoteBlock block)
    {
        var paragraph = new Paragraph();
        var run = new Run(new Text(block.Text));
        paragraph.Append(run);

        var style = block.Style ?? _styleMapping.GetStyle("blockquote");
        if (!string.IsNullOrEmpty(style))
        {
            _ensureStyle?.Invoke(style);
            paragraph.ParagraphProperties = new ParagraphProperties(
                new ParagraphStyleId { Val = style }
            );
        }

        body.Append(paragraph);
    }

    private void RenderCode(Body body, CodeBlock block)
    {
        var paragraph = new Paragraph();
        var run = new Run(new Text(block.Text));
        paragraph.Append(run);

        var style = block.Style ?? _styleMapping.GetStyle("code");
        if (!string.IsNullOrEmpty(style))
        {
            _ensureStyle?.Invoke(style);
            paragraph.ParagraphProperties = new ParagraphProperties(
                new ParagraphStyleId { Val = style }
            );
        }

        body.Append(paragraph);
    }

    private void RenderHorizontalRule(Body body)
    {
        var paragraph = new Paragraph(
            new ParagraphProperties(
                new ParagraphBorders(
                    new BottomBorder 
                    { 
                        Val = new DocumentFormat.OpenXml.EnumValue<BorderValues>(BorderValues.Single),
                        Size = 6,
                        Space = 1
                    }
                )
            )
        );
        body.Append(paragraph);
    }

    private void RenderCustom(Body body, CustomBlock block)
    {
        var paragraph = new Paragraph();
        var run = new Run(new Text(block.Text));
        paragraph.Append(run);

        var style = block.Style ?? _styleMapping.GetStyle(block.CustomType);
        if (!string.IsNullOrEmpty(style))
        {
            _ensureStyle?.Invoke(style);
            paragraph.ParagraphProperties = new ParagraphProperties(
                new ParagraphStyleId { Val = style }
            );
        }

        body.Append(paragraph);
    }

    private Run CreateRun(string text, string formatType)
    {
        var run = new Run();
        var runProperties = new RunProperties();

        switch (formatType.ToLowerInvariant())
        {
            case "bold":
                runProperties.Append(new Bold());
                break;
            case "italic":
                runProperties.Append(new Italic());
                break;
            case "underline":
                runProperties.Append(new Underline { Val = UnderlineValues.Single });
                break;
            case "strikethrough":
                runProperties.Append(new Strike());
                break;
            case "code":
                _ensureStyle?.Invoke("Code");
                runProperties.Append(new RunStyle { Val = "Code" });
                break;
        }

        run.Append(runProperties);
        run.Append(new Text(text));
        return run;
    }
}
