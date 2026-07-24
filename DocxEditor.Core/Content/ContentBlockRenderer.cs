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
    private readonly Action? _ensureNumbering;

    public ContentBlockRenderer(
        StyleMapping? styleMapping = null,
        Dictionary<string, Style>? cachedStyles = null,
        Action<string>? ensureStyle = null,
        Action? ensureNumbering = null)
    {
        _styleMapping = styleMapping ?? StyleMapping.Default;
        _cachedStyles = cachedStyles ?? new Dictionary<string, Style>();
        _ensureStyle = ensureStyle;
        _ensureNumbering = ensureNumbering;
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
        _ensureNumbering?.Invoke();

        for (int i = 0; i < block.Items.Count; i++)
        {
            var paragraph = new Paragraph();
            var run = new Run(new Text(block.Items[i]));
            paragraph.Append(run);

            // Add list properties
            var numberingProperties = new NumberingProperties(
                new NumberingLevelReference { Val = 0 },
                new NumberingId { Val = block.Ordered ? 1 : 2 }
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
        var table = new Table();
        
        // Add table properties
        var tableProperties = new TableProperties(
            new TableBorders(
                new TopBorder { Val = new DocumentFormat.OpenXml.EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new BottomBorder { Val = new DocumentFormat.OpenXml.EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new LeftBorder { Val = new DocumentFormat.OpenXml.EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new RightBorder { Val = new DocumentFormat.OpenXml.EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new InsideHorizontalBorder { Val = new DocumentFormat.OpenXml.EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new InsideVerticalBorder { Val = new DocumentFormat.OpenXml.EnumValue<BorderValues>(BorderValues.Single), Size = 4 }
            )
        );
        table.Append(tableProperties);

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
                runProperties.Append(new RunStyle { Val = "Code" });
                break;
        }

        run.Append(runProperties);
        run.Append(new Text(text));
        return run;
    }
}
