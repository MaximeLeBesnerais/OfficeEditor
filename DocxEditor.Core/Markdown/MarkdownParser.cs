using Markdig;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using DocxEditor.Core.Models;

namespace DocxEditor.Core.Markdown;

public class MarkdownParser
{
    public List<ContentBlock> Parse(string markdown, StyleMapping? styleMap = null)
    {
        var pipeline = new MarkdownPipelineBuilder()
            .UseAdvancedExtensions()
            .Build();

        var document = Markdig.Markdown.Parse(markdown, pipeline);
        var blocks = new List<ContentBlock>();

        foreach (var block in document)
        {
            var contentBlock = ConvertBlock(block, styleMap);
            if (contentBlock != null)
            {
                blocks.Add(contentBlock);
            }
        }

        return blocks;
    }

    private ContentBlock? ConvertBlock(Markdig.Syntax.Block block, StyleMapping? styleMap)
    {
        switch (block)
        {
            case Markdig.Syntax.HeadingBlock heading:
                return new Models.HeadingBlock
                {
                    Level = heading.Level,
                    Text = GetInlineText(heading.Inline),
                    Style = styleMap?.GetStyle($"heading{heading.Level}")
                };

            case Markdig.Syntax.ParagraphBlock paragraph:
                var inlineFormats = ExtractInlineFormats(paragraph.Inline);
                return new Models.ParagraphBlock
                {
                    Text = GetInlineText(paragraph.Inline),
                    InlineFormats = inlineFormats,
                    Style = styleMap?.GetStyle("paragraph")
                };

            case Markdig.Syntax.ListBlock list:
                var items = new List<string>();
                foreach (var item in list)
                {
                    if (item is ListItemBlock listItem)
                    {
                        foreach (var itemBlock in listItem)
                        {
                            if (itemBlock is Markdig.Syntax.ParagraphBlock itemParagraph)
                            {
                                items.Add(GetInlineText(itemParagraph.Inline));
                            }
                        }
                    }
                }
                return new Models.ListBlock
                {
                    Ordered = list.IsOrdered,
                    Items = items,
                    Style = styleMap?.GetStyle("paragraph")
                };

            case QuoteBlock quote:
                var quoteText = string.Join("\n", quote.Select(b => 
                    b is Markdig.Syntax.ParagraphBlock qb ? GetInlineText(qb.Inline) : ""));
                return new BlockquoteBlock
                {
                    Text = quoteText,
                    Style = styleMap?.GetStyle("blockquote")
                };

            case Markdig.Syntax.CodeBlock code:
                return new Models.CodeBlock
                {
                    Text = code.Lines.ToString(),
                    Language = (code as FencedCodeBlock)?.Info,
                    Style = styleMap?.GetStyle("code")
                };

            case ThematicBreakBlock:
                return new HorizontalRuleBlock();

            case Markdig.Extensions.Tables.Table table:
                return ConvertTable(table);

            default:
                return null;
        }
    }

    private TableBlock ConvertTable(Markdig.Extensions.Tables.Table table)
    {
        var rows = new List<Models.TableRow>();
        
        foreach (var row in table)
        {
            if (row is Markdig.Extensions.Tables.TableRow tableRow)
            {
                var cells = new List<Models.TableCell>();
                foreach (var cell in tableRow)
                {
                    if (cell is Markdig.Extensions.Tables.TableCell tableCell)
                    {
                        var cellText = string.Join(" ", tableCell.Select(b =>
                            b is Markdig.Syntax.ParagraphBlock pb ? GetInlineText(pb.Inline) : ""));
                        cells.Add(new Models.TableCell { Text = cellText });
                    }
                }
                rows.Add(new Models.TableRow { Cells = cells });
            }
        }

        return new TableBlock { Rows = rows };
    }

    private string GetInlineText(ContainerInline? inline)
    {
        if (inline == null) return string.Empty;
        
        var text = new System.Text.StringBuilder();
        foreach (var child in inline)
        {
            switch (child)
            {
                case LiteralInline literal:
                    text.Append(literal.Content.ToString());
                    break;
                case EmphasisInline emphasis:
                    text.Append(emphasis.ToString());
                    break;
                case CodeInline code:
                    text.Append(code.Content.ToString());
                    break;
                case LinkInline link:
                    text.Append(link.ToString());
                    break;
                default:
                    text.Append(child.ToString());
                    break;
            }
        }
        return text.ToString();
    }

    private List<InlineFormat> ExtractInlineFormats(ContainerInline? inline)
    {
        var formats = new List<InlineFormat>();
        if (inline == null) return formats;

        foreach (var child in inline)
        {
            switch (child)
            {
                case LiteralInline literal:
                    formats.Add(new InlineFormat { Type = "text", Text = literal.Content.ToString() });
                    break;
                case EmphasisInline emphasis:
                    var type = emphasis.DelimiterCount == 2 ? "bold" : "italic";
                    formats.Add(new InlineFormat { Type = type, Text = emphasis.ToString() });
                    break;
                case CodeInline code:
                    formats.Add(new InlineFormat { Type = "code", Text = code.Content.ToString() });
                    break;
                case LinkInline link:
                    formats.Add(new InlineFormat { Type = "text", Text = link.ToString() });
                    break;
                default:
                    formats.Add(new InlineFormat { Type = "text", Text = child.ToString() ?? "" });
                    break;
            }
        }

        return formats;
    }
}
