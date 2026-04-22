using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Models;

namespace DocxEditor.Core.Content;

public class ContentBlockBuilder
{
    private readonly List<ContentBlock> _blocks = new();

    public ContentBlockBuilder AddParagraph(string text, string? style = null)
    {
        _blocks.Add(new ParagraphBlock
        {
            Text = text,
            Style = style
        });
        return this;
    }

    public ContentBlockBuilder AddParagraph(string text, List<InlineFormat> inlineFormats, string? style = null)
    {
        _blocks.Add(new ParagraphBlock
        {
            Text = text,
            InlineFormats = inlineFormats,
            Style = style
        });
        return this;
    }

    public ContentBlockBuilder AddHeading(int level, string text, string? style = null)
    {
        _blocks.Add(new HeadingBlock
        {
            Level = level,
            Text = text,
            Style = style
        });
        return this;
    }

    public ContentBlockBuilder AddList(bool ordered, List<string> items, string? style = null)
    {
        _blocks.Add(new ListBlock
        {
            Ordered = ordered,
            Items = items,
            Style = style
        });
        return this;
    }

    public ContentBlockBuilder AddTable(List<List<string>> rows)
    {
        var tableRows = rows.Select(r => new Models.TableRow
        {
            Cells = r.Select(c => new Models.TableCell { Text = c }).ToList()
        }).ToList();

        _blocks.Add(new TableBlock
        {
            Rows = tableRows
        });
        return this;
    }

    public ContentBlockBuilder AddBlockquote(string text, string? style = null)
    {
        _blocks.Add(new BlockquoteBlock
        {
            Text = text,
            Style = style
        });
        return this;
    }

    public ContentBlockBuilder AddCode(string text, string? language = null, string? style = null)
    {
        _blocks.Add(new CodeBlock
        {
            Text = text,
            Language = language,
            Style = style
        });
        return this;
    }

    public ContentBlockBuilder AddHorizontalRule()
    {
        _blocks.Add(new HorizontalRuleBlock());
        return this;
    }

    public ContentBlockBuilder AddCustom(string customType, string text, string? style = null)
    {
        _blocks.Add(new CustomBlock
        {
            CustomType = customType,
            Text = text,
            Style = style
        });
        return this;
    }

    public List<ContentBlock> Build()
    {
        return _blocks;
    }
}
