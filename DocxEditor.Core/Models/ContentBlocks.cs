namespace DocxEditor.Core.Models;

public abstract record ContentBlock
{
    public string Type { get; init; } = string.Empty;
}

public record ParagraphBlock : ContentBlock
{
    public ParagraphBlock()
    {
        Type = "paragraph";
    }

    public string Text { get; init; } = string.Empty;
    public List<InlineFormat>? InlineFormats { get; init; }
    public new string? Style { get; init; }
}

public record HeadingBlock : ContentBlock
{
    public HeadingBlock()
    {
        Type = "heading";
    }

    public int Level { get; init; }
    public string Text { get; init; } = string.Empty;
    public new string? Style { get; init; }
}

public record ListBlock : ContentBlock
{
    public ListBlock()
    {
        Type = "list";
    }

    public bool Ordered { get; init; }
    public List<string> Items { get; init; } = new();
    public new string? Style { get; init; }
}

public record TableBlock : ContentBlock
{
    public TableBlock()
    {
        Type = "table";
    }

    public List<TableRow> Rows { get; init; } = new();
}

public record TableRow
{
    public List<TableCell> Cells { get; init; } = new();
}

public record TableCell
{
    public string Text { get; init; } = string.Empty;
}

public record BlockquoteBlock : ContentBlock
{
    public BlockquoteBlock()
    {
        Type = "blockquote";
    }

    public string Text { get; init; } = string.Empty;
    public new string? Style { get; init; }
}

public record CodeBlock : ContentBlock
{
    public CodeBlock()
    {
        Type = "code";
    }

    public string Text { get; init; } = string.Empty;
    public string? Language { get; init; }
    public new string? Style { get; init; }
}

public record HorizontalRuleBlock : ContentBlock
{
    public HorizontalRuleBlock()
    {
        Type = "horizontalRule";
    }
}

public record CustomBlock : ContentBlock
{
    public CustomBlock()
    {
        Type = "custom";
    }

    public string CustomType { get; init; } = string.Empty;
    public string Text { get; init; } = string.Empty;
    public new string? Style { get; init; }
}

public record InlineFormat
{
    public string Type { get; init; } = string.Empty;
    public string Text { get; init; } = string.Empty;
}
