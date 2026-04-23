namespace PptxEditor.Core.Models;

public sealed class TypstPresentation
{
    public List<TypstSlide> Slides { get; init; } = new();
    public string TempDirectory { get; init; } = string.Empty;
    public List<string> FontFiles { get; init; } = new();
}

public sealed class TypstSlide
{
    public int SlideIndex { get; init; }
    public SlideLayout Layout { get; init; } = new();
    public List<TypstElement> Elements { get; init; } = new();
}

public sealed class SlideLayout
{
    public double Width { get; init; }
    public double Height { get; init; }
    public string? BackgroundColor { get; init; }
}

public sealed class TypstElement
{
    public string Type { get; init; } = string.Empty;
    public uint Id { get; init; }
    public string Name { get; init; } = string.Empty;
    public double X { get; init; }
    public double Y { get; init; }
    public double Width { get; init; }
    public double Height { get; init; }
    public TypstTextElement? Text { get; init; }
    public TypstImageElement? Image { get; init; }
    public TypstTableElement? Table { get; init; }
}

public sealed class TypstTextElement
{
    public string Content { get; init; } = string.Empty;
    public TypstTextFormatting Formatting { get; init; } = new();
}

public sealed record TypstTextFormatting
{
    public double FontSize { get; init; } = 18.0;
    public bool Bold { get; init; }
    public bool Italic { get; init; }
    public bool Underline { get; init; }
    public string Color { get; init; } = "#000000";
    public string FontFamily { get; init; } = "Arial";
    public string Align { get; init; } = "left";
}

public sealed class TypstImageElement
{
    public string FileName { get; init; } = string.Empty;
    public string FullPath { get; init; } = string.Empty;
    public double Width { get; init; }
    public double Height { get; init; }
}

public sealed class TypstTableElement
{
    public List<List<TypstTableCell>> Rows { get; init; } = new();
    public List<double> ColumnWidths { get; init; } = new();
    public List<double> RowHeights { get; init; } = new();
    public string BorderColor { get; init; } = "#000000";
    public double BorderWidth { get; init; } = 1.0;
}

public sealed class TypstTableCell
{
    public string Content { get; init; } = string.Empty;
    public TypstTextFormatting Formatting { get; init; } = new();
    public string? BackgroundColor { get; init; }
    public int RowSpan { get; init; } = 1;
    public int ColSpan { get; init; } = 1;
}
