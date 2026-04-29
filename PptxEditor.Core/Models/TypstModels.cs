using System;

namespace PptxEditor.Core.Models;

public sealed class TypstPresentation
{
    public List<TypstSlide> Slides { get; init; } = new();
    public string TempDirectory { get; init; } = string.Empty;
    public List<string> FontFiles { get; init; } = new();
    public Dictionary<string, TypstFontMetrics> FontMetrics { get; init; } = new(StringComparer.OrdinalIgnoreCase);
    public Dictionary<string, string> ThemeFonts { get; init; } = new(StringComparer.OrdinalIgnoreCase);
}

public sealed class TypstFontMetrics
{
    public int UnitsPerEm { get; init; }
    public short TypoAscender { get; init; }
    public short TypoDescender { get; init; }
    public short TypoLineGap { get; init; }
    public short HheaAscender { get; init; }
    public short HheaDescender { get; init; }
    public short HheaLineGap { get; init; }
    public ushort WinAscent { get; init; }
    public ushort WinDescent { get; init; }
    public Dictionary<int, ushort> AdvanceWidths { get; init; } = new();
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
    public double Rotation { get; init; }
    public TypstTextElement? Text { get; init; }
    public TypstImageElement? Image { get; init; }
    public TypstTableElement? Table { get; init; }
    public TypstShapeElement? Shape { get; init; }
}

public sealed class TypstParagraph
{
    public string Content { get; set; } = "";
    public TypstTextFormatting Formatting { get; set; } = new();
    public int Level { get; set; }
    public string? BulletChar { get; set; }
    public string? AutoNumberType { get; set; }
    public bool HasBullet { get; set; }
    public double? LineSpacing { get; set; }
}

public sealed class TypstTextElement
{
    public List<TypstParagraph> Paragraphs { get; set; } = new();
    public string Content => string.Join("\n\n", Paragraphs.Select(p => p.Content));
    public TypstTextFormatting Formatting => Paragraphs.FirstOrDefault()?.Formatting ?? new();
    public bool AutoFit { get; init; }
    public double PaddingLeft { get; init; }
    public double PaddingTop { get; init; }
    public double PaddingRight { get; init; }
    public double PaddingBottom { get; init; }
    public double? LineSpacing { get; init; }
    public string VerticalAlign { get; init; } = "top";
    public int ParagraphCount { get; init; } = 1;
    public bool HasExplicitLineBreaks { get; init; }
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

public sealed class TypstShapeElement
{
    public string ShapeType { get; init; } = "rect";
    public string FillColor { get; init; } = string.Empty;
    public string StrokeColor { get; init; } = string.Empty;
    public double StrokeWidth { get; init; }
    public double CornerRadius { get; init; }
    public List<(double X, double Y)> Points { get; init; } = new();
}

public sealed class TableStyleDefinition
{
    public string StyleId { get; init; } = string.Empty;
    public Dictionary<string, TableStylePart> Parts { get; init; } = new(StringComparer.OrdinalIgnoreCase);
}

public sealed class TableStylePart
{
    public string? BackgroundColor { get; init; }
    public string? BorderTopColor { get; init; }
    public string? BorderBottomColor { get; init; }
    public string? BorderLeftColor { get; init; }
    public string? BorderRightColor { get; init; }
}
