namespace DocxEditor.Core.Models;

public sealed record TypstDocument
{
    public TypstPageSetup PageSetup { get; init; } = new();
    public string DefaultFontFamily { get; init; } = "Liberation Serif";
    public double DefaultFontSizePt { get; init; } = 11;
    public List<TypstBlock> HeaderBlocks { get; init; } = [];
    public List<TypstBlock> FooterBlocks { get; init; } = [];
    public TypstHeaderFooterSet HeaderFooter { get; init; } = new();
    public List<TypstBlock> Blocks { get; init; } = [];
    public List<string> Diagnostics { get; init; } = [];
    public string TempDirectory { get; init; } = string.Empty;
    public string AssetsDirectory { get; init; } = string.Empty;
}

public sealed record TypstHeaderFooterSet
{
    public List<TypstBlock> FirstHeaderBlocks { get; init; } = [];
    public List<TypstBlock> FirstFooterBlocks { get; init; } = [];
    public List<TypstBlock> DefaultHeaderBlocks { get; init; } = [];
    public List<TypstBlock> DefaultFooterBlocks { get; init; } = [];
    public bool HasTitlePage { get; init; }
    public bool ApplyDefaultAfterFirstPageBreak { get; init; }
}

public sealed record TypstPageSetup
{
    public double WidthInches { get; init; } = 8.27;
    public double HeightInches { get; init; } = 11.69;
    public TypstMargins Margins { get; init; } = new();
}

public sealed record TypstMargins
{
    public double TopInches { get; init; } = 1;
    public double RightInches { get; init; } = 1;
    public double BottomInches { get; init; } = 1;
    public double LeftInches { get; init; } = 1;
}

public abstract record TypstBlock;

public sealed record TypstParagraphBlock : TypstBlock
{
    public List<TypstInline> Inlines { get; init; } = [];
    public string? Alignment { get; init; }
    public double? SpaceBeforePt { get; init; }
    public double? SpaceAfterPt { get; init; }
    public double? LeadingPt { get; init; }
}

public sealed record TypstPageBreakBlock : TypstBlock;

public sealed record TypstPageSettingsBlock : TypstBlock
{
    public TypstPageSetup PageSetup { get; init; } = new();
    public TypstHeaderFooterSet HeaderFooter { get; init; } = new();
    public bool PageBreakBefore { get; init; } = true;
}

public sealed record TypstListBlock : TypstBlock
{
    public bool Ordered { get; init; }
    public List<TypstListItem> Items { get; init; } = [];
}

public sealed record TypstListItem
{
    public List<TypstInline> Inlines { get; init; } = [];
}

public sealed record TypstTableBlock : TypstBlock
{
    public List<TypstTableRow> Rows { get; init; } = [];
    public bool HasBorders { get; init; }
}

public sealed record TypstTableRow
{
    public List<TypstTableCell> Cells { get; init; } = [];
}

public sealed record TypstTableCell
{
    public List<TypstParagraphBlock> Paragraphs { get; init; } = [];
    public string? ShadingColor { get; init; }
}

public sealed record TypstImageBlock : TypstBlock
{
    public string Path { get; init; } = string.Empty;
    public double? WidthInches { get; init; }
    public double? HeightInches { get; init; }
    public double? XPt { get; init; }
    public double? YPt { get; init; }
    public bool IsUnsupportedFormat { get; init; }
}

public sealed record TypstShapeBlock : TypstBlock
{
    public List<TypstParagraphBlock> Paragraphs { get; init; } = [];
    public double? XPt { get; init; }
    public double? YPt { get; init; }
    public double? WidthPt { get; init; }
    public double? HeightPt { get; init; }
    public string? FillColor { get; init; }
    public string? StrokeColor { get; init; }
}

public sealed record TypstInline
{
    public TypstInlineKind Kind { get; init; } = TypstInlineKind.Text;
    public string Text { get; init; } = string.Empty;
    public string RawTypst { get; init; } = string.Empty;
    public bool Bold { get; init; }
    public bool Italic { get; init; }
    public bool Underline { get; init; }
    public string? Color { get; init; }
    public double? FontSizePt { get; init; }
    public string? FontFamily { get; init; }
}

public enum TypstInlineKind
{
    Text,
    LineBreak,
    Tab,
    RawTypst
}
