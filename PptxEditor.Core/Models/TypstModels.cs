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
    /// <summary>OS/2 sCapHeight (0 when the OS/2 table predates version 2).</summary>
    public short CapHeight { get; init; }
    public Dictionary<int, ushort> AdvanceWidths { get; init; } = new();
}

public sealed class TypstSlide
{
    private readonly List<string> _warnings = new();

    public int SlideIndex { get; init; }
    public SlideLayout Layout { get; init; } = new();
    public List<TypstElement> Elements { get; init; } = new();

    /// <summary>
    /// Human-readable warnings for slide content that could not be rendered faithfully —
    /// e.g. charts or other unsupported graphic frames replaced by a visible placeholder,
    /// or SmartArt approximated as positioned text. Empty when every element converted
    /// cleanly. Populated by <see cref="Converters.PptxToTypstConverter"/> during
    /// conversion; never mutated afterwards.
    /// </summary>
    public IReadOnlyList<string> Warnings => _warnings;

    /// <summary>Adds a conversion warning for this slide (no-op for null/blank text).</summary>
    public void AddWarning(string warning)
    {
        if (!string.IsNullOrWhiteSpace(warning))
        {
            _warnings.Add(warning);
        }
    }
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
    /// <summary>
    /// OOXML model identifier (e.g. a SmartArt <c>dsp:sp modelId</c> GUID) when the
    /// element originates from a model-driven part; null for regular slide shapes.
    /// Enables joining emitted shapes back to data-model nodes.
    /// </summary>
    public string? ModelId { get; init; }
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

public sealed class TypstTextRun
{
    public string Content { get; set; } = "";
    public TypstTextFormatting Formatting { get; set; } = new();
    public bool IsLineBreak { get; set; }
}

public sealed class TypstParagraph
{
    public string Content { get; set; } = "";
    public List<TypstTextRun> Runs { get; set; } = new();
    public TypstTextFormatting Formatting { get; set; } = new();
    public int Level { get; set; }
    public string? BulletChar { get; set; }
    public string? AutoNumberType { get; set; }
    public bool HasBullet { get; set; }
    /// <summary>
    /// Resolved bullet glyph color (#RRGGBB) from a:buClr, or the paragraph's first-run
    /// color when a:buClrTx (follow text) is specified. Null = inherit surrounding text color.
    /// </summary>
    public string? BulletColor { get; set; }
    public double? LineSpacing { get; set; }
    public double? SpaceBefore { get; set; }
    public double? SpaceAfter { get; set; }
    public double? MarginLeft { get; set; }
    public double? Indent { get; set; }
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
    public double TextBoxHeight { get; set; }
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
    public string? Caps { get; init; }
}

public sealed class TypstImageElement
{
    public string FileName { get; init; } = string.Empty;
    public string FullPath { get; init; } = string.Empty;
    public double Width { get; init; }
    public double Height { get; init; }
    public double CornerRadius { get; set; }
    /// <summary>Native pixel width read from image header, or null if unknown.</summary>
    public int? PixelWidth { get; init; }
    /// <summary>Native pixel height read from image header, or null if unknown.</summary>
    public int? PixelHeight { get; init; }
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
    public List<TypstParagraph> Paragraphs { get; init; } = new();
    public TypstTextFormatting Formatting { get; init; } = new();
    public string? BackgroundColor { get; init; }
    public TypstTableCellInsets? Insets { get; init; }
    public string? VerticalAlign { get; init; }
    public int RowSpan { get; init; } = 1;
    public int ColSpan { get; init; } = 1;
    public TableStylePart? StylePart { get; init; }
}

public sealed class TypstTableCellInsets
{
    public double? Left { get; init; }
    public double? Right { get; init; }
    public double? Top { get; init; }
    public double? Bottom { get; init; }
}

public sealed class TypstShapeElement
{
    public string ShapeType { get; init; } = "rect";
    public string FillColor { get; init; } = string.Empty;
    /// <summary>
    /// Linear gradient fill (from an OOXML <c>a:gradFill</c>), emitted as Typst
    /// <c>gradient.linear</c>. Mutually exclusive with <see cref="FillColor"/>:
    /// only set when no solid fill was found.
    /// </summary>
    public TypstGradientFill? FillGradient { get; init; }
    public string StrokeColor { get; init; } = string.Empty;
    public double StrokeWidth { get; init; }
    /// <summary>
    /// When true and no stroke color/width was resolved, emit an explicit
    /// <c>stroke: none</c> so the shape does not pick up Typst's default 1pt black
    /// stroke. Set by the SmartArt drawing extractor, where cached drawing shapes
    /// carry their full styling inline and a missing/empty <c>a:ln</c> means
    /// PowerPoint "no border" semantics. Not set on the regular slide-shape path:
    /// there a missing <c>a:ln</c> may still inherit a themed outline that the
    /// converter does not resolve, so the previous emission is preserved.
    /// </summary>
    public bool NoStroke { get; init; }
    public double CornerRadius { get; init; }
    public List<(double X, double Y)> Points { get; init; } = new();
    /// <summary>
    /// Multi-contour polygon (custGeom with several moveTo subpaths, e.g. a ring whose
    /// hole must stay transparent). When more than one subpath is present the shape is
    /// emitted as a Typst <c>#path(closed: true, fill-rule: "even-odd")</c> instead of a
    /// flat <c>#polygon</c>, which would fill the hole. Normalized 0..1 coordinates,
    /// same as <see cref="Points"/>.
    /// </summary>
    public List<List<(double X, double Y)>> Subpaths { get; init; } = new();
}

/// <summary>
/// Linear gradient fill: axis <paramref name="Angle"/> in degrees (OOXML <c>a:lin ang</c>
/// and Typst <c>gradient.linear</c> share the same clockwise-from-left→right convention)
/// plus the color stops.
/// </summary>
public sealed record TypstGradientFill(double Angle, IReadOnlyList<TypstGradientStop> Stops);

/// <summary>One gradient color stop: <c>#RRGGBB</c>/<c>#RRGGBBAA</c> hex color + offset in [0, 1].</summary>
public sealed record TypstGradientStop(string Color, double Offset);

public sealed class TableStyleDefinition
{
    public string StyleId { get; init; } = string.Empty;
    public Dictionary<string, TableStylePart> Parts { get; init; } = new(StringComparer.OrdinalIgnoreCase);
}

public sealed class TableStylePart
{
    // Existing
    public string? BackgroundColor { get; init; }
    public string? BorderTopColor { get; init; }
    public string? BorderBottomColor { get; init; }
    public string? BorderLeftColor { get; init; }
    public string? BorderRightColor { get; init; }

    // NEW: Text formatting
    public bool? TextBold { get; init; }
    public bool? TextItalic { get; init; }
    public string? TextColor { get; init; }
    public double? TextFontSize { get; init; } // in points

    // NEW: Border widths (in points)
    public double? BorderTopWidth { get; init; }
    public double? BorderBottomWidth { get; init; }
    public double? BorderLeftWidth { get; init; }
    public double? BorderRightWidth { get; init; }

    // NEW: Border presence
    public bool BorderTopNone { get; init; }
    public bool BorderBottomNone { get; init; }
    public bool BorderLeftNone { get; init; }
    public bool BorderRightNone { get; init; }
    public bool BorderInsideHNone { get; init; }
    public bool BorderInsideVNone { get; init; }

    public TableBorderState BorderTopState { get; init; } = TableBorderState.Inherit;
    public TableBorderState BorderBottomState { get; init; } = TableBorderState.Inherit;
    public TableBorderState BorderLeftState { get; init; } = TableBorderState.Inherit;
    public TableBorderState BorderRightState { get; init; } = TableBorderState.Inherit;
    public bool BorderTopExplicit { get; init; }
    public bool BorderBottomExplicit { get; init; }
    public bool BorderLeftExplicit { get; init; }
    public bool BorderRightExplicit { get; init; }
    public TableBorderState BorderInsideHState { get; init; } = TableBorderState.Inherit;
    public TableBorderState BorderInsideVState { get; init; } = TableBorderState.Inherit;
    public string? BorderInsideHColor { get; init; }
    public string? BorderInsideVColor { get; init; }
    public double? BorderInsideHWidth { get; init; }
    public double? BorderInsideVWidth { get; init; }
    public bool BackgroundCleared { get; init; }
}

public enum TableBorderState
{
    Inherit,
    Visible,
    None
}
