using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxEditor.Core.Exceptions;
using XlsxEditor.Core.Styles;

namespace XlsxEditor.Core.Builders;

/// <summary>
/// Describes the font aspect of a named style. <see cref="ColorArgb"/> accepts a
/// 6-digit RGB ("FF0000"), 8-digit ARGB ("FFFF0000") hex string or a common color name
/// ("red"), normalized to 8-digit ARGB before it is written to the stylesheet.
/// <see cref="Size"/> is in points and bounded by Excel's font-size range.
/// </summary>
public sealed record CellFontSpec
{
    /// <summary>Applies the &lt;b/&gt; (bold) font element.</summary>
    public bool Bold { get; init; }

    /// <summary>Applies the &lt;i/&gt; (italic) font element.</summary>
    public bool Italic { get; init; }

    /// <summary>Font color as a 6-digit RGB, 8-digit ARGB hex string or common color name.</summary>
    public string? ColorArgb { get; init; }

    /// <summary>Font size in points, between 1 and 409.5 (Excel's font-size limit).</summary>
    public double? Size { get; init; }

    internal bool IsDefault => !Bold && !Italic && ColorArgb is null && Size is null;

    internal void Validate(string context)
    {
        if (ColorArgb is not null)
        {
            CellStyleSpec.ValidateColorArgb(ColorArgb, context);
        }

        if (Size is { } size)
        {
            if (!double.IsFinite(size) || size <= 0)
            {
                throw new XlsxException(
                    $"{context} font size must be a finite number greater than zero (points). " +
                    $"Got {size.ToString(CultureInfo.InvariantCulture)}.");
            }

            if (size > 409.5)
            {
                throw new XlsxException(
                    $"{context} font size must not exceed 409.5 points (Excel's maximum). " +
                    $"Got {size.ToString(CultureInfo.InvariantCulture)}.");
            }
        }
    }
}

/// <summary>
/// Describes the fill aspect of a named style. <see cref="ColorArgb"/> accepts a 6-digit
/// RGB, 8-digit ARGB hex string or a common color name (normalized to 8-digit ARGB before
/// it is written to the stylesheet). <see cref="Pattern"/> selects the Excel pattern type
/// ("solid", "gray125", …); when null and a color is present the fill is solid, matching
/// the pre-pattern contract. Only pattern fills are supported (gradient fills remain out
/// of scope for the style builder).
/// </summary>
public sealed record CellFillSpec
{
    /// <summary>Foreground fill color as a 6-digit RGB, 8-digit ARGB hex string or color name.</summary>
    public string? SolidColorArgb { get; init; }

    /// <summary>Excel pattern type (solid, gray125, …); null means solid when a color is set.</summary>
    public PatternValues? Pattern { get; init; }

    internal bool IsDefault => Pattern is null && SolidColorArgb is null;

    internal void Validate(string context)
    {
        if (SolidColorArgb is not null)
        {
            CellStyleSpec.ValidateColorArgb(SolidColorArgb, context);
        }
    }
}

/// <summary>
/// Describes one edge (left/right/top/bottom) of a cell border: an Excel border style
/// plus an optional color. <see cref="Style"/> uses the SpreadsheetML
/// <see cref="BorderStyleValues"/> enum ("Thin", "Medium", "Dashed", "Dotted", …); a null
/// style means the edge carries no border. <see cref="ColorArgb"/> follows the shared
/// color contract (hex or a common color name).
/// </summary>
public sealed record CellEdgeSpec
{
    /// <summary>Border style (Thin, Medium, Dashed, Dotted, …); null means no border.</summary>
    public BorderStyleValues? Style { get; init; }

    /// <summary>Border color as a 6-digit RGB, 8-digit ARGB hex string or color name.</summary>
    public string? ColorArgb { get; init; }

    internal bool IsDefault => Style is null && ColorArgb is null;

    internal void Validate(string context)
    {
        if (ColorArgb is not null)
        {
            CellStyleSpec.ValidateColorArgb(ColorArgb, context);
        }
    }
}

/// <summary>
/// Describes the border aspect of a named style: up to four edges, each with a style and
/// optional color. An empty border (no edges with a style) is treated as "no border".
/// </summary>
public sealed record CellBorderSpec
{
    public CellEdgeSpec? Left { get; init; }

    public CellEdgeSpec? Right { get; init; }

    public CellEdgeSpec? Top { get; init; }

    public CellEdgeSpec? Bottom { get; init; }

    internal bool IsDefault =>
        (Left is null || Left.IsDefault)
        && (Right is null || Right.IsDefault)
        && (Top is null || Top.IsDefault)
        && (Bottom is null || Bottom.IsDefault);

    internal void Validate(string context)
    {
        Left?.Validate(context);
        Right?.Validate(context);
        Top?.Validate(context);
        Bottom?.Validate(context);
    }
}

/// <summary>
/// Describes the alignment aspect of a named style: horizontal and vertical placement
/// plus optional wrap-text. Alignment values come from the SpreadsheetML enum of the
/// same name used in the &lt;alignment/&gt; element.
/// </summary>
public sealed record CellAlignmentSpec
{
    /// <summary>Horizontal placement (e.g. Center, Left, Right); null leaves it unset.</summary>
    public HorizontalAlignmentValues? Horizontal { get; init; }

    /// <summary>Vertical placement (e.g. Center, Top, Bottom); null leaves it unset.</summary>
    public VerticalAlignmentValues? Vertical { get; init; }

    /// <summary>Wraps text within the cell when true (wrapText="1").</summary>
    public bool WrapText { get; init; }

    internal bool IsDefault => Horizontal is null && Vertical is null && !WrapText;

    internal void Validate(string context)
    {
        // Horizontal/Vertical are strongly-typed SpreadsheetML enums, so an invalid
        // value cannot be produced through the public API without a deliberate cast;
        // the SDK's serialization rejects such values at save time.
    }
}

/// <summary>
/// A named cell style: an immutable description of font, fill, border, alignment and
/// number format. Styles are defined workbook-wide via
/// <see cref="IWorkbookBuilder.DefineStyle"/> and referenced from cell writes by
/// <paramref name="Name"/>. Names are a builder-session concept: after a save/reopen,
/// a style must be re-defined under its name before it can be referenced again — the
/// structural dedup makes that cheap, because re-defining an identical spec reuses the
/// existing cellXf instead of appending a duplicate. The style builder is append-only
/// and deduplicated: every style resolves to a cellXf index in the workbook stylesheet,
/// structurally identical styles share one index, and existing style definitions are
/// never mutated when a workbook is opened.
/// </summary>
public sealed record CellStyleSpec
{
    /// <summary>
    /// The style's unique, case-insensitive name, used to reference it from cell writes.
    /// Must be a non-empty, non-whitespace string.
    /// </summary>
    public string Name { get; init; } = string.Empty;

    /// <summary>Optional font properties (bold, italic, color, size).</summary>
    public CellFontSpec? Font { get; init; }

    /// <summary>Optional fill properties (pattern type + foreground color).</summary>
    public CellFillSpec? Fill { get; init; }

    /// <summary>Optional border properties (per-edge style + color).</summary>
    public CellBorderSpec? Border { get; init; }

    /// <summary>Optional alignment properties (horizontal, vertical, wrap).</summary>
    public CellAlignmentSpec? Alignment { get; init; }

    /// <summary>
    /// Optional Excel number-format code (e.g. "0.00", "$#,##0.00", "yyyy-mm-dd").
    /// Built-in codes map to their built-in numFmt id; any other code is registered as
    /// a custom numFmt (id &gt;= 164) on first use.
    /// </summary>
    public string? NumberFormat { get; init; }

    /// <summary>
    /// Validates the style's name and every component, throwing an
    /// <see cref="XlsxException"/> describing the first invalid aspect. It must run
    /// before any stylesheet mutation so a rejected style never leaves a partial
    /// entry behind.
    /// </summary>
    internal void Validate()
    {
        if (string.IsNullOrWhiteSpace(Name))
        {
            throw new XlsxException(
                "A named style must have a non-empty name; null, empty and whitespace " +
                "names are not allowed.");
        }

        Font?.Validate($"Style '{Name}'");
        Fill?.Validate($"Style '{Name}'");
        Border?.Validate($"Style '{Name}'");
        Alignment?.Validate($"Style '{Name}'");

        if (NumberFormat is not null && string.IsNullOrWhiteSpace(NumberFormat))
        {
            throw new XlsxException(
                $"Style '{Name}' has a whitespace-only number format; use null to " +
                "leave the number format unset.");
        }
    }

    /// <summary>
    /// Validates a color against the shared style color contract: a 6-digit RGB, an
    /// 8-digit ARGB hex string, or a common color name. Throws an <see cref="XlsxException"/>
    /// describing the first invalid aspect.
    /// </summary>
    internal static void ValidateColorArgb(string color, string context)
    {
        if (!ExcelColor.IsValid(color))
        {
            throw new XlsxException(
                $"{context} color '{color}' must be a 6-digit RGB (e.g. 'FF0000'), an " +
                "8-digit ARGB (e.g. 'FFFF0000') hex string, or a common color name " +
                "(e.g. 'red', 'white', 'darkgray').");
        }
    }

    /// <summary>
    /// Normalizes a 6-digit RGB to 8-digit ARGB (opaque alpha), resolves a common color
    /// name to its hex value, and upper-cases an 8-digit ARGB. The input must already be
    /// validated by <see cref="ValidateColorArgb"/>.
    /// </summary>
    internal static string NormalizeColorArgb(string color) => ExcelColor.ToArgb(color, "Style");
}
