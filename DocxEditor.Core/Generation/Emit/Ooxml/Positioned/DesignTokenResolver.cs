using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Positioned;

/// <summary>
/// Resolves generation-vocabulary references against the (optional) design tokens at emit
/// time: palette token names → #RRGGBB hex, font slot names ("display"/"body") → family
/// names, and typography token lookup for text content. The parser keeps palette token
/// names in the model (not resolved values), so every emitter resolves them here. Raw
/// #RRGGBB literals pass through unchanged. Unknown tokens fall back to the supplied
/// default rather than throwing, since the parser has already validated references — this
/// class is defensive against callers that build the model by hand.
/// </summary>
internal sealed class DesignTokenResolver
{
    private readonly DesignTokens? _design;

    public DesignTokenResolver(DesignTokens? design)
    {
        _design = design;
    }

    /// <summary>
    /// Resolves a color reference (palette token or #RRGGBB) to its RRGGBB hex without the
    /// leading '#'. Unknown references fall back to <paramref name="fallbackHex"/>.
    /// </summary>
    public string ResolveHex(string? color, string fallbackHex)
    {
        if (string.IsNullOrWhiteSpace(color))
        {
            return fallbackHex;
        }

        if (color.StartsWith('#'))
        {
            return color.Length == 7 ? color[1..].ToUpperInvariant() : fallbackHex;
        }

        if (_design is not null && _design.Palette.TryGetValue(color, out string? hex) && !string.IsNullOrWhiteSpace(hex))
        {
            return hex.StartsWith('#') ? hex[1..].ToUpperInvariant() : hex.ToUpperInvariant();
        }

        return fallbackHex;
    }

    /// <summary>
    /// Resolves a font reference: the slot names "display"/"body" map to the design font
    /// slots, anything else is treated as a raw family name. Returns null when the slot is
    /// undefined (caller falls back to a document default).
    /// </summary>
    public string? ResolveFontFamily(string? slotOrFamily)
    {
        if (string.IsNullOrWhiteSpace(slotOrFamily))
        {
            return null;
        }

        if (slotOrFamily.Equals("display", StringComparison.OrdinalIgnoreCase))
        {
            return _design?.Fonts.Display;
        }

        if (slotOrFamily.Equals("body", StringComparison.OrdinalIgnoreCase))
        {
            return _design?.Fonts.Body;
        }

        return slotOrFamily;
    }

    /// <summary>Looks up a named typography token, or null when unknown/undefined.</summary>
    public TypographyToken? TryGetTypographyToken(string? name)
    {
        return name is not null && _design is not null && _design.Typography.TryGetValue(name, out TypographyToken? token)
            ? token
            : null;
    }

    /// <summary>
    /// Default font family for body text: the design page default font, then the design
    /// "body" slot, then the built-in "Calibri".
    /// </summary>
    public string DefaultBodyFontFamily =>
        ResolveFontFamily(_design?.Page.DefaultFontFamily)
        ?? _design?.Fonts.Body
        ?? "Calibri";

    /// <summary>Default text color (RRGGBB) for body text: design default or black.</summary>
    public string DefaultTextHex => ResolveHex(_design?.Page.DefaultTextColor, "000000");

    /// <summary>Default shape corner radius: the design shape default (0 = square).</summary>
    public double DefaultCornerRadiusPt => _design?.Shapes.CornerRadiusPt ?? 0;

    /// <summary>
    /// Resolves the effective fill hex for a box primitive, or null for no fill. Explicit
    /// values win; otherwise the design <c>shapes.defaultFill</c> is used.
    /// </summary>
    public string? ResolveBoxFillHex(string? fill)
    {
        if (fill is not null)
        {
            return ResolveHex(fill, "000000");
        }

        return _design?.Shapes.DefaultFill is { } defaultFill ? ResolveHex(defaultFill, "000000") : null;
    }

    /// <summary>
    /// Resolves the effective stroke for a box primitive, or null for no stroke. Explicit
    /// strokes win; otherwise the design <c>shapes.defaultStroke</c> is used.
    /// </summary>
    public StrokeSpec? ResolveBoxStroke(StrokeSpec? stroke)
    {
        if (stroke is not null)
        {
            return stroke;
        }

        if (_design?.Shapes.DefaultStrokeColor is { } defaultColor)
        {
            return new StrokeSpec { Color = defaultColor, WidthPt = _design.Shapes.DefaultStrokeWidthPt };
        }

        return null;
    }

    /// <summary>
    /// Resolves the effective stroke for a line. A line with no stroke at all is
    /// invisible, so unlike box primitives the fallback is a visible default (black 1pt)
    /// rather than none.
    /// </summary>
    public StrokeSpec ResolveLineStroke(StrokeSpec? stroke) =>
        ResolveBoxStroke(stroke)
        ?? new StrokeSpec { Color = "000000", WidthPt = 1 };
}
