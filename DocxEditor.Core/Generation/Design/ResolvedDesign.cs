using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Design;

/// <summary>
/// The fully-resolved design state: foundation defaults merged with the document's <c>design</c>
/// tokens. Colors are normalized #RRGGBB literals, fonts are resolved families, typography tokens
/// are resolved <see cref="ResolvedRunFormatting"/> values, and shape/page defaults are plain
/// values — emitters consume this record (via <see cref="DocxDesignResolver.ResolveAll"/>) without
/// interpreting token strings.
/// </summary>
public sealed record ResolvedDesign
{
    /// <summary>Palette as token name → #RRGGBB (never empty in practice; empty when no palette).</summary>
    public IReadOnlyDictionary<string, string> Palette { get; init; } = new Dictionary<string, string>();

    /// <summary>Body font family. Null when the design does not specify one (inherit document default).</summary>
    public string? BodyFontFamily { get; init; }

    /// <summary>Display/heading font family. Null when the design does not specify one (inherit document default).</summary>
    public string? DisplayFontFamily { get; init; }

    /// <summary>Resolved typography tokens, keyed by token name.</summary>
    public IReadOnlyDictionary<string, ResolvedRunFormatting> Typography { get; init; } =
        new Dictionary<string, ResolvedRunFormatting>();

    /// <summary>Named spacing values in points, keyed by token name.</summary>
    public IReadOnlyDictionary<string, double> Spacing { get; init; } = new Dictionary<string, double>();

    /// <summary>Resolved shape defaults (fill/stroke/corner) in plain values.</summary>
    public ResolvedShapeDefaults Shapes { get; init; } = new();

    /// <summary>Resolved page defaults (size, orientation, margins, fonts).</summary>
    public ResolvedPageDefaults Page { get; init; } = new();

    /// <summary>Body font, falling back to the built-in default when the design is silent.</summary>
    public string BodyFontFamilyOrDefault => BodyFontFamily ?? DocxDesignDefaults.BodyFont;

    /// <summary>Display font, falling back to the built-in default when the design is silent.</summary>
    public string DisplayFontFamilyOrDefault => DisplayFontFamily ?? DocxDesignDefaults.DisplayFont;
}

/// <summary>Resolved shape defaults with token references already converted to #RRGGBB.</summary>
public sealed record ResolvedShapeDefaults
{
    /// <summary>Default corner radius in points (0 = square).</summary>
    public double CornerRadiusPt { get; init; }

    /// <summary>Default fill as #RRGGBB. Null = no default fill.</summary>
    public string? DefaultFillHex { get; init; }

    /// <summary>Default stroke color as #RRGGBB. Null = no default stroke.</summary>
    public string? DefaultStrokeColorHex { get; init; }

    /// <summary>Default stroke width in points.</summary>
    public double DefaultStrokeWidthPt { get; init; } = 1;
}

/// <summary>Resolved page defaults with token references already converted to concrete values.</summary>
public sealed record ResolvedPageDefaults
{
    /// <summary>Default page size (A4 when unspecified).</summary>
    public PageSize PageSize { get; init; } = PageSize.Default;

    /// <summary>Default orientation (portrait).</summary>
    public PageOrientation Orientation { get; init; } = PageOrientation.Portrait;

    /// <summary>Default margins (1 inch on every edge when unspecified).</summary>
    public Margins Margins { get; init; } = Margins.Defaults;

    /// <summary>Default body font family (slot-resolved). Null = inherit document default.</summary>
    public string? BodyFontFamily { get; init; }

    /// <summary>Default body text color as #RRGGBB. Null = inherit.</summary>
    public string? DefaultTextColorHex { get; init; }

    /// <summary>Effective (orientation-aware) page dimensions for the default page size.</summary>
    public (double WidthPt, double HeightPt) EffectiveSize() => PageSize.EffectiveSize(Orientation);
}
