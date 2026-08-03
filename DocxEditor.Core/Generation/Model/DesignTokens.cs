namespace DocxEditor.Core.Generation.Model;

/// <summary>
/// Design tokens (vocabulary §3.1): the palette, font slots, typography scale, spacing
/// scale, shape defaults and page defaults that content references. Colors resolve
/// against <see cref="Palette"/> by name; raw #RRGGBB hex is accepted but warned
/// (off-token drift). All records are immutable.
/// </summary>
public sealed record DesignTokens
{
    /// <summary>
    /// Named built-in theme reference (see <c>DesignThemeCatalog</c>). Null = the default
    /// <c>editorial</c> theme. Theme values are the base layer that every field in this block
    /// overrides, so an empty <c>design</c> still yields a coherent editorial look.
    /// </summary>
    public string? Theme { get; init; }

    /// <summary>Named colors, token name → #RRGGBB. May be empty (theme palette then applies).</summary>
    public IReadOnlyDictionary<string, string> Palette { get; init; } = new Dictionary<string, string>();

    /// <summary>Font slot tokens (display / body). Optional.</summary>
    public FontTokens Fonts { get; init; } = new();

    /// <summary>Named typography tokens referenced by <c>token</c> on text content.</summary>
    public IReadOnlyDictionary<string, TypographyToken> Typography { get; init; } = new Dictionary<string, TypographyToken>();

    /// <summary>Named spacing values in points referenced by paragraph <c>spacing.before/after</c>.</summary>
    public IReadOnlyDictionary<string, double> Spacing { get; init; } = new Dictionary<string, double>();

    /// <summary>Default shape treatment (corner radius, fills, stroke).</summary>
    public ShapeDefaults Shapes { get; init; } = new();

    /// <summary>Default page geometry and typography (size, orientation, margins, fonts).</summary>
    public PageDefaults Page { get; init; } = new();

    /// <summary>Layout guardrails and density (spacing scale, minimum body size, table width).</summary>
    public LayoutDefaults Layout { get; init; } = new();
}

/// <summary>
/// Layout guardrails and density knobs. <see cref="Density"/> scales the theme's paragraph
/// spacing; <see cref="MinBodySizePt"/> and <see cref="MaxTableWidthPt"/> are advisory
/// thresholds that surface emitter warnings when violated (no pagination prediction).
/// </summary>
public sealed record LayoutDefaults
{
    /// <summary>Typographic density; scales theme spacing. Null = theme default (comfortable).</summary>
    public Density? Density { get; init; }

    /// <summary>Minimum recommended body-text size in points (≥ 0). Undersized body text warns.</summary>
    public double? MinBodySizePt { get; init; }

    /// <summary>Maximum recommended table width in points (> 0). Wider tables warn.</summary>
    public double? MaxTableWidthPt { get; init; }
}

/// <summary>Font slot tokens referenced by <c>font</c> fields ("display" | "body").</summary>
public sealed record FontTokens
{
    /// <summary>Display/heading typeface family name.</summary>
    public string? Display { get; init; }

    /// <summary>Body typeface family name.</summary>
    public string? Body { get; init; }
}

/// <summary>One named typography token bundling font, size and emphasis for text content.</summary>
public sealed record TypographyToken
{
    /// <summary>Font slot ("display" | "body") or a raw family name.</summary>
    public string? FontFamily { get; init; }

    /// <summary>Font size in points (&gt; 0).</summary>
    public double? SizePt { get; init; }

    /// <summary>Palette token name or #RRGGBB literal.</summary>
    public string? Color { get; init; }

    /// <summary>Bold weight.</summary>
    public bool Bold { get; init; }

    /// <summary>Italic style.</summary>
    public bool Italic { get; init; }

    /// <summary>Underline.</summary>
    public bool Underline { get; init; }

    /// <summary>All-caps rendering (<c>w:caps</c>), used by small display roles like eyebrow.</summary>
    public bool AllCaps { get; init; }
}

/// <summary>Default shape treatment applied by positioned box primitives.</summary>
public sealed record ShapeDefaults
{
    /// <summary>Default corner radius in points (0 = square).</summary>
    public double CornerRadiusPt { get; init; }

    /// <summary>Default fill (palette token or hex).</summary>
    public string? DefaultFill { get; init; }

    /// <summary>Default stroke color (palette token or hex).</summary>
    public string? DefaultStrokeColor { get; init; }

    /// <summary>Default stroke width in points.</summary>
    public double DefaultStrokeWidthPt { get; init; } = 1;
}

/// <summary>Default page geometry and typography used when a section omits its own setup.</summary>
public sealed record PageDefaults
{
    /// <summary>Default named page size (null = A4).</summary>
    public PageSizeName? PageSize { get; init; }

    /// <summary>Default orientation (portrait).</summary>
    public PageOrientation Orientation { get; init; } = PageOrientation.Portrait;

    /// <summary>Default margins (null = 1 inch on every edge).</summary>
    public Margins? Margins { get; init; }

    /// <summary>Default font slot ("display" | "body") or family name for body text.</summary>
    public string? DefaultFontFamily { get; init; }

    /// <summary>Default text color (palette token or hex).</summary>
    public string? DefaultTextColor { get; init; }
}
