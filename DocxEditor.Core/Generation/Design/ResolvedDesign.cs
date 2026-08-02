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
    /// <summary>The resolved theme's catalog name (defaults to <c>editorial</c>).</summary>
    public string ThemeName { get; init; } = DesignThemeCatalog.DefaultTheme;

    /// <summary>Palette as token name → #RRGGBB (theme palette merged with document overrides).</summary>
    public IReadOnlyDictionary<string, string> Palette { get; init; } = new Dictionary<string, string>();

    /// <summary>Body font family (theme default, overridden by the document). Null when unset.</summary>
    public string? BodyFontFamily { get; init; }

    /// <summary>Display/heading font family (theme default, overridden by the document). Null when unset.</summary>
    public string? DisplayFontFamily { get; init; }

    /// <summary>Resolved typography tokens, keyed by token name.</summary>
    public IReadOnlyDictionary<string, ResolvedRunFormatting> Typography { get; init; } =
        new Dictionary<string, ResolvedRunFormatting>();

    /// <summary>Named spacing values in points, keyed by token name.</summary>
    public IReadOnlyDictionary<string, double> Spacing { get; init; } = new Dictionary<string, double>();

    /// <summary>Resolved semantic-role defaults (theme roles, density-scaled), keyed by role.</summary>
    public IReadOnlyDictionary<TextRole, ResolvedRoleFormatting> Roles { get; init; } =
        new Dictionary<TextRole, ResolvedRoleFormatting>();

    /// <summary>Resolved layout guardrails and density scale.</summary>
    public ResolvedLayout Layout { get; init; } = new();

    /// <summary>Resolved shape defaults (fill/stroke/corner) in plain values.</summary>
    public ResolvedShapeDefaults Shapes { get; init; } = new();

    /// <summary>Resolved page defaults (size, orientation, margins, fonts).</summary>
    public ResolvedPageDefaults Page { get; init; } = new();

    /// <summary>Body font, falling back to the built-in default when the design is silent.</summary>
    public string BodyFontFamilyOrDefault => BodyFontFamily ?? DocxDesignDefaults.BodyFont;

    /// <summary>Display font, falling back to the built-in default when the design is silent.</summary>
    public string DisplayFontFamilyOrDefault => DisplayFontFamily ?? DocxDesignDefaults.DisplayFont;

    /// <summary>Resolved formatting for a role, falling back to the <see cref="TextRole.Body"/> default.</summary>
    public ResolvedRoleFormatting RoleOrDefault(TextRole role) =>
        Roles.TryGetValue(role, out var formatting) ? formatting : Roles.GetValueOrDefault(TextRole.Body) ?? ResolvedRoleFormatting.Empty;
}

/// <summary>
/// Fully-resolved formatting for one semantic <see cref="TextRole"/>: run formatting plus
/// paragraph formatting (already density-scaled) and keep-with-next/keep-lines flags. Emitters
/// use this when a paragraph's content carries a role (or infers one from a heading level).
/// </summary>
public sealed record ResolvedRoleFormatting
{
    /// <summary>Identity/default role formatting — nothing set.</summary>
    public static ResolvedRoleFormatting Empty { get; } = new();

    /// <summary>Resolved run formatting (font, size, color, emphasis, caps).</summary>
    public ResolvedRunFormatting Run { get; init; } = ResolvedRunFormatting.Empty;

    /// <summary>Resolved paragraph formatting (alignment, density-scaled spacing).</summary>
    public ResolvedParagraphFormatting Paragraph { get; init; } = ResolvedParagraphFormatting.Empty;

    /// <summary>Keep the paragraph with the next one.</summary>
    public bool KeepNext { get; init; }

    /// <summary>Keep all lines of the paragraph on the same page.</summary>
    public bool KeepLines { get; init; }

    /// <summary>True when nothing is set.</summary>
    public bool IsEmpty => Run.IsEmpty && Paragraph.IsEmpty && !KeepNext && !KeepLines;
}

/// <summary>Resolved layout guardrails and density knobs in plain values.</summary>
public sealed record ResolvedLayout
{
    /// <summary>Resolved density (theme default when the document is silent).</summary>
    public Density Density { get; init; } = Density.Comfortable;

    /// <summary>Deterministic spacing multiplier for the resolved density.</summary>
    public double DensityScale { get; init; } = 1.0;

    /// <summary>Minimum recommended body-text size in points (guardrail threshold).</summary>
    public double MinBodySizePt { get; init; } = 9;

    /// <summary>Maximum recommended table width in points; null = the page text width.</summary>
    public double? MaxTableWidthPt { get; init; }
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
