using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Design;

/// <summary>
/// A named, built-in design theme: the base layer of the design system. A theme supplies
/// display/body font slots, a palette (token → #RRGGBB), a full set of <see cref="TextRole"/>
/// formatting defaults, a default density, a body-size readability floor and default page
/// geometry. The document's <c>design</c> block overrides theme values field-by-field, and
/// content-level token/direct formatting overrides the resolved roles — so a theme is always
/// the fallback, never the master. See <see cref="DesignThemeCatalog"/> for the shipped themes
/// (<c>editorial</c> is the default).
/// </summary>
public sealed record DesignTheme
{
    /// <summary>Catalog name referenced by <c>design.theme</c>.</summary>
    public required string Name { get; init; }

    /// <summary>Display/heading typeface family.</summary>
    public required string DisplayFontFamily { get; init; }

    /// <summary>Body typeface family.</summary>
    public required string BodyFontFamily { get; init; }

    /// <summary>Theme palette, token name → #RRGGBB.</summary>
    public required IReadOnlyDictionary<string, string> Palette { get; init; }

    /// <summary>Default formatting for every <see cref="TextRole"/> (must cover all roles).</summary>
    public required IReadOnlyDictionary<TextRole, ThemeRoleFormatting> Roles { get; init; }

    /// <summary>Default typographic density (spacing scale).</summary>
    public Density Density { get; init; } = Density.Comfortable;

    /// <summary>Readability floor for body-size text in points (guardrail threshold).</summary>
    public double MinBodySizePt { get; init; } = 9;

    /// <summary>Default page size when a section and <c>design.page</c> are both silent.</summary>
    public PageSizeName PageSize { get; init; } = PageSizeName.A4;

    /// <summary>Default margins when a section and <c>design.page</c> are both silent.</summary>
    public Margins Margins { get; init; } = Margins.Defaults;

    /// <summary>Returns the theme's palette merged with document-level overrides (document wins).</summary>
    public IReadOnlyDictionary<string, string> EffectivePalette(IReadOnlyDictionary<string, string>? overrides)
    {
        if (overrides is null || overrides.Count == 0)
        {
            return Palette;
        }
        var merged = new Dictionary<string, string>(Palette, StringComparer.Ordinal);
        foreach (var (token, hex) in overrides)
        {
            merged[token] = hex;
        }
        return merged;
    }
}

/// <summary>
/// One theme's default formatting for a <see cref="TextRole"/>: run properties (font slot or
/// raw family, size, palette color, emphasis, all-caps) plus paragraph properties (alignment,
/// before/after spacing, line multiple, keep-with-next). <see cref="FontFamily"/> and
/// <see cref="Color"/> are resolved against the theme at resolution time; spacing values are
/// density-scaled before emitters see them.
/// </summary>
public sealed record ThemeRoleFormatting
{
    /// <summary>Font slot ("display" | "body") or a raw family name.</summary>
    public string? FontFamily { get; init; }

    /// <summary>Font size in points (&gt; 0).</summary>
    public double? SizePt { get; init; }

    /// <summary>Palette token name resolved to #RRGGBB.</summary>
    public string? Color { get; init; }

    /// <summary>Bold weight.</summary>
    public bool Bold { get; init; }

    /// <summary>Italic style.</summary>
    public bool Italic { get; init; }

    /// <summary>Underline.</summary>
    public bool Underline { get; init; }

    /// <summary>All-caps rendering (<c>w:caps</c>).</summary>
    public bool AllCaps { get; init; }

    /// <summary>Default paragraph alignment.</summary>
    public TextAlignment? Alignment { get; init; }

    /// <summary>Space before in points (density-scaled at resolution).</summary>
    public double? BeforePt { get; init; }

    /// <summary>Space after in points (density-scaled at resolution).</summary>
    public double? AfterPt { get; init; }

    /// <summary>Line-spacing multiple.</summary>
    public double? LineMultiple { get; init; }

    /// <summary>Keep the paragraph with the next one (headings).</summary>
    public bool KeepNext { get; init; }

    /// <summary>Keep all lines of the paragraph together.</summary>
    public bool KeepLines { get; init; }
}
