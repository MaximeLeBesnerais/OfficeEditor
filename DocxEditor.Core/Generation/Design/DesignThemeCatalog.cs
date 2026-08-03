using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Design;

/// <summary>
/// The built-in named theme catalog. Every document resolves against a theme — the default is
/// <c>editorial</c>, a Word-native editorial-report treatment (Georgia display / Arial body,
/// navy/blue/teal/coral accents on pale surfaces, 1-inch margins) — regardless of whether the
/// JSON declares one. The document's <c>design</c> block overrides theme values field-by-field,
/// so the theme is the fallback layer, never the master.
/// </summary>
public static class DesignThemeCatalog
{
    /// <summary>The default theme name applied when <c>design.theme</c> is omitted.</summary>
    public const string DefaultTheme = "editorial";

    /// <summary>The default editorial-report theme (Georgia / Arial, navy-blue-teal-coral).</summary>
    public static DesignTheme Editorial { get; } = new()
    {
        Name = "editorial",
        DisplayFontFamily = "Georgia",
        BodyFontFamily = "Arial",
        Palette = new Dictionary<string, string>
        {
            ["ink"] = "#1C2733",
            ["paper"] = "#FFFFFF",
            ["white"] = "#FFFFFF",
            ["pale"] = "#EEF2F6",
            ["primary"] = "#1F3A5F",
            ["accent"] = "#2E6FB7",
            ["teal"] = "#1F7A6E",
            ["coral"] = "#E4674A",
            ["muted"] = "#5A6B7B",
            ["border"] = "#D6DEE6"
        },
        Roles = EditorialRoles(),
        Density = Density.Comfortable,
        MinBodySizePt = 9,
        PageSize = PageSizeName.A4,
        Margins = Margins.Defaults
    };

    /// <summary>A clean corporate-blue theme (optional alternative to <c>editorial</c>).</summary>
    public static DesignTheme Corporate { get; } = new()
    {
        Name = "corporate",
        DisplayFontFamily = "Trebuchet MS",
        BodyFontFamily = "Arial",
        Palette = new Dictionary<string, string>
        {
            ["ink"] = "#22262C",
            ["paper"] = "#FFFFFF",
            ["white"] = "#FFFFFF",
            ["pale"] = "#F2F6FB",
            ["primary"] = "#14498C",
            ["accent"] = "#3B82C4",
            ["teal"] = "#1E7B85",
            ["coral"] = "#B4552F",
            ["muted"] = "#5B6573",
            ["border"] = "#DCE4EE"
        },
        Roles = CorporateRoles(),
        Density = Density.Comfortable,
        MinBodySizePt = 9,
        PageSize = PageSizeName.Letter,
        Margins = Margins.Defaults
    };

    /// <summary>Theme lookup keyed case-insensitively by catalog name.</summary>
    public static readonly IReadOnlyDictionary<string, DesignTheme> Themes = new Dictionary<string, DesignTheme>(StringComparer.OrdinalIgnoreCase)
    {
        [Editorial.Name] = Editorial,
        [Corporate.Name] = Corporate
    };

    /// <summary>
    /// Resolves a theme name to its definition. Unknown names fall back to <see cref="Editorial"/>
    /// (the resolver records a deterministic warning); null resolves to the default theme.
    /// </summary>
    public static DesignTheme Resolve(string? name) =>
        name is not null && Themes.TryGetValue(name, out var theme) ? theme : Editorial;

    /// <summary>Returns the theme definition for a name, or null when unknown.</summary>
    public static DesignTheme? TryGet(string name) =>
        Themes.TryGetValue(name, out var theme) ? theme : null;

    private static Dictionary<TextRole, ThemeRoleFormatting> EditorialRoles() => new()
    {
        [TextRole.Title] = Role("display", 26, "primary", bold: true, after: 2),
        [TextRole.Subtitle] = Role("body", 13, "muted", before: 2, after: 14, line: 1.25),
        [TextRole.Eyebrow] = Role("display", 11, "coral", bold: true, allCaps: true, before: 2, after: 8),
        [TextRole.Heading1] = Role("display", 20, "primary", bold: true, before: 22, after: 6, keepNext: true, keepLines: true),
        [TextRole.Heading2] = Role("display", 16, "primary", bold: true, before: 18, after: 4, keepNext: true, keepLines: true),
        [TextRole.Heading3] = Role("display", 13, "teal", bold: true, before: 14, after: 4, keepNext: true, keepLines: true),
        [TextRole.Heading4] = Role("display", 11.5, "teal", bold: true, before: 10, after: 3, keepNext: true, keepLines: true),
        [TextRole.Heading5] = Role("display", 10.5, "ink", bold: true, before: 8, after: 2, keepNext: true, keepLines: true),
        [TextRole.Heading6] = Role("display", 10, "ink", bold: true, italic: true, before: 6, after: 2, keepNext: true, keepLines: true),
        [TextRole.Body] = Role("body", 11, "ink", after: 8, line: 1.3),
        [TextRole.Muted] = Role("body", 10.5, "muted", after: 6, line: 1.25),
        [TextRole.Label] = Role("body", 9.5, "accent", bold: true, after: 2),
        [TextRole.Metric] = Role("display", 24, "primary", bold: true, after: 2),
        [TextRole.MetricLabel] = Role("body", 9.5, "muted", after: 4),
        [TextRole.TableHeader] = Role("body", 10, "white", bold: true, after: 0, line: 1.15),
        [TextRole.TableBody] = Role("body", 10, "ink", after: 0, line: 1.15),
        [TextRole.Callout] = Role("body", 11, "ink", before: 2, after: 2, line: 1.25),
        [TextRole.Footer] = Role("body", 9, "muted", alignment: TextAlignment.Center, after: 0)
    };

    private static Dictionary<TextRole, ThemeRoleFormatting> CorporateRoles() => new()
    {
        [TextRole.Title] = Role("display", 24, "primary", bold: true, after: 2),
        [TextRole.Subtitle] = Role("body", 12.5, "muted", before: 2, after: 12, line: 1.2),
        [TextRole.Eyebrow] = Role("display", 10.5, "accent", bold: true, allCaps: true, before: 2, after: 8),
        [TextRole.Heading1] = Role("display", 18, "primary", bold: true, before: 20, after: 6, keepNext: true, keepLines: true),
        [TextRole.Heading2] = Role("display", 14.5, "primary", bold: true, before: 16, after: 4, keepNext: true, keepLines: true),
        [TextRole.Heading3] = Role("display", 12.5, "accent", bold: true, before: 12, after: 3, keepNext: true, keepLines: true),
        [TextRole.Heading4] = Role("display", 11, "accent", bold: true, before: 9, after: 3, keepNext: true, keepLines: true),
        [TextRole.Heading5] = Role("display", 10.5, "ink", bold: true, before: 7, after: 2, keepNext: true, keepLines: true),
        [TextRole.Heading6] = Role("display", 10, "ink", bold: true, italic: true, before: 6, after: 2, keepNext: true, keepLines: true),
        [TextRole.Body] = Role("body", 11, "ink", after: 8, line: 1.25),
        [TextRole.Muted] = Role("body", 10, "muted", after: 6, line: 1.2),
        [TextRole.Label] = Role("body", 9.5, "accent", bold: true, after: 2),
        [TextRole.Metric] = Role("display", 22, "primary", bold: true, after: 2),
        [TextRole.MetricLabel] = Role("body", 9, "muted", after: 4),
        [TextRole.TableHeader] = Role("body", 10, "white", bold: true, after: 0, line: 1.1),
        [TextRole.TableBody] = Role("body", 10, "ink", after: 0, line: 1.1),
        [TextRole.Callout] = Role("body", 11, "ink", before: 2, after: 2, line: 1.2),
        [TextRole.Footer] = Role("body", 9, "muted", alignment: TextAlignment.Center, after: 0)
    };

    private static ThemeRoleFormatting Role(
        string? font,
        double? size,
        string? color = null,
        bool bold = false,
        bool italic = false,
        bool allCaps = false,
        TextAlignment? alignment = null,
        double? before = null,
        double? after = null,
        double? line = null,
        bool keepNext = false,
        bool keepLines = false) =>
        new()
        {
            FontFamily = font,
            SizePt = size,
            Color = color,
            Bold = bold,
            Italic = italic,
            AllCaps = allCaps,
            Alignment = alignment,
            BeforePt = before,
            AfterPt = after,
            LineMultiple = line,
            KeepNext = keepNext,
            KeepLines = keepLines
        };
}
