using System.Globalization;
using System.Text.RegularExpressions;
using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Design;

/// <summary>
/// Resolves the generation vocabulary's design tokens into plain, immutable values emitters can
/// write out directly. It merges the active built-in theme (<see cref="DesignThemeCatalog"/>,
/// default <c>editorial</c>), the model's <see cref="PageSize.Default"/>/<see cref="Margins.Defaults"/>
/// and the document's <c>design</c> block, so callers never re-interpret palette/font/token
/// strings at emit time. The document's block overrides theme values field-by-field.
///
/// Resolution contract, deliberately symmetric with the parser's behavior:
/// <list type="bullet">
/// <item>Strict methods (<see cref="ResolveColor"/>, <see cref="ResolveTypography"/>,
/// <see cref="ResolveSpacing"/>) throw <see cref="DocxDesignResolutionException"/> on an unresolved
/// token name — use them when the emitter must have a value.</item>
/// <item><c>Try*</c> methods never throw; they record a deterministic
/// <see cref="DesignResolutionWarning"/> (off-palette hex, undefined font slot, unknown token) and
/// return a fallback — use them when the emitter can degrade. Warnings are readable from
/// <see cref="Warnings"/> and flow into <see cref="Contracts.DocxGenerationResult.Warnings"/>.</item>
/// </list>
///
/// The resolver is stateful only in its accumulated warning list; instance reuse across documents
/// would mix warnings, so create one per document (it is cheap — resolution is a single pass and
/// <see cref="ResolveAll"/> is cached).
/// </summary>
public sealed class DocxDesignResolver
{
    private static readonly Regex HexColorPattern = new(@"^#[0-9a-fA-F]{6}$", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private readonly DesignTokens? _design;
    private readonly List<DesignResolutionWarning> _warnings = [];
    private ResolvedDesign? _resolved;
    private DesignTheme? _theme;
    private IReadOnlyDictionary<string, string>? _mergedPalette;

    public DocxDesignResolver(DesignTokens? design = null)
    {
        _design = design;
    }

    /// <summary>
    /// Warnings accumulated across every resolution performed on this instance, in the order they
    /// were recorded. Deterministic for a given document and resolution walk.
    /// </summary>
    public IReadOnlyList<DesignResolutionWarning> Warnings => _warnings;

    /// <summary>Records a warning from another component of the subsystem (e.g. the style manager).</summary>
    internal void RecordWarning(DesignResolutionWarning warning) => _warnings.Add(warning);

    /// <summary>Records a warning from another component of the subsystem (e.g. the style manager).</summary>
    internal void RecordWarning(string code, string message, string? context = null) =>
        _warnings.Add(new DesignResolutionWarning(code, message, context));

    /// <summary>
    /// Resolves the entire design block once and caches the result. Theme, shape and page default
    /// tokens are resolved softly (warnings, never thrown) because they are fallbacks by nature;
    /// explicit references in content are resolved strictly through <see cref="ResolveColor"/> and friends.
    /// </summary>
    public ResolvedDesign ResolveAll()
    {
        if (_resolved is not null)
        {
            return _resolved;
        }

        var theme = ResolveTheme();
        var palette = ResolveMergedPalette();
        var designFonts = _design?.Fonts;
        var fonts = new FontTokens
        {
            Display = designFonts?.Display ?? theme.DisplayFontFamily,
            Body = designFonts?.Body ?? theme.BodyFontFamily
        };

        var typography = new Dictionary<string, ResolvedRunFormatting>(StringComparer.Ordinal);
        if (_design?.Typography is { } typographyTokens)
        {
            foreach (var (tokenName, token) in typographyTokens)
            {
                typography[tokenName] = ResolveTypographyToken(token, $"design.typography.{tokenName}");
            }
        }

        var shapes = _design?.Shapes ?? new ShapeDefaults();
        var resolvedShapes = new ResolvedShapeDefaults
        {
            CornerRadiusPt = shapes.CornerRadiusPt,
            DefaultFillHex = TryResolveColor(shapes.DefaultFill, "design.shapes.defaultFill"),
            DefaultStrokeColorHex = TryResolveColor(shapes.DefaultStrokeColor, "design.shapes.defaultStroke"),
            DefaultStrokeWidthPt = shapes.DefaultStrokeWidthPt
        };

        var layout = ResolveLayout(theme, _design?.Layout);

        var roles = new Dictionary<TextRole, ResolvedRoleFormatting>();
        foreach (var (role, roleFormatting) in theme.Roles)
        {
            roles[role] = ResolveRoleFormatting(roleFormatting, theme, palette, layout.DensityScale, $"theme '{theme.Name}' role '{role}'");
        }

        var page = _design?.Page ?? new PageDefaults();
        var resolvedPage = new ResolvedPageDefaults
        {
            PageSize = page.PageSize is { } sizeName ? PageSize.Named(sizeName) : PageSize.Named(theme.PageSize),
            Orientation = page.Orientation,
            Margins = page.Margins ?? theme.Margins,
            BodyFontFamily = TryResolveFontFamily(page.DefaultFontFamily, "design.page.defaultFont"),
            DefaultTextColorHex = TryResolveColor(page.DefaultTextColor, "design.page.defaultTextColor")
        };

        _resolved = new ResolvedDesign
        {
            ThemeName = theme.Name,
            Palette = palette,
            BodyFontFamily = fonts.Body,
            DisplayFontFamily = fonts.Display,
            Typography = typography,
            Spacing = _design?.Spacing ?? new Dictionary<string, double>(),
            Roles = roles,
            Layout = layout,
            Shapes = resolvedShapes,
            Page = resolvedPage
        };
        return _resolved;
    }

    // ---- theme ----

    /// <summary>
    /// The active built-in theme: <c>design.theme</c> when named and known, otherwise the default
    /// <c>editorial</c> theme. An unknown name records a deterministic warning and falls back.
    /// </summary>
    public DesignTheme ResolveTheme()
    {
        if (_theme is not null)
        {
            return _theme;
        }
        var name = _design?.Theme;
        if (name is not null && !DesignThemeCatalog.Themes.ContainsKey(name))
        {
            RecordWarning(
                "UnknownTheme",
                $"unknown theme '{name}'; falling back to the default '{DesignThemeCatalog.DefaultTheme}' theme.",
                "design.theme");
        }
        _theme = DesignThemeCatalog.Resolve(name);
        return _theme;
    }

    /// <summary>The theme palette merged with document-level overrides (document wins). Cached.</summary>
    private IReadOnlyDictionary<string, string> ResolveMergedPalette() =>
        _mergedPalette ??= ResolveTheme().EffectivePalette(_design?.Palette);

    /// <summary>
    /// Resolves a semantic text role to its theme-provided formatting. Returns null (with a
    /// warning) when the role is unknown — role names are validated by the parser, so this only
    /// fires for programmatically built models. Never throws.
    /// </summary>
    public ResolvedRoleFormatting? TryResolveRole(TextRole? role, string? context = null)
    {
        if (role is null)
        {
            return null;
        }
        if (ResolveAll().Roles.TryGetValue(role.Value, out var resolved))
        {
            return resolved;
        }
        RecordWarning("UnknownTextRole", $"unknown text role '{role}'.", context);
        return null;
    }

    /// <summary>Strict role resolution: <see cref="ResolvedRoleFormatting.Empty"/> for unknown roles.</summary>
    public ResolvedRoleFormatting ResolveRole(TextRole role, string? context = null) =>
        TryResolveRole(role, context) ?? ResolvedRoleFormatting.Empty;

    // ---- colors ----

    /// <summary>Resolves a palette token or #RRGGBB literal to a normalized #RRGGBB. Returns null
    /// (with a warning) when the value is neither; never throws. Palette tokens resolve against the
    /// theme palette merged with the document palette.</summary>
    public string? TryResolveColor(string? value, string? context = null)
    {
        if (value is null)
        {
            return null;
        }
        if (ResolveMergedPalette().TryGetValue(value, out var hex))
        {
            return NormalizeHex(hex);
        }
        if (HexColorPattern.IsMatch(value))
        {
            RecordWarning(
                "OffPaletteColor",
                $"raw hex color '{value}' is not declared in the design palette; emitters should prefer palette tokens.",
                context);
            return NormalizeHex(value);
        }
        RecordWarning(
            "UnknownColorToken",
            $"color '{value}' is neither a design palette token nor a #RRGGBB hex literal.",
            context);
        return null;
    }

    /// <summary>Strict color resolution: throws <see cref="DocxDesignResolutionException"/> when
    /// <paramref name="value"/> is not a palette token and not #RRGGBB hex.</summary>
    public string ResolveColor(string value, string? context = null) =>
        TryResolveColor(value, context)
        ?? throw new DocxDesignResolutionException(value, "color", Suggest(value, ResolveAll().Palette.Keys));

    // ---- fonts ----

    /// <summary>Resolves a font reference — the "display"/"body" slots or a raw family name — to a
    /// concrete family. An undefined slot falls back to the built-in font with a warning. Never throws.</summary>
    public string? TryResolveFontFamily(string? value, string? context = null)
    {
        if (value is null)
        {
            return null;
        }
        if (string.Equals(value, "display", StringComparison.OrdinalIgnoreCase))
        {
            var family = _design?.Fonts.Display ?? ResolveTheme().DisplayFontFamily;
            if (family is not null)
            {
                return family;
            }
            RecordWarning(
                "UndefinedFontSlot",
                $"font slot '{value}' is not defined in the design; falling back to the built-in '{DocxDesignDefaults.DisplayFont}'.",
                context);
            return DocxDesignDefaults.DisplayFont;
        }
        if (string.Equals(value, "body", StringComparison.OrdinalIgnoreCase))
        {
            var family = _design?.Fonts.Body ?? ResolveTheme().BodyFontFamily;
            if (family is not null)
            {
                return family;
            }
            RecordWarning(
                "UndefinedFontSlot",
                $"font slot '{value}' is not defined in the design; falling back to the built-in '{DocxDesignDefaults.BodyFont}'.",
                context);
            return DocxDesignDefaults.BodyFont;
        }
        return value;
    }

    /// <summary>Strict font resolution; an undefined slot still resolves (built-in font) with a
    /// warning — the only hard failure for fonts is a null input, which never occurs for the strict
    /// overload's callers.</summary>
    public string ResolveFontFamily(string value, string? context = null) =>
        TryResolveFontFamily(value, context) ?? DocxDesignDefaults.BodyFont;

    // ---- typography ----

    /// <summary>Resolves a named typography token to <see cref="ResolvedRunFormatting"/>. Returns
    /// null (with a warning) when the token is unknown; never throws.</summary>
    public ResolvedRunFormatting? TryResolveTypography(string? token, string? context = null)
    {
        if (token is null)
        {
            return null;
        }
        if (ResolveAll().Typography.TryGetValue(token, out var resolved))
        {
            return resolved;
        }
        RecordWarning("UnknownTypographyToken", $"unknown typography token '{token}'.", context);
        return null;
    }

    /// <summary>Strict typography resolution: throws <see cref="DocxDesignResolutionException"/>
    /// when the token is unknown.</summary>
    public ResolvedRunFormatting ResolveTypography(string token, string? context = null) =>
        TryResolveTypography(token, context)
        ?? throw new DocxDesignResolutionException(token, "typography token", Suggest(token, ResolveAll().Typography.Keys));

    private ResolvedRunFormatting ResolveTypographyToken(TypographyToken token, string context) =>
        new()
        {
            FontFamily = TryResolveFontFamily(token.FontFamily, context),
            FontSizePt = token.SizePt,
            ColorHex = TryResolveColor(token.Color, context),
            Bold = token.Bold ? true : null,
            Italic = token.Italic ? true : null,
            Underline = token.Underline ? true : null,
            AllCaps = token.AllCaps ? true : null
        };

    // ---- roles ----

    /// <summary>
    /// Resolves a role's theme formatting: font slot to the resolved family, color token to the
    /// merged palette hex, and paragraph spacing density-scaled. Theme roles are defined in the
    /// theme palette, so color/font resolution always succeeds.
    /// </summary>
    private static ResolvedRoleFormatting ResolveRoleFormatting(
        ThemeRoleFormatting formatting,
        DesignTheme theme,
        IReadOnlyDictionary<string, string> palette,
        double densityScale,
        string context)
    {
        var run = new ResolvedRunFormatting
        {
            FontFamily = ResolveRoleFont(formatting.FontFamily, theme),
            FontSizePt = formatting.SizePt,
            ColorHex = ResolvePaletteColor(formatting.Color, palette, context),
            Bold = formatting.Bold ? true : null,
            Italic = formatting.Italic ? true : null,
            Underline = formatting.Underline ? true : null,
            AllCaps = formatting.AllCaps ? true : null
        };
        var paragraph = new ResolvedParagraphFormatting
        {
            Alignment = formatting.Alignment,
            SpaceBeforePt = formatting.BeforePt is { } before ? before * densityScale : null,
            SpaceAfterPt = formatting.AfterPt is { } after ? after * densityScale : null,
            LineSpacingMultiple = formatting.LineMultiple
        };
        return new ResolvedRoleFormatting
        {
            Run = run,
            Paragraph = paragraph,
            KeepNext = formatting.KeepNext,
            KeepLines = formatting.KeepLines
        };
    }

    private static string? ResolveRoleFont(string? font, DesignTheme theme) =>
        font is null
            ? null
            : string.Equals(font, "display", StringComparison.OrdinalIgnoreCase)
                ? theme.DisplayFontFamily
                : string.Equals(font, "body", StringComparison.OrdinalIgnoreCase)
                    ? theme.BodyFontFamily
                    : font;

    private static string? ResolvePaletteColor(string? token, IReadOnlyDictionary<string, string> palette, string context)
    {
        if (token is null)
        {
            return null;
        }
        if (palette.TryGetValue(token, out var hex))
        {
            return NormalizeHex(hex);
        }
        if (HexColorPattern.IsMatch(token))
        {
            return NormalizeHex(token);
        }
        // Theme roles always reference palette tokens; this only guards against a document
        // overriding a token away. Fall back silently to the theme default.
        return null;
    }

    /// <summary>
    /// Resolves the layout guardrails/density: document values win, theme values fill gaps, and
    /// the density scale is derived deterministically.
    /// </summary>
    private static ResolvedLayout ResolveLayout(DesignTheme theme, LayoutDefaults? layout) =>
        new()
        {
            Density = layout?.Density ?? theme.Density,
            DensityScale = (layout?.Density ?? theme.Density).Scale(),
            MinBodySizePt = layout?.MinBodySizePt ?? theme.MinBodySizePt,
            MaxTableWidthPt = layout?.MaxTableWidthPt
        };

    // ---- spacing ----

    /// <summary>Resolves a spacing reference: a numeric pt value is passed through, a token name is
    /// looked up in <c>design.spacing</c>. Returns null (with a warning) when neither; never throws.</summary>
    public double? TryResolveSpacing(string? value, string? context = null)
    {
        if (value is null)
        {
            return null;
        }
        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) && number >= 0)
        {
            return number;
        }
        if (ResolveAll().Spacing.TryGetValue(value, out var pt))
        {
            return pt;
        }
        RecordWarning("UnknownSpacingToken", $"unknown spacing token '{value}'.", context);
        return null;
    }

    /// <summary>Strict spacing resolution: throws <see cref="DocxDesignResolutionException"/> when
    /// <paramref name="value"/> is neither a number nor a known token.</summary>
    public double ResolveSpacing(string value, string? context = null) =>
        TryResolveSpacing(value, context)
        ?? throw new DocxDesignResolutionException(value, "spacing token", Suggest(value, ResolveAll().Spacing.Keys));

    // ---- composed formatting ----

    /// <summary>
    /// Resolves one run's formatting, overlaying the optional typography token base with the run's
    /// own properties (run wins where set). This is the canonical way emitters obtain final run
    /// formatting for a run that sits under a paragraph token.
    /// </summary>
    public ResolvedRunFormatting ResolveRunFormatting(Run run, string? token, string? context = null)
    {
        var runFormatting = new ResolvedRunFormatting
        {
            FontFamily = TryResolveFontFamily(run.FontFamily, context),
            FontSizePt = run.FontSizePt,
            ColorHex = TryResolveColor(run.Color, context),
            Bold = run.Bold ? true : null,
            Italic = run.Italic ? true : null,
            Underline = run.Underline ? true : null,
            AllCaps = run.AllCaps ? true : null
        };
        var tokenFormatting = TryResolveTypography(token, context) ?? ResolvedRunFormatting.Empty;
        return tokenFormatting.Overlay(runFormatting);
    }

    /// <summary>Resolves a paragraph's alignment and spacing (already pt-valued in the model) to a
    /// plain <see cref="ResolvedParagraphFormatting"/>.</summary>
    public ResolvedParagraphFormatting ResolveParagraph(TextAlignment? alignment, ParagraphSpacing? spacing, string? context = null) =>
        new()
        {
            Alignment = alignment,
            SpaceBeforePt = spacing?.BeforePt,
            SpaceAfterPt = spacing?.AfterPt,
            LineSpacingMultiple = spacing?.LineMultiple
        };

    /// <summary>
    /// Resolves a full text content model to <see cref="ResolvedText"/>: token base formatting,
    /// per-run formatting (composed with the token), and paragraph formatting. Empty content yields
    /// <see cref="ResolvedText.Empty"/>.
    /// </summary>
    public ResolvedText ResolveText(TextModel? content, string? context = null)
    {
        if (content is null)
        {
            return ResolvedText.Empty;
        }
        var tokenFormatting = TryResolveTypography(content.Token, context) ?? ResolvedRunFormatting.Empty;
        var paragraph = ResolveParagraph(content.Alignment, content.Spacing, context);

        IReadOnlyList<ResolvedRun> runs;
        if (content.Text is not null)
        {
            runs = [new ResolvedRun(content.Text, tokenFormatting)];
        }
        else
        {
            var resolved = new List<ResolvedRun>(content.Runs?.Count ?? 0);
            if (content.Runs is not null)
            {
                foreach (var run in content.Runs)
                {
                    resolved.Add(new ResolvedRun(run.Text, ResolveRunFormatting(run, content.Token, context)));
                }
            }
            runs = resolved;
        }

        return new ResolvedText
        {
            Text = content.Text,
            Runs = runs,
            TokenFormatting = tokenFormatting,
            Paragraph = paragraph
        };
    }

    // ---- shapes ----

    /// <summary>
    /// Resolves a positioned primitive's fill/stroke/corner radius, merging explicit values with
    /// the design <c>shapes</c> defaults. The result is token-free and directly consumable by the
    /// positioned emitter.
    /// </summary>
    public ResolvedShapeStyle ResolveShape(string? fill, StrokeSpec? stroke, double? cornerRadiusPt, string? context = null)
    {
        var shapes = ResolveAll().Shapes;
        return new ResolvedShapeStyle
        {
            FillHex = TryResolveColor(fill, context) ?? shapes.DefaultFillHex,
            StrokeColorHex = TryResolveColor(stroke?.Color, context) ?? shapes.DefaultStrokeColorHex,
            StrokeWidthPt = stroke?.WidthPt ?? shapes.DefaultStrokeWidthPt,
            CornerRadiusPt = cornerRadiusPt ?? shapes.CornerRadiusPt
        };
    }

    // ---- page ----

    /// <summary>
    /// Resolves a section's page setup (or the design/built-in page defaults when the section omits
    /// it) to a plain <see cref="ResolvedPageFormat"/> with orientation-aware dimensions.
    /// </summary>
    public ResolvedPageFormat ResolvePage(PageSetup? setup, string? context = null)
    {
        var page = ResolveAll().Page;
        var size = setup?.PageSize ?? page.PageSize;
        var orientation = setup?.Orientation ?? page.Orientation;
        var margins = setup?.Margins ?? page.Margins;
        var (width, height) = size.EffectiveSize(orientation);
        return new ResolvedPageFormat
        {
            WidthPt = width,
            HeightPt = height,
            Orientation = orientation,
            Margins = margins,
            Columns = setup?.Columns,
            BreakType = setup?.BreakType
        };
    }

    private static string NormalizeHex(string hex) => hex.ToUpperInvariant();

    private static string? Suggest(string value, IEnumerable<string> candidates)
    {
        string? best = null;
        var bestDistance = int.MaxValue;
        foreach (var candidate in candidates)
        {
            var distance = Levenshtein(value, candidate);
            if (distance < bestDistance)
            {
                bestDistance = distance;
                best = candidate;
            }
        }
        return bestDistance <= 2 ? $"Did you mean '{best}'?" : null;
    }

    private static int Levenshtein(string a, string b)
    {
        a = a.ToLowerInvariant();
        b = b.ToLowerInvariant();
        var previous = new int[b.Length + 1];
        var current = new int[b.Length + 1];
        for (var j = 0; j <= b.Length; j++)
        {
            previous[j] = j;
        }
        for (var i = 1; i <= a.Length; i++)
        {
            current[0] = i;
            for (var j = 1; j <= b.Length; j++)
            {
                var cost = a[i - 1] == b[j - 1] ? 0 : 1;
                current[j] = Math.Min(Math.Min(current[j - 1] + 1, previous[j] + 1), previous[j - 1] + cost);
            }
            (previous, current) = (current, previous);
        }
        return previous[b.Length];
    }
}
