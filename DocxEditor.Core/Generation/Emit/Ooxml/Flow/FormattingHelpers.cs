using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Generation.Model;
using DocxEditor.Core.Generation.Schema;
using Model = DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Flow;

/// <summary>
/// A resolved bundle of default character formatting for a text run, derived from a
/// typography token or heading defaults. Individual run properties override these.
/// </summary>
internal sealed record ResolvedTextFormat(
    string? FontFamily,
    double? SizePt,
    string? Color,
    bool Bold,
    bool Italic,
    bool Underline);

/// <summary>
/// Local, replaceable OpenXML formatting helpers for the flow-first emitter. The design
/// branch may augment or replace these during integration; everything here is direct
/// (non-style) formatting, plus the blank-document default styles the flow branch owns.
/// </summary>
internal static class FormattingHelpers
{
    public const int TwipsPerPoint = 20;
    public const int HalfPointsPerPoint = 2;
    public const long EmusPerPoint = 12700;

    private static readonly Regex HexColorPattern = new(@"^#[0-9a-fA-F]{6}$", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    // Plain "RRGGBB" (no '#') only ever comes from a previous resolution step inside the
    // emitter (token/hex are resolved once, then carried as hex on ResolvedTextFormat).
    private static readonly Regex BareHexPattern = new(@"^[0-9a-fA-F]{6}$", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    /// <summary>Points → twips (1/1440 inch), as a schema string.</summary>
    public static string Twips(double points) =>
        ((int)Math.Round(points * TwipsPerPoint, MidpointRounding.AwayFromZero)).ToString(CultureInfo.InvariantCulture);

    /// <summary>Points → twips as an integer (page size/margin attributes are numeric).</summary>
    public static int TwipsInt(double points) =>
        (int)Math.Round(points * TwipsPerPoint, MidpointRounding.AwayFromZero);

    /// <summary>Points → half-points (font sizes), as a schema string.</summary>
    public static string HalfPoints(double points) =>
        ((int)Math.Round(points * HalfPointsPerPoint, MidpointRounding.AwayFromZero)).ToString(CultureInfo.InvariantCulture);

    /// <summary>Points → EMUs (1/914400 inch) for drawing extents.</summary>
    public static long Emus(double points) =>
        (long)Math.Round(points * EmusPerPoint, MidpointRounding.AwayFromZero);

    /// <summary>Line-spacing multiple → the 240ths-of-a-line schema value.</summary>
    public static string LineMultiple(double multiple) =>
        ((int)Math.Round(multiple * 240, MidpointRounding.AwayFromZero)).ToString(CultureInfo.InvariantCulture);

    /// <summary>Model alignment → paragraph justification (justify = both).</summary>
    public static JustificationValues? Justification(Model.TextAlignment? alignment) => alignment switch
    {
        Model.TextAlignment.Left => JustificationValues.Left,
        Model.TextAlignment.Center => JustificationValues.Center,
        Model.TextAlignment.Right => JustificationValues.Right,
        Model.TextAlignment.Justify => JustificationValues.Both,
        _ => null
    };

    /// <summary>Model alignment → table alignment (justify is not valid for tables).</summary>
    public static TableRowAlignmentValues? TableJustification(Model.TextAlignment? alignment) => alignment switch
    {
        Model.TextAlignment.Left => TableRowAlignmentValues.Left,
        Model.TextAlignment.Center => TableRowAlignmentValues.Center,
        Model.TextAlignment.Right => TableRowAlignmentValues.Right,
        _ => null
    };

    /// <summary>
    /// Resolves a color: a palette token name or a raw #RRGGBB literal → "RRGGBB". An
    /// unresolvable color warns and returns null (the property is skipped).
    /// </summary>
    public static string? ResolveColor(OoxmlEmitContext context, string? color, string path)
    {
        if (color is null)
        {
            return null;
        }

        if (context.Design?.Palette is { Count: > 0 } palette && palette.TryGetValue(color, out var hex))
        {
            return StripHash(hex);
        }

        if (HexColorPattern.IsMatch(color))
        {
            return StripHash(color);
        }

        if (BareHexPattern.IsMatch(color))
        {
            // Already resolved to hex by an earlier step; accept silently.
            return color;
        }

        context.Warn(path, $"color '{color}' could not be resolved (not a palette token and not #RRGGBB); the color was skipped.");
        return null;
    }

    /// <summary>
    /// Resolves a font reference: the "display"/"body" slots resolve to the design font slot,
    /// anything else is treated as a raw family name. An undefined slot warns and returns null.
    /// </summary>
    public static string? ResolveFontFamily(OoxmlEmitContext context, string? font, string path)
    {
        if (font is null)
        {
            return null;
        }

        if (string.Equals(font, "display", StringComparison.OrdinalIgnoreCase))
        {
            if (context.Design?.Fonts.Display is { } display)
            {
                return display;
            }
            context.Warn(path, "font slot 'display' is not defined in the design tokens; the font was skipped.");
            return null;
        }

        if (string.Equals(font, "body", StringComparison.OrdinalIgnoreCase))
        {
            if (context.Design?.Fonts.Body is { } body)
            {
                return body;
            }
            context.Warn(path, "font slot 'body' is not defined in the design tokens; the font was skipped.");
            return null;
        }

        return font;
    }

    /// <summary>
    /// Resolves a typography token into a <see cref="ResolvedTextFormat"/> applying to the
    /// paragraph's runs by default. Unknown tokens warn (the parser already errors, but the
    /// model can be built programmatically) and return null.
    /// </summary>
    public static ResolvedTextFormat? ResolveTypographyToken(OoxmlEmitContext context, string? token, string path)
    {
        if (token is null)
        {
            return null;
        }

        if (context.Design?.Typography is { Count: > 0 } typography && typography.TryGetValue(token, out var resolved))
        {
            return new ResolvedTextFormat(
                ResolveFontFamily(context, resolved.FontFamily, path),
                resolved.SizePt,
                ResolveColor(context, resolved.Color, path),
                resolved.Bold,
                resolved.Italic,
                resolved.Underline);
        }

        context.Warn(path, $"typography token '{token}' is not defined in the design tokens; its formatting was skipped.");
        return null;
    }

    /// <summary>
    /// Default direct formatting for a heading level (used when the heading carries neither an
    /// explicit style nor a typography token): bold + a display font at a level-scaled size.
    /// </summary>
    public static ResolvedTextFormat HeadingDefaults(OoxmlEmitContext context, int level)
    {
        var size = Math.Max(24 - (level - 1) * 2, 12);
        return new ResolvedTextFormat(
            ResolveFontFamily(context, context.Design?.Fonts.Display, null, silent: true),
            size,
            ResolveColor(context, context.Design?.Page.DefaultTextColor, null, silent: true),
            Bold: true,
            Italic: false,
            Underline: false);
    }

    /// <summary>Builds one run from a model run, merging the token/heading defaults.</summary>
    public static DocumentFormat.OpenXml.Wordprocessing.Run BuildRun(
        OoxmlEmitContext context,
        Model.Run run,
        ResolvedTextFormat? defaults,
        string path)
    {
        var runProperties = new RunProperties();

        if (run.Style is not null && context.TryEnsureStyle(run.Style, StyleValues.Character, path))
        {
            runProperties.RunStyle = new RunStyle { Val = run.Style };
        }

        var font = ResolveFontFamily(context, run.FontFamily ?? defaults?.FontFamily, path);
        if (font is not null)
        {
            runProperties.RunFonts = new RunFonts { Ascii = font, HighAnsi = font, EastAsia = font };
        }

        var sizePt = run.FontSizePt ?? defaults?.SizePt;
        if (sizePt is { } size)
        {
            var halfPoints = HalfPoints(size);
            runProperties.FontSize = new FontSize { Val = halfPoints };
            runProperties.FontSizeComplexScript = new FontSizeComplexScript { Val = halfPoints };
        }

        var color = ResolveColor(context, run.Color ?? defaults?.Color, path);
        if (color is not null)
        {
            runProperties.Color = new Color { Val = color };
        }

        if (run.Bold || defaults?.Bold is true)
        {
            runProperties.Bold = new Bold();
        }
        if (run.Italic || defaults?.Italic is true)
        {
            runProperties.Italic = new Italic();
        }
        if (run.Underline || defaults?.Underline is true)
        {
            runProperties.Underline = new Underline { Val = UnderlineValues.Single };
        }

        var wordRun = new DocumentFormat.OpenXml.Wordprocessing.Run();
        if (runProperties.HasChildren)
        {
            wordRun.Append(runProperties);
        }
        AppendTextWithBreaks(wordRun, run.Text);
        return wordRun;
    }

    /// <summary>Builds a run from plain text using the token/heading defaults.</summary>
    public static DocumentFormat.OpenXml.Wordprocessing.Run BuildTextRun(
        OoxmlEmitContext context,
        string text,
        ResolvedTextFormat? defaults,
        string path)
    {
        var wordRun = new DocumentFormat.OpenXml.Wordprocessing.Run();
        var runProperties = new RunProperties();

        var font = ResolveFontFamily(context, defaults?.FontFamily, path);
        if (font is not null)
        {
            runProperties.RunFonts = new RunFonts { Ascii = font, HighAnsi = font, EastAsia = font };
        }

        if (defaults?.SizePt is { } size)
        {
            var halfPoints = HalfPoints(size);
            runProperties.FontSize = new FontSize { Val = halfPoints };
            runProperties.FontSizeComplexScript = new FontSizeComplexScript { Val = halfPoints };
        }

        var color = ResolveColor(context, defaults?.Color, path);
        if (color is not null)
        {
            runProperties.Color = new Color { Val = color };
        }

        if (defaults?.Bold is true)
        {
            runProperties.Bold = new Bold();
        }
        if (defaults?.Italic is true)
        {
            runProperties.Italic = new Italic();
        }
        if (defaults?.Underline is true)
        {
            runProperties.Underline = new Underline { Val = UnderlineValues.Single };
        }

        if (runProperties.HasChildren)
        {
            wordRun.Append(runProperties);
        }
        AppendTextWithBreaks(wordRun, text);
        return wordRun;
    }

    /// <summary>Appends text segments to a run, converting newlines to <c>w:br</c>.</summary>
    public static void AppendTextWithBreaks(DocumentFormat.OpenXml.Wordprocessing.Run run, string text)
    {
        var segments = text.Split('\n');
        for (var i = 0; i < segments.Length; i++)
        {
            if (i > 0)
            {
                run.Append(new Break());
            }

            var segment = segments[i];
            if (segment.Length > 0)
            {
                run.Append(new Text(segment) { Space = SpaceProcessingModeValues.Preserve });
            }
        }
    }

    /// <summary>
    /// The blank-document default "Normal" paragraph style, seeded from the design page
    /// defaults (font slot/family + text color) when present.
    /// </summary>
    public static Style BuildNormalDefaultStyle(DesignTokens? design)
    {
        var runProperties = new StyleRunProperties();

        var fontSlot = design?.Page.DefaultFontFamily;
        var family = fontSlot switch
        {
            null => null,
            "display" => design?.Fonts.Display,
            "body" => design?.Fonts.Body,
            _ => fontSlot
        };
        if (family is not null)
        {
            runProperties.RunFonts = new RunFonts { Ascii = family, HighAnsi = family, EastAsia = family };
        }

        if (design?.Page.DefaultTextColor is { } colorToken && ResolveColorSilently(design, colorToken) is { } hex)
        {
            runProperties.Color = new Color { Val = hex };
        }

        // CT_Style child order: name, aliases, basedOn, …, pPr, rPr, ….
        return new Style(
            new StyleName { Val = "Normal" },
            runProperties)
        {
            Type = StyleValues.Paragraph,
            StyleId = "Normal",
            Default = true
        };
    }

    /// <summary>
    /// On-demand default style for a blank document when an explicitly referenced style ID
    /// does not exist. Templates never reach this path (unknown template styles warn + skip).
    /// </summary>
    public static Style BuildDefaultStyle(string styleId, StyleValues kind)
    {
        if (styleId.StartsWith("Heading", StringComparison.Ordinal) &&
            int.TryParse(styleId.AsSpan("Heading".Length), out var level) && level is >= 1 and <= 6)
        {
            var before = (240 - (level - 1) * 30).ToString(CultureInfo.InvariantCulture);
            var size = Math.Max(32 - (level - 1) * 4, 18).ToString(CultureInfo.InvariantCulture);
            return new Style(
                new StyleName { Val = $"heading {level}" },
                new BasedOn { Val = "Normal" },
                new StyleParagraphProperties(
                    new KeepNext(),
                    new SpacingBetweenLines { Before = before, After = "60" },
                    new OutlineLevel { Val = level - 1 }),
                new StyleRunProperties(
                    new Bold(),
                    new BoldComplexScript(),
                    new FontSize { Val = size },
                    new FontSizeComplexScript { Val = size }))
            {
                Type = StyleValues.Paragraph,
                StyleId = styleId
            };
        }

        if (kind == StyleValues.Table)
        {
            return new Style(
                new StyleName { Val = styleId },
                new StyleTableProperties(
                    new TableBorders(
                        new TopBorder { Val = BorderValues.Single, Size = 4 },
                        new LeftBorder { Val = BorderValues.Single, Size = 4 },
                        new BottomBorder { Val = BorderValues.Single, Size = 4 },
                        new RightBorder { Val = BorderValues.Single, Size = 4 },
                        new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                        new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 })))
            {
                Type = StyleValues.Table,
                StyleId = styleId
            };
        }

        if (kind == StyleValues.Character)
        {
            return new Style(
                new StyleName { Val = styleId })
            {
                Type = StyleValues.Character,
                StyleId = styleId
            };
        }

        return new Style(
            new StyleName { Val = styleId },
            new BasedOn { Val = "Normal" })
        {
            Type = StyleValues.Paragraph,
            StyleId = styleId
        };
    }

    private static string StripHash(string hex) => hex[0] == '#' ? hex[1..] : hex;

    private static string? ResolveColorSilently(DesignTokens design, string color)
    {
        if (design.Palette.TryGetValue(color, out var hex))
        {
            return StripHash(hex);
        }
        return HexColorPattern.IsMatch(color) ? StripHash(color) : null;
    }

    private static string? ResolveFontFamily(OoxmlEmitContext context, string? font, string? path, bool silent)
    {
        return silent ? ResolveFontFamilySilently(context, font) : ResolveFontFamily(context, font, path ?? string.Empty);
    }

    private static string? ResolveFontFamilySilently(OoxmlEmitContext context, string? font)
    {
        if (font is null)
        {
            return null;
        }
        return font.ToLowerInvariant() switch
        {
            "display" => context.Design?.Fonts.Display,
            "body" => context.Design?.Fonts.Body,
            _ => font
        };
    }

    private static string? ResolveColor(OoxmlEmitContext context, string? color, string? path, bool silent)
    {
        return silent ? ResolveColorSilently(context.Design ?? new DesignTokens(), color ?? string.Empty) : ResolveColor(context, color, path ?? string.Empty);
    }
}
