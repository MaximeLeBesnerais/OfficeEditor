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

        if (BareHexPattern.IsMatch(color))
        {
            // Already resolved to hex by an earlier step; accept silently.
            return color;
        }
        return context.DesignResolver.TryResolveColor(color, path) is { } resolved
            ? StripHash(resolved)
            : null;
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

        return context.DesignResolver.TryResolveFontFamily(font, path);
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

        if (context.DesignResolver.TryResolveTypography(token, path) is { } resolved)
        {
            return new ResolvedTextFormat(
                resolved.FontFamily,
                resolved.FontSizePt,
                resolved.ColorHex is { } color ? StripHash(color) : null,
                resolved.Bold is true,
                resolved.Italic is true,
                resolved.Underline is true);
        }
        return null;
    }

    /// <summary>
    /// Default direct formatting for a heading level (used when the heading carries neither an
    /// explicit style nor a typography token): bold + a display font at a level-scaled size.
    /// </summary>
    public static ResolvedTextFormat HeadingDefaults(OoxmlEmitContext context, int level)
    {
        var size = Math.Max(24 - (level - 1) * 2, 12);
        var design = context.DesignResolver.ResolveAll();
        return new ResolvedTextFormat(
            design.DisplayFontFamilyOrDefault,
            size,
            design.Page.DefaultTextColorHex is { } color ? StripHash(color) : null,
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

        if (context.ResolveStyle(run.Style, StyleValues.Character, path) is { } runStyleId)
        {
            runProperties.RunStyle = new RunStyle { Val = runStyleId };
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

    private static string StripHash(string hex) => hex[0] == '#' ? hex[1..] : hex;

}
