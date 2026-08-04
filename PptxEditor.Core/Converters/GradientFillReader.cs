using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using PptxEditor.Core.Models;

namespace PptxEditor.Core.Converters;

/// <summary>
/// Shared reader for OOXML linear gradient fills (<c>a:gradFill</c>) on shape
/// properties. Both the main slide-shape path (<see cref="PptxToTypstConverter"/>,
/// typed <c>p:spPr</c>) and the SmartArt drawing extractor (untyped <c>dsp:spPr</c>)
/// call this so the two paths produce identical <see cref="TypstGradientFill"/>
/// values for the same markup.
///
/// Also hosts the shared DrawingML color helpers (<see cref="ParseHexColor"/>,
/// <see cref="ApplyColorModifiers"/>, <see cref="FormatHexColor"/>) so gradient
/// stops and solid fills apply tint/alpha transforms identically.
/// </summary>
internal static class GradientFillReader
{
    private const string DrawingmlNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// <summary>
    /// Reads a linear gradient fill (<c>a:gradFill</c>) from the given shape-properties
    /// element, including per-stop <c>a:alpha</c> opacities. Returns null when there is
    /// no gradient, when it has fewer than two resolvable stops, or when a stop color
    /// cannot be resolved.
    /// Decks like AetherLink use full-bleed gradient rectangles as slide backgrounds;
    /// SmartArt drawing parts use gradients as the primary shape fill.
    /// </summary>
    /// <param name="shapeProperties">Shape-properties element whose direct
    /// <c>a:gradFill</c> child is read (any element type — typed
    /// <c>p:spPr</c> or untyped <c>dsp:spPr</c>).</param>
    /// <param name="resolveSchemeColor">Scheme color name → <c>#RRGGBB</c>, or null
    /// when scheme colors cannot be resolved (stops referencing them then fail and
    /// the whole gradient is rejected, matching the previous behavior).</param>
    internal static TypstGradientFill? TryReadLinearGradient(
        OpenXmlElement shapeProperties,
        Func<string, string?>? resolveSchemeColor,
        bool themeStyleCompatibility = false)
    {
        var gradFill = GetChild(shapeProperties, "gradFill");
        var stopList = gradFill == null ? null : GetChild(gradFill, "gsLst");
        if (stopList == null)
            return null;

        var stops = new List<TypstGradientStop>();
        foreach (var gs in stopList.Elements()
                     .Where(e => e.LocalName == "gs" && e.NamespaceUri == DrawingmlNs))
        {
            var color = ReadStopColor(gs, resolveSchemeColor, themeStyleCompatibility);
            if (color == null)
                return null;

            var pos = ReadLongAttribute(gs, "pos") ?? 0;
            stops.Add(new TypstGradientStop(color, Math.Clamp(pos / 100000.0, 0.0, 1.0)));
        }

        if (stops.Count < 2)
            return null;

        // OOXML a:lin ang is in 60000ths of a degree; OOXML and Typst gradient.linear
        // both measure the axis clockwise from the left→right direction, so degrees
        // pass through verbatim (same convention as the generation pipeline's emitters).
        var lin = GetChild(gradFill!, "lin");
        var angle = (lin == null ? 0 : ReadLongAttribute(lin, "ang") ?? 0) / 60000.0;
        return new TypstGradientFill(angle, stops);
    }

    private static string? ReadStopColor(
        OpenXmlElement stop,
        Func<string, string?>? resolveSchemeColor,
        bool themeStyleCompatibility)
    {
        if (TryResolveBaseColor(stop, resolveSchemeColor, out var baseColor, out var modifiers))
        {
            // Unlike a solid fill, a fully transparent gradient stop is meaningful
            // (fade-out) — keep it as #RRGGBB00 instead of treating it as noFill.
            return ApplyColorModifiers(baseColor, modifiers, themeStyleCompatibility)
                ?? FormatHexColor(baseColor, 0);
        }

        return null;
    }

    /// <summary>
    /// Resolves one DrawingML color container. The order is the canonical OOXML
    /// order used throughout the converter: explicit sRGB, scheme color, system
    /// color's portable <c>lastClr</c>, then preset color. Modifiers remain attached
    /// to the selected color element and are applied in document order.
    /// </summary>
    internal static string? ResolveColor(
        OpenXmlElement colorContainer,
        Func<string, string?>? resolveSchemeColor = null,
        bool themeStyleCompatibility = false)
    {
        return TryResolveBaseColor(colorContainer, resolveSchemeColor, out var color, out var modifiers)
            ? ApplyColorModifiers(color, modifiers, themeStyleCompatibility)
            : null;
    }

    private static bool TryResolveBaseColor(
        OpenXmlElement colorContainer,
        Func<string, string?>? resolveSchemeColor,
        out (byte R, byte G, byte B) color,
        out IEnumerable<OpenXmlElement> modifiers)
    {
        var colorElements = colorContainer.LocalName is
            ("srgbClr" or "scrgbClr" or "hslClr" or "schemeClr" or "sysClr" or "prstClr")
            ? new[] { colorContainer }
            : colorContainer.ChildElements.Cast<OpenXmlElement>();
        foreach (var colorElement in colorElements)
        {
            if (colorElement.LocalName is not
                    ("srgbClr" or "scrgbClr" or "hslClr" or "schemeClr" or "sysClr" or "prstClr")
                || colorElement.NamespaceUri != DrawingmlNs)
                continue;

            var raw = ReadAttribute(colorElement, "val");
            var resolved = colorElement.LocalName switch
            {
                "srgbClr" => TryParseHexColor(raw, out var srgb) ? FormatHexColor(srgb) : null,
                "scrgbClr" => TryReadScrgbColor(colorElement, out var scrgb) ? FormatHexColor(scrgb) : null,
                "hslClr" => TryReadHslColor(colorElement, out var hsl) ? FormatHexColor(hsl) : null,
                "schemeClr" => string.IsNullOrEmpty(raw) ? null : resolveSchemeColor?.Invoke(raw),
                "sysClr" => IsSixHexDigits(ReadAttribute(colorElement, "lastClr"))
                    ? ReadAttribute(colorElement, "lastClr")
                    : null,
                "prstClr" => PresetColors.TryGetValue(raw ?? string.Empty, out var preset) ? preset : null,
                _ => null
            };

            if (TryParseHexColor(resolved, out color))
            {
                modifiers = colorElement.ChildElements;
                return true;
            }
        }

        color = default;
        modifiers = Array.Empty<OpenXmlElement>();
        return false;
    }

    private static bool TryReadScrgbColor(
        OpenXmlElement colorElement, out (byte R, byte G, byte B) color)
    {
        color = default;
        return TryReadPercentage(colorElement, "r", out var red)
            && TryReadPercentage(colorElement, "g", out var green)
            && TryReadPercentage(colorElement, "b", out var blue)
            && SetColor(red, green, blue, out color);
    }

    private static bool TryReadHslColor(
        OpenXmlElement colorElement, out (byte R, byte G, byte B) color)
    {
        var hue = ReadAttribute(colorElement, "hue");
        var saturation = ReadAttribute(colorElement, "sat");
        var luminance = ReadAttribute(colorElement, "lum");
        if (!double.TryParse(hue, NumberStyles.Float, CultureInfo.InvariantCulture, out var hueValue)
            || !double.TryParse(saturation, NumberStyles.Float, CultureInfo.InvariantCulture, out var saturationValue)
            || !double.TryParse(luminance, NumberStyles.Float, CultureInfo.InvariantCulture, out var luminanceValue))
        {
            color = default;
            return false;
        }

        color = HslToRgb(
            hueValue / 60000.0,
            Math.Clamp(saturationValue / 100000.0, 0.0, 1.0),
            Math.Clamp(luminanceValue / 100000.0, 0.0, 1.0));
        return true;
    }

    private static bool TryReadPercentage(OpenXmlElement element, string attributeName, out double value)
        => double.TryParse(
            ReadAttribute(element, attributeName),
            NumberStyles.Float,
            CultureInfo.InvariantCulture,
            out value)
            && value is >= 0 and <= 100000;

    private static bool SetColor(double red, double green, double blue, out (byte R, byte G, byte B) color)
    {
        color = (
            Clamp(red / 100000.0 * 255.0),
            Clamp(green / 100000.0 * 255.0),
            Clamp(blue / 100000.0 * 255.0));
        return true;
    }

    private static bool IsSixHexDigits(string? value)
        => value != null && Regex.IsMatch(value, "^[0-9A-Fa-f]{6}$", RegexOptions.CultureInvariant);

    private static readonly IReadOnlyDictionary<string, string> PresetColors =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["black"] = "#000000", ["white"] = "#FFFFFF", ["red"] = "#FF0000",
            ["green"] = "#008000", ["blue"] = "#0000FF", ["yellow"] = "#FFFF00",
            ["cyan"] = "#00FFFF", ["magenta"] = "#FF00FF", ["gray"] = "#808080",
            ["grey"] = "#808080", ["dkGray"] = "#404040", ["darkGray"] = "#404040",
            ["ltGray"] = "#C0C0C0", ["lightGray"] = "#C0C0C0",
            ["orange"] = "#FFA500", ["purple"] = "#800080", ["brown"] = "#A52A2A",
            ["pink"] = "#FFC0CB", ["gold"] = "#FFD700", ["navy"] = "#000080",
            ["olive"] = "#808000", ["maroon"] = "#800000", ["teal"] = "#008080",
            ["silver"] = "#C0C0C0", ["lime"] = "#00FF00", ["aqua"] = "#00FFFF"
            , ["aliceBlue"] = "#F0F8FF", ["antiqueWhite"] = "#FAEBD7", ["aquaMarine"] = "#7FFFD4",
            ["azure"] = "#F0FFFF", ["beige"] = "#F5F5DC", ["bisque"] = "#FFE4C4",
            ["blanchedAlmond"] = "#FFEBCD", ["blueViolet"] = "#8A2BE2", ["burlyWood"] = "#DEB887",
            ["cadetBlue"] = "#5F9EA0", ["chartreuse"] = "#7FFF00", ["chocolate"] = "#D2691E",
            ["coral"] = "#FF7F50", ["cornflowerBlue"] = "#6495ED", ["cornsilk"] = "#FFF8DC",
            ["crimson"] = "#DC143C", ["darkBlue"] = "#00008B", ["darkCyan"] = "#008B8B",
            ["darkGoldenrod"] = "#B8860B", ["darkGreen"] = "#006400", ["darkKhaki"] = "#BDB76B",
            ["darkMagenta"] = "#8B008B", ["darkOliveGreen"] = "#556B2F", ["darkOrange"] = "#FF8C00",
            ["darkOrchid"] = "#9932CC", ["darkRed"] = "#8B0000", ["darkSalmon"] = "#E9967A",
            ["darkSeaGreen"] = "#8FBC8F", ["darkSlateBlue"] = "#483D8B", ["darkSlateGray"] = "#2F4F4F",
            ["darkTurquoise"] = "#00CED1", ["darkViolet"] = "#9400D3", ["deepPink"] = "#FF1493",
            ["deepSkyBlue"] = "#00BFFF", ["dimGray"] = "#696969", ["dodgerBlue"] = "#1E90FF",
            ["firebrick"] = "#B22222", ["floralWhite"] = "#FFFAF0", ["forestGreen"] = "#228B22",
            ["gainsboro"] = "#DCDCDC", ["ghostWhite"] = "#F8F8FF", ["goldenrod"] = "#DAA520",
            ["greenYellow"] = "#ADFF2F", ["honeydew"] = "#F0FFF0", ["hotPink"] = "#FF69B4",
            ["indianRed"] = "#CD5C5C", ["indigo"] = "#4B0082", ["ivory"] = "#FFFFF0",
            ["khaki"] = "#F0E68C", ["lavender"] = "#E6E6FA", ["lavenderBlush"] = "#FFF0F5",
            ["lawnGreen"] = "#7CFC00", ["lemonChiffon"] = "#FFFACD", ["lightBlue"] = "#ADD8E6",
            ["lightCoral"] = "#F08080", ["lightCyan"] = "#E0FFFF", ["lightGoldenrodYellow"] = "#FAFAD2",
            ["lightGreen"] = "#90EE90", ["lightPink"] = "#FFB6C1", ["lightSalmon"] = "#FFA07A",
            ["lightSeaGreen"] = "#20B2AA", ["lightSkyBlue"] = "#87CEFA", ["lightSlateGray"] = "#778899",
            ["lightSteelBlue"] = "#B0C4DE", ["lightYellow"] = "#FFFFE0", ["limeGreen"] = "#32CD32",
            ["linen"] = "#FAF0E6", ["mediumAquamarine"] = "#66CDAA", ["mediumBlue"] = "#0000CD",
            ["mediumOrchid"] = "#BA55D3", ["mediumPurple"] = "#9370DB", ["mediumSeaGreen"] = "#3CB371",
            ["mediumSlateBlue"] = "#7B68EE", ["mediumSpringGreen"] = "#00FA9A", ["mediumTurquoise"] = "#48D1CC",
            ["mediumVioletRed"] = "#C71585", ["midnightBlue"] = "#191970", ["mintCream"] = "#F5FFFA",
            ["mistyRose"] = "#FFE4E1", ["moccasin"] = "#FFE4B5", ["navajoWhite"] = "#FFDEAD",
            ["oldLace"] = "#FDF5E6", ["oliveDrab"] = "#6B8E23", ["orangeRed"] = "#FF4500",
            ["orchid"] = "#DA70D6", ["paleGoldenrod"] = "#EEE8AA", ["paleGreen"] = "#98FB98",
            ["paleTurquoise"] = "#AFEEEE", ["paleVioletRed"] = "#DB7093", ["papayaWhip"] = "#FFEFD5",
            ["peachPuff"] = "#FFDAB9", ["peru"] = "#CD853F", ["plum"] = "#DDA0DD",
            ["powderBlue"] = "#B0E0E6", ["rosyBrown"] = "#BC8F8F", ["royalBlue"] = "#4169E1",
            ["saddleBrown"] = "#8B4513", ["salmon"] = "#FA8072", ["sandyBrown"] = "#F4A460",
            ["seaGreen"] = "#2E8B57", ["seashell"] = "#FFF5EE", ["sienna"] = "#A0522D",
            ["skyBlue"] = "#87CEEB", ["slateBlue"] = "#6A5ACD", ["slateGray"] = "#708090",
            ["snow"] = "#FFFAFA", ["springGreen"] = "#00FF7F", ["steelBlue"] = "#4682B4",
            ["tan"] = "#D2B48C", ["thistle"] = "#D8BFD8", ["tomato"] = "#FF6347",
            ["turquoise"] = "#40E0D0", ["violet"] = "#EE82EE", ["wheat"] = "#F5DEB3",
            ["whiteSmoke"] = "#F5F5F5", ["yellowGreen"] = "#9ACD32"
        };

    internal static (byte R, byte G, byte B) ParseHexColor(string hex)
    {
        return TryParseHexColor(hex, out var color) ? color : ((byte)0, (byte)0, (byte)0);
    }

    internal static bool TryParseHexColor(string? hex, out (byte R, byte G, byte B) color)
    {
        var value = hex?.TrimStart('#');
        if (value == null || !IsSixHexDigits(value))
        {
            color = default;
            return false;
        }

        color = (
            byte.Parse(value[..2], NumberStyles.HexNumber, CultureInfo.InvariantCulture),
            byte.Parse(value[2..4], NumberStyles.HexNumber, CultureInfo.InvariantCulture),
            byte.Parse(value[4..6], NumberStyles.HexNumber, CultureInfo.InvariantCulture));
        return true;
    }

    internal static string FormatHexColor((byte R, byte G, byte B) color, byte alpha = 255)
    {
        if (alpha == 255)
            return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
        return $"#{color.R:X2}{color.G:X2}{color.B:X2}{alpha:X2}";
    }

    /// <summary>
    /// Applies DrawingML color transforms (<c>a:tint</c>, <c>a:shade</c>,
    /// <c>a:satMod</c>, <c>a:lumMod</c>, <c>a:lumOff</c>, <c>a:alpha</c>) in
    /// document order. Returns null when the result is fully transparent
    /// (visually identical to <c>a:noFill</c> for solid fills) — gradient callers
    /// substitute an explicit <c>#RRGGBB00</c> stop instead.
    /// </summary>
    internal static string? ApplyColorModifiers(
        (byte R, byte G, byte B) color,
        IEnumerable<OpenXmlElement> modifiers,
        bool themeStyleCompatibility = false)
    {
        var alpha = 1.0;
        foreach (var mod in modifiers)
        {
            switch (mod.LocalName)
            {
                case "tint":
                    if (TryGetOoxmlVal(mod, out var tintVal))
                        color = ApplyTint(color, tintVal);
                    break;
                case "alpha":
                    if (TryGetOoxmlVal(mod, out var alphaVal))
                        alpha = Math.Clamp(alphaVal / 100000.0, 0.0, 1.0);
                    break;
                case "alphaMod":
                    if (TryGetOoxmlVal(mod, out var alphaModVal))
                        alpha = Math.Clamp(alpha * alphaModVal / 100000.0, 0.0, 1.0);
                    break;
                case "alphaOff":
                    if (TryGetOoxmlVal(mod, out var alphaOffVal))
                        alpha = Math.Clamp(alpha + alphaOffVal / 100000.0, 0.0, 1.0);
                    break;
                case "shade": // darken: exact HSL-lightness scale when L<0.5 (no clamping)
                    if (TryGetOoxmlVal(mod, out var shadeVal))
                        color = themeStyleCompatibility
                            ? ApplyThemeStyleShade(color, shadeVal / 100000.0)
                            : ScaleChannels(color, shadeVal / 100000.0);
                    break;
                case "satMod": // HSL saturation multiply (hue/lightness preserved)
                    if (TryGetOoxmlVal(mod, out var satModVal))
                        color = ApplySatMod(color, satModVal / 100000.0);
                    break;
                case "lumMod":
                    if (TryGetOoxmlVal(mod, out var lumModVal))
                        color = ScaleChannels(color, lumModVal / 100000.0);
                    break;
                case "lumOff":
                    if (TryGetOoxmlVal(mod, out var lumOffVal))
                        color = OffsetChannels(color, lumOffVal / 100000.0);
                    break;
                case "hue":
                    if (TryGetOoxmlVal(mod, out var hueVal))
                        color = ApplyHslTransform(color, h => hueVal / 60000.0, null, null);
                    break;
                case "hueMod":
                    if (TryGetOoxmlVal(mod, out var hueModVal))
                        color = ApplyHslTransform(color, h => h * hueModVal / 100000.0, null, null);
                    break;
                case "hueOff":
                    if (TryGetOoxmlVal(mod, out var hueOffVal))
                        color = ApplyHslTransform(color, h => h + hueOffVal / 60000.0, null, null);
                    break;
                case "sat":
                    if (TryGetOoxmlVal(mod, out var satVal))
                        color = ApplyHslTransform(color, null, _ => satVal / 100000.0, null);
                    break;
                case "satOff":
                    if (TryGetOoxmlVal(mod, out var satOffVal))
                        color = ApplyHslTransform(color, null, s => s + satOffVal / 100000.0, null);
                    break;
                case "lum":
                    if (TryGetOoxmlVal(mod, out var lumVal))
                        color = ApplyHslTransform(color, null, null, _ => lumVal / 100000.0);
                    break;
            }
        }
        // Fully transparent solid fill is visually identical to a:noFill — report no
        // color so callers treat the fill/stroke/text color as absent.
        return alpha <= 0 ? null : FormatHexColor(color, ClampAlpha(alpha));
    }

    internal static (byte R, byte G, byte B) ApplyTint((byte R, byte G, byte B) color, int tint)
    {
        // DrawingML tint values specify how much of the source color to keep;
        // the remainder is blended toward white. Blend in linear light so very
        // light tints match PowerPoint's rendered colors more closely.
        double sourceWeight = Math.Clamp(tint / 100000.0, 0.0, 1.0);
        return (
            ApplyTintChannel(color.R, sourceWeight),
            ApplyTintChannel(color.G, sourceWeight),
            ApplyTintChannel(color.B, sourceWeight)
        );
    }

    private static byte ApplyTintChannel(byte channel, double sourceWeight)
    {
        var linear = Math.Pow(channel / 255.0, 2.2);
        var tinted = linear * sourceWeight + (1.0 - sourceWeight);
        return (byte)Math.Round(Math.Pow(tinted, 1.0 / 2.2) * 255.0, MidpointRounding.AwayFromZero);
    }

    // lumMod/lumOff are HSL-luminance transforms in the spec; the per-channel
    // approximation matches PowerPoint for grayscale bases (e.g. bg1 + lumMod 75%
    // = #BFBFBF) and stays close for colored bases.
    private static (byte R, byte G, byte B) ScaleChannels((byte R, byte G, byte B) color, double factor)
    {
        return (Clamp(color.R * factor), Clamp(color.G * factor), Clamp(color.B * factor));
    }

    /// <summary>
    /// ECMA-376 <c>a:satMod</c>: multiplies the HSL saturation, preserving hue and
    /// lightness. (Scaling channels toward gray would shift lightness and distort
    /// saturated accent colors, so a real HSL round-trip is required.)
    /// </summary>
    private static (byte R, byte G, byte B) ApplySatMod((byte R, byte G, byte B) color, double factor)
    {
        var (h, s, l) = RgbToHsl(color);
        s = Math.Clamp(s * factor, 0.0, 1.0);
        return HslToRgb(h, s, l);
    }

    private static (double H, double S, double L) RgbToHsl((byte R, byte G, byte B) color)
    {
        double r = color.R / 255.0, g = color.G / 255.0, b = color.B / 255.0;
        var max = Math.Max(r, Math.Max(g, b));
        var min = Math.Min(r, Math.Min(g, b));
        var l = (max + min) / 2;
        var d = max - min;
        if (d == 0)
            return (0, 0, l);

        var s = l <= 0.5 ? d / (max + min) : d / (2 - max - min);
        double h;
        if (max == r) h = ((g - b) / d) % 6;
        else if (max == g) h = (b - r) / d + 2;
        else h = (r - g) / d + 4;
        return ((h * 60 + 360) % 360, s, l);
    }

    private static (byte R, byte G, byte B) HslToRgb(double h, double s, double l)
    {
        var c = (1 - Math.Abs(2 * l - 1)) * s;
        var x = c * (1 - Math.Abs((h / 60) % 2 - 1));
        var m = l - c / 2;
        double r, g, b;
        if (h < 60) (r, g, b) = (c, x, 0);
        else if (h < 120) (r, g, b) = (x, c, 0);
        else if (h < 180) (r, g, b) = (0, c, x);
        else if (h < 240) (r, g, b) = (0, x, c);
        else if (h < 300) (r, g, b) = (x, 0, c);
        else (r, g, b) = (c, 0, x);
        return (Clamp((r + m) * 255), Clamp((g + m) * 255), Clamp((b + m) * 255));
    }

    private static (byte R, byte G, byte B) ApplyHslTransform(
        (byte R, byte G, byte B) color,
        Func<double, double>? hue,
        Func<double, double>? saturation,
        Func<double, double>? luminance)
    {
        var (h, s, l) = RgbToHsl(color);
        if (hue != null)
            h = (hue(h) % 360.0 + 360.0) % 360.0;
        if (saturation != null)
            s = Math.Clamp(saturation(s), 0.0, 1.0);
        if (luminance != null)
            l = Math.Clamp(luminance(l), 0.0, 1.0);
        return HslToRgb(h, s, l);
    }

    private static (byte R, byte G, byte B) ApplyThemeStyleShade(
        (byte R, byte G, byte B) color, double shade)
    {
        var (h, s, l) = RgbToHsl(color);
        l *= (1.0 + Math.Clamp(shade, 0.0, 1.0)) / 2.0;
        return HslToRgb(h, s, l);
    }

    private static (byte R, byte G, byte B) OffsetChannels((byte R, byte G, byte B) color, double offset)
    {
        return (Clamp(color.R + 255 * offset), Clamp(color.G + 255 * offset), Clamp(color.B + 255 * offset));
    }

    private static byte Clamp(double v)
    {
        return (byte)Math.Round(Math.Clamp(v, 0, 255), MidpointRounding.AwayFromZero);
    }

    private static byte ClampAlpha(double alpha)
        => (byte)Math.Clamp(alpha * 255.0, 0.0, 255.0);

    private static bool TryGetOoxmlVal(OpenXmlElement element, out int value)
    {
        var match = Regex.Match(element.OuterXml, @"\bval=""(-?\d+)""");
        return int.TryParse(
            match.Success ? match.Groups[1].Value : null,
            NumberStyles.Integer,
            CultureInfo.InvariantCulture,
            out value);
    }

    private static long? ReadLongAttribute(OpenXmlElement element, string attributeName)
    {
        var raw = ReadAttribute(element, attributeName);
        return long.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value)
            ? value
            : null;
    }

        // Raw XML attribute read via regex on OuterXml — SDK attribute parsing is
    // unreliable on missing attributes and on untyped (dsp:*) elements.
    private static string? ReadAttribute(OpenXmlElement element, string attributeName)
    {
        var match = Regex.Match(
            element.OuterXml,
            $@"\b{Regex.Escape(attributeName)}\s*=\s*""([^""]*)""",
            RegexOptions.CultureInvariant);
        return match.Success ? match.Groups[1].Value : null;
    }

    private static OpenXmlElement? GetChild(OpenXmlElement parent, string localName)
    {
        return parent.Elements()
            .FirstOrDefault(e => e.LocalName == localName && e.NamespaceUri == DrawingmlNs);
    }
}
