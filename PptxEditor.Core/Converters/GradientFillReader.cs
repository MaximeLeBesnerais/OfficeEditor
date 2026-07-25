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
    /// SmartArt drawing parts  use gradients as the primary shape fill.
    /// </summary>
    /// <param name="shapeProperties">Shape-properties element whose direct
    /// <c>a:gradFill</c> child is read (any element type — typed
    /// <c>p:spPr</c> or untyped <c>dsp:spPr</c>).</param>
    /// <param name="resolveSchemeColor">Scheme color name → <c>#RRGGBB</c>, or null
    /// when scheme colors cannot be resolved (stops referencing them then fail and
    /// the whole gradient is rejected, matching the previous behavior).</param>
    internal static TypstGradientFill? TryReadLinearGradient(
        OpenXmlElement shapeProperties, Func<string, string?>? resolveSchemeColor)
    {
        var gradFill = GetChild(shapeProperties, "gradFill");
        var stopList = gradFill == null ? null : GetChild(gradFill, "gsLst");
        if (stopList == null)
            return null;

        var stops = new List<TypstGradientStop>();
        foreach (var gs in stopList.Elements()
                     .Where(e => e.LocalName == "gs" && e.NamespaceUri == DrawingmlNs))
        {
            var color = ReadStopColor(gs, resolveSchemeColor);
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

    private static string? ReadStopColor(OpenXmlElement stop, Func<string, string?>? resolveSchemeColor)
    {
        var rgb = GetChild(stop, "srgbClr");
        var hex = rgb == null ? null : ReadAttribute(rgb, "val");
        if (!string.IsNullOrEmpty(hex))
        {
            // Unlike a solid fill, a fully transparent gradient stop is meaningful
            // (fade-out) — keep it as #RRGGBB00 instead of treating it as noFill.
            return ApplyColorModifiers(ParseHexColor(hex!), rgb!.ChildElements)
                ?? FormatHexColor(ParseHexColor(hex!), 0);
        }

        var schemeColor = GetChild(stop, "schemeClr");
        var schemeName = schemeColor == null ? null : ReadAttribute(schemeColor, "val");
        if (!string.IsNullOrEmpty(schemeName))
        {
            var resolved = resolveSchemeColor?.Invoke(schemeName!);
            if (!string.IsNullOrEmpty(resolved))
            {
                return ApplyColorModifiers(ParseHexColor(resolved!), schemeColor!.ChildElements)
                    ?? FormatHexColor(ParseHexColor(resolved!), 0);
            }
        }

        return null;
    }

    internal static (byte R, byte G, byte B) ParseHexColor(string hex)
    {
        hex = hex.TrimStart('#');
        if (hex.Length == 6)
        {
            return (
                (byte)int.Parse(hex.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
                (byte)int.Parse(hex.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
                (byte)int.Parse(hex.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture)
            );
        }
        return (0, 0, 0);
    }

    internal static string FormatHexColor((byte R, byte G, byte B) color, byte alpha = 255)
    {
        if (alpha == 255)
            return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
        return $"#{color.R:X2}{color.G:X2}{color.B:X2}{alpha:X2}";
    }

    /// <summary>
    /// Applies DrawingML color transforms (<c>a:tint</c>, <c>a:alpha</c>) to a base
    /// color. Returns null when the result is fully transparent (visually identical
    /// to <c>a:noFill</c> for solid fills) — gradient callers substitute an explicit
    /// <c>#RRGGBB00</c> stop instead.
    /// </summary>
    internal static string? ApplyColorModifiers((byte R, byte G, byte B) color, OpenXmlElementList modifiers)
    {
        byte alpha = 255;
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
                        alpha = ApplyAlpha(alphaVal);
                    break;
                case "lumMod":
                    if (TryGetOoxmlVal(mod, out var lumModVal))
                        color = ScaleChannels(color, lumModVal / 100000.0);
                    break;
                case "lumOff":
                    if (TryGetOoxmlVal(mod, out var lumOffVal))
                        color = OffsetChannels(color, lumOffVal / 100000.0);
                    break;
            }
        }
        // Fully transparent solid fill is visually identical to a:noFill — report no
        // color so callers treat the fill/stroke/text color as absent.
        return alpha == 0 ? null : FormatHexColor(color, alpha);
    }

    private static (byte R, byte G, byte B) ApplyTint((byte R, byte G, byte B) color, int tint)
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

    private static (byte R, byte G, byte B) OffsetChannels((byte R, byte G, byte B) color, double offset)
    {
        return (Clamp(color.R + 255 * offset), Clamp(color.G + 255 * offset), Clamp(color.B + 255 * offset));
    }

    private static byte Clamp(double v)
    {
        return (byte)Math.Round(Math.Clamp(v, 0, 255), MidpointRounding.AwayFromZero);
    }

    private static byte ApplyAlpha(int alpha)
    {
        return (byte)(alpha / 100000.0 * 255);
    }

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

    // Raw XML attribute read per AGENTS.pptx.md rule 1 — SDK attribute parsing is
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
