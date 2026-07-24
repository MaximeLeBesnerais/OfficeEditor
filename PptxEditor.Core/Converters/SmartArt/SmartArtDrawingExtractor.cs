using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using PptxEditor.Core.Models;

namespace PptxEditor.Core.Converters.SmartArt;

/// <summary>
/// Extracts shape geometry, fills, and strokes from SmartArt diagram drawing-part
/// shapes (<c>dsp:sp</c> elements).  Only reads pre-rendered shapes — does not
/// perform DataModel-driven layout.
/// </summary>
internal static class SmartArtDrawingExtractor
{
    private const string Diagram2006Ns = "http://schemas.openxmlformats.org/drawing/2006/diagram";
    private const string Diagram2008Ns = "http://schemas.microsoft.com/office/drawing/2008/diagram";
    private const string DrawingmlNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static readonly HashSet<string> DiagramNamespaces = new(StringComparer.Ordinal)
    {
        Diagram2006Ns, Diagram2008Ns
    };

    /// <summary>
    /// Preset geometry name → (output ShapeType, polygon-point list for "polygon" shapes).
    /// Point coordinates are normalised to [0,1] in each axis.
    /// </summary>
    private static readonly Dictionary<string, (ShapeType Type, List<(double, double)>? Points)> PresetGeometryMap = new(StringComparer.Ordinal)
    {
        ["rect"] = (ShapeType.Rect, null),
        ["roundRect"] = (ShapeType.Rect, null),
        ["round1Rect"] = (ShapeType.Rect, null),
        ["round2SameRect"] = (ShapeType.Rect, null),
        ["ellipse"] = (ShapeType.Ellipse, null),
        ["rightArrow"] = (ShapeType.RightArrow, null),
        ["chevron"] = (ShapeType.Chevron, null),
        ["triangle"] = (ShapeType.Triangle, null),
        ["diamond"] = (ShapeType.Diamond, null),
        ["pentagon"] = (ShapeType.Pentagon, null),
        ["hexagon"] = (ShapeType.Hexagon, null),
        // NOTE: blockArc, pie and donut were removed deliberately — the previous
        // polygon approximations were geometrically wrong (blockArc and donut shared
        // the same octagon points; pie used bounding-box corners outside the
        // ellipse). Unknown presets return null and degrade to the text fallback.
    };

    /// <summary>
    /// Pre-computed polygon points (normalised 0..1) for non-rect/ellipse presets.
    /// </summary>
    private static readonly Dictionary<ShapeType, List<(double, double)>> PolygonPoints = new()
    {
        [ShapeType.Chevron] = new()
        {
            // OOXML chevron (adj = 0.5): a rectangle with an arrow notch — 6 points.
            (0, 0), (0.5, 0), (1, 0.5), (0.5, 1), (0, 1), (0.5, 0.5)
        },
        [ShapeType.RightArrow] = new()
        {
            (0, 0.25), (0.6, 0.25), (0.6, 0), (1, 0.5),
            (0.6, 1), (0.6, 0.75), (0, 0.75)
        },
        [ShapeType.Triangle] = new()
        {
            (0.5, 0), (1, 1), (0, 1)
        },
        [ShapeType.Diamond] = new()
        {
            (0.5, 0), (1, 0.5), (0.5, 1), (0, 0.5)
        },
        [ShapeType.Pentagon] = new()
        {
            (0.5, 0), (1, 0.38), (0.81, 1), (0.19, 1), (0, 0.38)
        },
        [ShapeType.Hexagon] = new()
        {
            (0.25, 0), (0.75, 0), (1, 0.5), (0.75, 1), (0.25, 1), (0, 0.5)
        },
    };

    /// <summary>
    /// Computes the bounding box (in points) over the drawing-space geometry
    /// (<c>dsp:spPr/a:xfrm</c>) of the given diagram shapes. Returns null when no
    /// shape carries usable geometry. Used by the converter to normalise drawing
    /// space onto the graphic frame's extents.
    /// </summary>
    internal static (double MinX, double MinY, double Width, double Height)? ComputeBoundingBox(
        IEnumerable<OpenXmlElement> dspShapes)
    {
        double minX = double.MaxValue, minY = double.MaxValue;
        double maxX = double.MinValue, maxY = double.MinValue;
        var found = false;

        foreach (var shape in dspShapes)
        {
            var spPr = GetChild(shape, "spPr", DiagramNamespaces);
            var xfrm = spPr == null ? null : GetChild(spPr, "xfrm", DrawingmlNs);
            var off = xfrm == null ? null : GetChild(xfrm, "off", DrawingmlNs);
            var ext = xfrm == null ? null : GetChild(xfrm, "ext", DrawingmlNs);
            if (off == null || ext == null) continue;

            var x = ReadEmuAsPt(off, "x");
            var y = ReadEmuAsPt(off, "y");
            var cx = ReadEmuAsPt(ext, "cx");
            var cy = ReadEmuAsPt(ext, "cy");
            if (x == null || y == null || cx == null || cy == null) continue;

            found = true;
            minX = Math.Min(minX, x.Value);
            minY = Math.Min(minY, y.Value);
            maxX = Math.Max(maxX, x.Value + cx.Value);
            maxY = Math.Max(maxY, y.Value + cy.Value);
        }

        return found ? (minX, minY, maxX - minX, maxY - minY) : null;
    }

    /// <summary>
    /// Attempts to extract a shape element from a &lt;dsp:sp&gt; diagram shape.
    /// Returns null when the shape has no recognisable geometry or cannot be rendered.
    /// </summary>
    /// <param name="schemeColors">Theme-aware scheme colour map (scheme name → "#RRGGBB").
    /// May be empty; a static Office fallback is used when the map is missing an entry.</param>
    /// <param name="modelId">Optional <c>dsp:sp modelId</c> carried onto the emitted
    /// element so shapes can be joined back to data-model nodes.</param>
    public static TypstElement? TryExtractShape(
        OpenXmlElement dspShape,
        double offX, double offY,
        double scaleX, double scaleY,
        double frameX, double frameY,
        double shapeW, double shapeH,
        IReadOnlyDictionary<string, string>? schemeColors = null,
        string? modelId = null)
    {
        var spPr = GetChild(dspShape, "spPr", DiagramNamespaces);
        if (spPr == null) return null;

        var geometry = ReadGeometry(spPr, shapeW, shapeH);
        if (geometry == null) return null;

        var fillColor = ReadFillColor(spPr, schemeColors);
        var (strokeColor, strokeWidth) = ReadStroke(spPr, schemeColors);

        var x = offX + (frameX + geometry.OffsetX) * scaleX;
        var y = offY + (frameY + geometry.OffsetY) * scaleY;
        var w = geometry.Width * scaleX;
        var h = geometry.Height * scaleY;
        var rotation = geometry.Rotation ?? 0.0;

        return geometry.ShapeType switch
        {
            ShapeType.Rect => BuildRect(x, y, w, h, rotation, fillColor, strokeColor, strokeWidth, geometry.CornerRadius, modelId),

            ShapeType.Ellipse => new TypstElement
            {
                Type = "Shape",
                X = x, Y = y, Width = w, Height = h, Rotation = rotation,
                ModelId = modelId,
                Shape = new TypstShapeElement
                {
                    ShapeType = "ellipse",
                    FillColor = fillColor ?? string.Empty,
                    StrokeColor = strokeColor ?? string.Empty,
                    StrokeWidth = strokeWidth
                }
            },

            _ => BuildPolygon(x, y, w, h, rotation, fillColor, strokeColor, strokeWidth, geometry.ShapeType, modelId)
        };
    }

    private static TypstElement BuildRect(double x, double y, double w, double h, double rotation,
        string? fillColor, string? strokeColor, double strokeWidth, double cornerRadius, string? modelId)
    {
        return new TypstElement
        {
            Type = "Shape",
            X = x, Y = y, Width = w, Height = h, Rotation = rotation,
            ModelId = modelId,
            Shape = new TypstShapeElement
            {
                ShapeType = "rect",
                FillColor = fillColor ?? string.Empty,
                StrokeColor = strokeColor ?? string.Empty,
                StrokeWidth = strokeWidth,
                CornerRadius = cornerRadius
            }
        };
    }

    private static TypstElement BuildPolygon(double x, double y, double w, double h, double rotation,
        string? fillColor, string? strokeColor, double strokeWidth, ShapeType shapeType, string? modelId)
    {
        var points = PolygonPoints.TryGetValue(shapeType, out var pts)
            ? pts
            : new List<(double, double)> { (0, 0), (1, 0), (1, 1), (0, 1) };

        return new TypstElement
        {
            Type = "Shape",
            X = x, Y = y, Width = w, Height = h, Rotation = rotation,
            ModelId = modelId,
            Shape = new TypstShapeElement
            {
                ShapeType = "polygon",
                FillColor = fillColor ?? string.Empty,
                StrokeColor = strokeColor ?? string.Empty,
                StrokeWidth = strokeWidth,
                Points = points
            }
        };
    }

    private static DiagramGeometry? ReadGeometry(OpenXmlElement spPr, double shapeW, double shapeH)
    {
        var xfrm = GetChild(spPr, "xfrm", DrawingmlNs);
        if (xfrm == null) return null;

        var off = GetChild(xfrm, "off", DrawingmlNs);
        var ext = GetChild(xfrm, "ext", DrawingmlNs);
        if (off == null || ext == null) return null;

        var offsetX = ReadEmuAsPt(off, "x");
        var offsetY = ReadEmuAsPt(off, "y");
        var width = ReadEmuAsPt(ext, "cx");
        var height = ReadEmuAsPt(ext, "cy");

        if (offsetX == null || offsetY == null || width == null || height == null)
            return null;

        var rotationDeg = ReadRotation(xfrm);

        // NOTE: xfrm flipH/flipV are not applied — mirrored diagram shapes render
        // unflipped (documented in SMARTART-REPORT.md §4.4 known limitations).

        var prstGeom = GetChild(spPr, "prstGeom", DrawingmlNs);
        if (prstGeom == null) return null;

        var prstValue = ReadAttribute(prstGeom, "prst");
        if (string.IsNullOrEmpty(prstValue)) return null;

        if (!PresetGeometryMap.TryGetValue(prstValue, out var mapping))
            return null;

        var cornerRadius = 0.0;
        if (mapping.Type == ShapeType.Rect)
        {
            cornerRadius = ReadCornerRadius(prstGeom, shapeW, shapeH);
        }

        return new DiagramGeometry
        {
            ShapeType = mapping.Type,
            OffsetX = offsetX.Value,
            OffsetY = offsetY.Value,
            Width = width.Value,
            Height = height.Value,
            Rotation = rotationDeg,
            CornerRadius = cornerRadius
        };
    }

    private static double? ReadRotation(OpenXmlElement xfrm)
    {
        var rotStr = ReadAttribute(xfrm, "rot");
        if (string.IsNullOrEmpty(rotStr))
            return null;

        if (!long.TryParse(rotStr, NumberStyles.Integer, CultureInfo.InvariantCulture, out var rot60000))
            return null;

        return rot60000 / 60000.0;
    }

    private static double ReadCornerRadius(OpenXmlElement prstGeom, double shapeW, double shapeH)
    {
        var avLst = GetChild(prstGeom, "avLst", DrawingmlNs);
        if (avLst == null) return 0;

        var gdList = avLst.Elements()
            .Where(e => e.LocalName == "gd" && e.NamespaceUri == DrawingmlNs)
            .ToList();

        foreach (var gd in gdList)
        {
            var name = ReadAttribute(gd, "name");
            if (name != "adj") continue;

            var fmla = ReadAttribute(gd, "fmla");
            if (string.IsNullOrEmpty(fmla)) continue;

            var valPart = fmla.StartsWith("val ", StringComparison.Ordinal)
                ? fmla.Substring(4).Trim()
                : fmla.Trim();

            if (!long.TryParse(valPart, NumberStyles.Integer, CultureInfo.InvariantCulture, out var adjVal))
                continue;

            // adj is a fraction of the smaller shape dimension in 100000ths
            // (e.g. val 10000 = 10% of min(w,h)) — NOT an EMU value. Mirrors
            // PptxToTypstConverter.ExtractShapeCornerRadius.
            var minSide = Math.Min(shapeW, shapeH);
            var radiusPt = adjVal / 100000.0 * minSide;
            return Math.Min(radiusPt, minSide * 0.5);
        }

        return 0;
    }

    private static string? ReadFillColor(OpenXmlElement spPr, IReadOnlyDictionary<string, string>? schemeColors)
    {
        var solidFill = GetChild(spPr, "solidFill", DrawingmlNs);
        if (solidFill == null) return null;

        var srgbClr = GetChild(solidFill, "srgbClr", DrawingmlNs);
        if (srgbClr != null)
        {
            var val = ReadAttribute(srgbClr, "val");
            if (!string.IsNullOrEmpty(val))
                return ApplyColorTransforms(NormalizeHexColor(val), srgbClr);
        }

        var schemeClr = GetChild(solidFill, "schemeClr", DrawingmlNs);
        if (schemeClr != null)
        {
            var val = ReadAttribute(schemeClr, "val");
            if (!string.IsNullOrEmpty(val))
            {
                var resolved = ResolveSchemeColor(val, schemeColors);
                return resolved == null ? null : ApplyColorTransforms(resolved, schemeClr);
            }
        }

        return null;
    }

    private static (string? color, double width) ReadStroke(OpenXmlElement spPr, IReadOnlyDictionary<string, string>? schemeColors)
    {
        var ln = GetChild(spPr, "ln", DrawingmlNs);
        if (ln == null) return (null, 0);

        var wStr = ReadAttribute(ln, "w");
        var strokeWidth = 0.0;
        if (!string.IsNullOrEmpty(wStr) &&
            long.TryParse(wStr, NumberStyles.Integer, CultureInfo.InvariantCulture, out var wEmu))
        {
            strokeWidth = EmuToPt(wEmu);
        }

        var solidFill = GetChild(ln, "solidFill", DrawingmlNs);
        if (solidFill == null) return (null, strokeWidth);

        var srgbClr = GetChild(solidFill, "srgbClr", DrawingmlNs);
        if (srgbClr != null)
        {
            var val = ReadAttribute(srgbClr, "val");
            if (!string.IsNullOrEmpty(val))
                return (NormalizeHexColor(val), strokeWidth);
        }

        var schemeClr = GetChild(solidFill, "schemeClr", DrawingmlNs);
        if (schemeClr != null)
        {
            var val = ReadAttribute(schemeClr, "val");
            if (!string.IsNullOrEmpty(val))
            {
                var resolved = ResolveSchemeColor(val, schemeColors);
                return (resolved == null ? null : ApplyColorTransforms(resolved, schemeClr), strokeWidth);
            }
        }

        return (null, strokeWidth);
    }

    /// <summary>
    /// Applies OOXML colour transforms (<c>a:tint</c>, <c>a:shade</c>, <c>a:lumMod</c>,
    /// <c>a:lumOff</c>) in document order. Per-channel arithmetic — a close approximation of
    /// the HSL-space spec (ECMA-376) that matches PowerPoint for common tint/shade usage
    /// (e.g. SmartArt connector fills like accent1 + tint 60% = pale accent). Saturation and
    /// alpha transforms are not applied (rare in diagram drawing parts).
    /// </summary>
    private static string ApplyColorTransforms(string hex, OpenXmlElement colorElement)
    {
        var hexValue = hex.TrimStart('#');
        if (hexValue.Length != 6)
            return NormalizeHexColor(hex);

        double r = Convert.ToInt32(hexValue[..2], 16);
        double g = Convert.ToInt32(hexValue[2..4], 16);
        double b = Convert.ToInt32(hexValue[4..6], 16);

        foreach (var child in colorElement.ChildElements)
        {
            var valAttr = ReadAttribute(child, "val");
            if (string.IsNullOrEmpty(valAttr) ||
                !double.TryParse(valAttr, NumberStyles.Float, CultureInfo.InvariantCulture, out var rawVal))
                continue;

            var f = rawVal / 100000.0;
            switch (child.LocalName)
            {
                case "tint": // mix toward white
                    r = r * f + 255 * (1 - f); g = g * f + 255 * (1 - f); b = b * f + 255 * (1 - f);
                    break;
                case "shade": // mix toward black
                case "lumMod": // luminance multiply (per-channel approximation)
                    r *= f; g *= f; b *= f;
                    break;
                case "lumOff": // luminance offset (per-channel approximation)
                    r += 255 * f; g += 255 * f; b += 255 * f;
                    break;
            }
        }

        return $"#{ClampChannel(r):X2}{ClampChannel(g):X2}{ClampChannel(b):X2}";
    }

    private static int ClampChannel(double v)
    {
        return (int)Math.Round(Math.Clamp(v, 0, 255), MidpointRounding.AwayFromZero);
    }

    /// <summary>
    /// Resolves a scheme colour name (e.g. "accent1", "lt1") through the theme-aware
    /// map, falling back to the static Office theme defaults. Always returns the
    /// pipeline-canonical "#RRGGBB" form (used by the Typst emitters' rgb("…") calls).
    /// </summary>
    internal static string? ResolveSchemeColor(string schemeName, IReadOnlyDictionary<string, string>? schemeColors)
    {
        var canonical = schemeName switch
        {
            "bg1" => "lt1",
            "tx1" => "dk1",
            "bg2" => "lt2",
            "tx2" => "dk2",
            _ => schemeName
        };

        if (schemeColors != null && schemeColors.TryGetValue(canonical, out var themeRgb))
            return NormalizeHexColor(themeRgb);

        return StaticSchemeColorFallback(schemeName);
    }

    private static string NormalizeHexColor(string rgb)
    {
        return rgb.StartsWith('#') ? rgb : "#" + rgb;
    }

    private static string? StaticSchemeColorFallback(string schemeName)
    {
        return schemeName switch
        {
            "accent1" => "#4472C4",
            "accent2" => "#ED7D31",
            "accent3" => "#A5A5A5",
            "accent4" => "#FFC000",
            "accent5" => "#5B9BD5",
            "accent6" => "#70AD47",
            "lt1" => "#FFFFFF",
            "dk1" => "#000000",
            "lt2" => "#E7E6E6",
            "dk2" => "#44546A",
            _ => null
        };
    }

    private static double EmuToPt(long emu)
    {
        return emu / 12700.0;
    }

    private static double? ReadEmuAsPt(OpenXmlElement element, string attributeName)
    {
        var val = ReadAttribute(element, attributeName);
        if (string.IsNullOrEmpty(val)) return null;

        if (!long.TryParse(val, NumberStyles.Integer, CultureInfo.InvariantCulture, out var emu))
            return null;

        return EmuToPt(emu);
    }

    private static string? ReadAttribute(OpenXmlElement element, string attributeName)
    {
        var match = Regex.Match(
            element.OuterXml,
            $@"\b{Regex.Escape(attributeName)}\s*=\s*""([^""]*)""",
            RegexOptions.CultureInvariant);

        return match.Success ? match.Groups[1].Value : null;
    }

    private static OpenXmlElement? GetChild(OpenXmlElement parent, string localName, HashSet<string> namespaces)
    {
        return parent.Elements()
            .FirstOrDefault(e => e.LocalName == localName && namespaces.Contains(e.NamespaceUri));
    }

    private static OpenXmlElement? GetChild(OpenXmlElement parent, string localName, string namespaceUri)
    {
        return parent.Elements()
            .FirstOrDefault(e => e.LocalName == localName && e.NamespaceUri == namespaceUri);
    }

    private sealed class DiagramGeometry
    {
        public ShapeType ShapeType { get; init; }
        public double OffsetX { get; init; }
        public double OffsetY { get; init; }
        public double Width { get; init; }
        public double Height { get; init; }
        public double? Rotation { get; init; }
        public double CornerRadius { get; init; }
    }

    private enum ShapeType
    {
        Rect,
        Ellipse,
        RightArrow,
        Chevron,
        Triangle,
        Diamond,
        Pentagon,
        Hexagon
    }
}
