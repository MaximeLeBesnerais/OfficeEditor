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
        ["blockArc"] = (ShapeType.BlockArc, null),
        ["pie"] = (ShapeType.Pie, null),
        ["donut"] = (ShapeType.Donut, null),
    };

    /// <summary>
    /// Pre-computed polygon points (normalised 0..1) for non-rect/ellipse presets.
    /// </summary>
    private static readonly Dictionary<ShapeType, List<(double, double)>> PolygonPoints = new()
    {
        [ShapeType.Chevron] = new()
        {
            (0, 0.2), (0.55, 0), (0.55, 0.28), (1, 0.28),
            (1, 0.72), (0.55, 0.72), (0.55, 1), (0, 0.8)
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
        [ShapeType.BlockArc] = new()
        {
            (0.2, 0), (0.8, 0), (1, 0.2), (1, 0.8), (0.8, 1), (0.2, 1), (0, 0.8), (0, 0.2)
        },
        [ShapeType.Pie] = new()
        {
            (0.5, 0.5), (1, 0), (1, 0.5), (0.5, 1), (0, 0.5), (0, 0)
        },
        [ShapeType.Donut] = new()
        {
            (0.2, 0), (0.8, 0), (1, 0.2), (1, 0.8), (0.8, 1), (0.2, 1), (0, 0.8), (0, 0.2)
        },
    };

    /// <summary>
    /// Attempts to extract a shape element from a &lt;dsp:sp&gt; diagram shape.
    /// Returns null when the shape has no recognisable geometry or cannot be rendered.
    /// </summary>
    /// <param name="schemeColors">Theme-aware scheme colour map (scheme name → "#RRGGBB").
    /// May be empty; a static Office fallback is used when the map is missing an entry.</param>
    public static TypstElement? TryExtractShape(
        OpenXmlElement dspShape,
        double offX, double offY,
        double scaleX, double scaleY,
        double frameX, double frameY,
        double shapeW, double shapeH,
        IReadOnlyDictionary<string, string>? schemeColors = null)
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
            ShapeType.Rect => BuildRect(x, y, w, h, rotation, fillColor, strokeColor, strokeWidth, geometry.CornerRadius),

            ShapeType.Ellipse => new TypstElement
            {
                Type = "Shape",
                X = x, Y = y, Width = w, Height = h, Rotation = rotation,
                Shape = new TypstShapeElement
                {
                    ShapeType = "ellipse",
                    FillColor = fillColor ?? string.Empty,
                    StrokeColor = strokeColor ?? string.Empty,
                    StrokeWidth = strokeWidth
                }
            },

            _ => BuildPolygon(x, y, w, h, rotation, fillColor, strokeColor, strokeWidth, geometry.ShapeType)
        };
    }

    private static TypstElement BuildRect(double x, double y, double w, double h, double rotation,
        string? fillColor, string? strokeColor, double strokeWidth, double cornerRadius)
    {
        return new TypstElement
        {
            Type = "Shape",
            X = x, Y = y, Width = w, Height = h, Rotation = rotation,
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
        string? fillColor, string? strokeColor, double strokeWidth, ShapeType shapeType)
    {
        var points = PolygonPoints.TryGetValue(shapeType, out var pts)
            ? pts
            : new List<(double, double)> { (0, 0), (1, 0), (1, 1), (0, 1) };

        return new TypstElement
        {
            Type = "Shape",
            X = x, Y = y, Width = w, Height = h, Rotation = rotation,
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

            var radiusPt = EmuToPt(adjVal);
            var minSide = Math.Min(shapeW, shapeH);
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
                return val;
        }

        var schemeClr = GetChild(solidFill, "schemeClr", DrawingmlNs);
        if (schemeClr != null)
        {
            var val = ReadAttribute(schemeClr, "val");
            if (!string.IsNullOrEmpty(val))
            {
                return ResolveSchemeColor(val, schemeColors);
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
                return (val, strokeWidth);
        }

        var schemeClr = GetChild(solidFill, "schemeClr", DrawingmlNs);
        if (schemeClr != null)
        {
            var val = ReadAttribute(schemeClr, "val");
            if (!string.IsNullOrEmpty(val))
            {
                return (ResolveSchemeColor(val, schemeColors), strokeWidth);
            }
        }

        return (null, strokeWidth);
    }

    private static string? ResolveSchemeColor(string schemeName, IReadOnlyDictionary<string, string>? schemeColors)
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
            return themeRgb.StartsWith("#") ? themeRgb.Substring(1) : themeRgb;

        return StaticSchemeColorFallback(schemeName);
    }

    private static string? StaticSchemeColorFallback(string schemeName)
    {
        return schemeName switch
        {
            "accent1" => "4472C4",
            "accent2" => "ED7D31",
            "accent3" => "A5A5A5",
            "accent4" => "FFC000",
            "accent5" => "5B9BD5",
            "accent6" => "70AD47",
            "lt1" => "FFFFFF",
            "dk1" => "000000",
            "lt2" => "F2F2F2",
            "dk2" => "4472C4",
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
        Hexagon,
        BlockArc,
        Pie,
        Donut
    }
}
