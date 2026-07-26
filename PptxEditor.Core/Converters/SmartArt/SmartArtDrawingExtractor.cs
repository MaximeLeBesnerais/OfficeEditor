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
        ["line"] = (ShapeType.Rect, null),
        ["downArrow"] = (ShapeType.DownArrow, null),
        ["upArrow"] = (ShapeType.UpArrow, null),
        ["leftArrow"] = (ShapeType.LeftArrow, null),
        ["leftRightArrow"] = (ShapeType.LeftRightArrow, null),
        ["upDownArrow"] = (ShapeType.UpDownArrow, null),
        ["trapezoid"] = (ShapeType.Trapezoid, null),
        ["circularArrow"] = (ShapeType.CircularArrow, null),
        ["leftCircularArrow"] = (ShapeType.LeftCircularArrow, null),
        ["gear6"] = (ShapeType.Gear6, null),
        ["gear9"] = (ShapeType.Gear9, null),
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
        [ShapeType.DownArrow] = new()
        {
            (0.25, 0), (0.75, 0), (0.75, 0.6), (1, 0.6),
            (0.5, 1), (0, 0.6), (0.25, 0.6)
        },
        [ShapeType.UpArrow] = new()
        {
            (0.25, 1), (0.75, 1), (0.75, 0.4), (1, 0.4),
            (0.5, 0), (0, 0.4), (0.25, 0.4)
        },
        [ShapeType.LeftArrow] = new()
        {
            (1, 0.25), (0.4, 0.25), (0.4, 0), (0, 0.5),
            (0.4, 1), (0.4, 0.75), (1, 0.75)
        },
        [ShapeType.LeftRightArrow] = new()
        {
            (0, 0.5), (0.4, 0), (0.4, 0.25), (0.6, 0.25), (0.6, 0),
            (1, 0.5), (0.6, 1), (0.6, 0.75), (0.4, 0.75), (0.4, 1)
        },
        [ShapeType.UpDownArrow] = new()
        {
            (0.5, 0), (1, 0.4), (0.75, 0.4), (0.75, 0.6), (1, 0.6),
            (0.5, 1), (0, 0.6), (0.25, 0.6), (0.25, 0.4), (0, 0.4)
        },
        [ShapeType.Trapezoid] = new()
        {
            (0.25, 0), (0.75, 0), (1, 1), (0, 1)
        },
    };

    /// <summary>
    /// Presets whose faithful outline requires elliptical arcs. Their polygon
    /// points are computed per shape by <see cref="SmartArtPresetGeometry"/>
    /// (honoring a:avLst adjustments) instead of the static table above.
    /// </summary>
    private static bool IsArcBasedPreset(ShapeType shapeType)
    {
        return shapeType is ShapeType.CircularArrow or ShapeType.LeftCircularArrow
            or ShapeType.Gear6 or ShapeType.Gear9;
    }

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
    /// Maps the drawing-space bounding box onto the graphic frame with a uniform
    /// (aspect-preserving) scale, centred within the frame. Returns the draw scale
    /// plus the adjusted frame origin expected by <see cref="TryExtractShape"/>
    /// (final = off + (frame + shapeOff) · scale).
    ///
    /// Calibration (sales_acceleration_deck slide 15): PowerPoint's cached
    /// dsp:drawing is authored in frame coordinates — the drawing bbox width equals
    /// the frame width (359.86pt vs 360pt) and the content is vertically centred
    /// with symmetric 24pt internal margins — so the previous per-axis bbox stretch
    /// over-sized shapes (~20% too tall) whenever the content did not span the full
    /// frame. Uniform fit + centring reproduces PowerPoint's effective layout box
    /// for cached drawings, and degrades gracefully (fit + centre) when the frame
    /// was resized after the cache was generated.
    /// </summary>
    internal static (double ScaleX, double ScaleY, double FrameX, double FrameY) ComputeFrameFit(
        (double MinX, double MinY, double Width, double Height) bounds,
        (double X, double Y, double Width, double Height) frame)
    {
        var scale = Math.Min(frame.Width / bounds.Width, frame.Height / bounds.Height);
        var centerOffsetX = (frame.Width - bounds.Width * scale) / 2;
        var centerOffsetY = (frame.Height - bounds.Height * scale) / 2;

        return (
            scale,
            scale,
            (frame.X + centerOffsetX) / scale - bounds.MinX,
            (frame.Y + centerOffsetY) / scale - bounds.MinY);
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
        // Many SmartArt drawing parts carry their colors as a:gradFill on the shape
        //  rather than a:solidFill — read gradients through the same
        // shared reader the main slide-shape path uses.
        var fillGradient = fillColor == null
            ? GradientFillReader.TryReadLinearGradient(
                spPr, name => ResolveSchemeColor(name, schemeColors))
            : null;
        var (strokeColor, strokeWidth) = ReadStroke(spPr, schemeColors);
        // Cached drawing shapes carry their styling inline; a missing or fill-less
        // a:ln means "no border" in PowerPoint. Flag it so the Typst emitter writes
        // an explicit stroke: none instead of inheriting Typst's 1pt black default
        // (which drew a visible black box around text-container shapes).
        var noStroke = string.IsNullOrEmpty(strokeColor) || strokeWidth <= 0;

        var x = offX + (frameX + geometry.OffsetX) * scaleX;
        var y = offY + (frameY + geometry.OffsetY) * scaleY;
        var w = geometry.Width * scaleX;
        var h = geometry.Height * scaleY;
        var rotation = geometry.Rotation ?? 0.0;

        return geometry.ShapeType switch
        {
            ShapeType.Rect => BuildRect(x, y, w, h, rotation, fillColor, fillGradient, strokeColor, strokeWidth, noStroke, geometry.CornerRadius, modelId),

            ShapeType.Ellipse => new TypstElement
            {
                Type = "Shape",
                X = x, Y = y, Width = w, Height = h, Rotation = rotation,
                ModelId = modelId,
                Shape = new TypstShapeElement
                {
                    ShapeType = "ellipse",
                    FillColor = fillColor ?? string.Empty,
                    FillGradient = fillGradient,
                    StrokeColor = strokeColor ?? string.Empty,
                    StrokeWidth = strokeWidth,
                    NoStroke = noStroke
                }
            },

            _ => BuildPolygon(x, y, w, h, rotation, fillColor, fillGradient, strokeColor, strokeWidth, noStroke, geometry, modelId)
        };
    }

    private static TypstElement BuildRect(double x, double y, double w, double h, double rotation,
        string? fillColor, TypstGradientFill? fillGradient, string? strokeColor, double strokeWidth, bool noStroke, double cornerRadius, string? modelId)
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
                FillGradient = fillGradient,
                StrokeColor = strokeColor ?? string.Empty,
                StrokeWidth = strokeWidth,
                NoStroke = noStroke,
                CornerRadius = cornerRadius
            }
        };
    }

    private static TypstElement BuildPolygon(double x, double y, double w, double h, double rotation,
        string? fillColor, TypstGradientFill? fillGradient, string? strokeColor, double strokeWidth, bool noStroke, DiagramGeometry geometry, string? modelId)
    {
        // Arc-based presets (gears, circular arrows) evaluate their ECMA-376
        // preset definition per shape so a:avLst adjustments are honored;
        // everything else uses the static normalized polygon table.
        List<(double, double)>? points = null;
        if (IsArcBasedPreset(geometry.ShapeType))
        {
            points = SmartArtPresetGeometry.TryBuildNormalizedPoints(
                geometry.PrstName, geometry.Width, geometry.Height, geometry.Adjustments);
        }

        points ??= PolygonPoints.TryGetValue(geometry.ShapeType, out var pts)
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
                FillGradient = fillGradient,
                StrokeColor = strokeColor ?? string.Empty,
                StrokeWidth = strokeWidth,
                NoStroke = noStroke,
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
        // unflipped (documented in docs/SMARTART-REPORT.md §4.4 known limitations).

        var prstGeom = GetChild(spPr, "prstGeom", DrawingmlNs);
        if (prstGeom == null) return null;

        var prstValue = ReadAttribute(prstGeom, "prst");
        if (string.IsNullOrEmpty(prstValue)) return null;

        if (!PresetGeometryMap.TryGetValue(prstValue, out var mapping))
            return null;

        var cornerRadius = 0.0;
        if (mapping.Type == ShapeType.Rect)
        {
            cornerRadius = ReadCornerRadius(prstGeom, prstValue, shapeW, shapeH);
        }

        return new DiagramGeometry
        {
            ShapeType = mapping.Type,
            PrstName = prstValue,
            Adjustments = ReadAdjustments(prstGeom),
            OffsetX = offsetX.Value,
            OffsetY = offsetY.Value,
            Width = width.Value,
            Height = height.Value,
            Rotation = rotationDeg,
            CornerRadius = cornerRadius
        };
    }

    /// <summary>
    /// Reads the literal adjustment values (<c>a:gd name="adjN" fmla="val V"/></c>)
    /// from a preset geometry's avLst. Non-literal formulas are ignored — cached
    /// diagram drawing parts always carry plain <c>val</c> adjustments.
    /// </summary>
    private static IReadOnlyDictionary<string, double> ReadAdjustments(OpenXmlElement prstGeom)
    {
        var avLst = GetChild(prstGeom, "avLst", DrawingmlNs);
        if (avLst == null) return new Dictionary<string, double>(StringComparer.Ordinal);

        var adjustments = new Dictionary<string, double>(StringComparer.Ordinal);
        foreach (var gd in avLst.Elements()
                     .Where(e => e.LocalName == "gd" && e.NamespaceUri == DrawingmlNs))
        {
            var name = ReadAttribute(gd, "name");
            var fmla = ReadAttribute(gd, "fmla");
            if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(fmla)) continue;

            var valPart = fmla.StartsWith("val ", StringComparison.Ordinal)
                ? fmla.Substring(4).Trim()
                : null;
            if (valPart == null) continue;

            if (double.TryParse(valPart, NumberStyles.Float, CultureInfo.InvariantCulture, out var adjVal))
            {
                adjustments[name] = adjVal;
            }
        }

        return adjustments;
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

    /// <summary>
    /// ECMA-376 default corner-radius adjustment for the rounded-rectangle family
    /// (<c>roundRect</c>, <c>round1Rect</c>, <c>round2SameRect</c>): 1/6 of
    /// min(w,h). Cached SmartArt drawings almost always carry an empty avLst, which
    /// must still render rounded — not collapse to a sharp rectangle.
    /// </summary>
    private const long DefaultRoundedRectAdj = 16667;

    private static double ReadCornerRadius(OpenXmlElement prstGeom, string prstName, double shapeW, double shapeH)
    {
        var avLst = GetChild(prstGeom, "avLst", DrawingmlNs);
        var gdList = avLst == null
            ? new List<OpenXmlElement>()
            : avLst.Elements()
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

            return AdjToCornerRadius(adjVal, shapeW, shapeH);
        }

        // No explicit adj: the rounded-rectangle family falls back to the ECMA
        // default; every other rect-mapped preset (rect, line) stays sharp.
        var isRoundedRect = prstName is "roundRect" or "round1Rect" or "round2SameRect";
        return isRoundedRect ? AdjToCornerRadius(DefaultRoundedRectAdj, shapeW, shapeH) : 0;
    }

    private static double AdjToCornerRadius(long adjVal, double shapeW, double shapeH)
    {
        // adj is a fraction of the smaller shape dimension in 100000ths
        // (e.g. val 10000 = 10% of min(w,h)) — NOT an EMU value. Mirrors
        // PptxToTypstConverter.ExtractShapeCornerRadius.
        var minSide = Math.Min(shapeW, shapeH);
        var radiusPt = adjVal / 100000.0 * minSide;
        return Math.Min(radiusPt, minSide * 0.5);
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
                return (ApplyColorTransforms(NormalizeHexColor(val), srgbClr), strokeWidth);
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
    /// <c>a:lumOff</c>, <c>a:alpha</c>) in document order. <c>a:tint</c> delegates to
    /// <see cref="GradientFillReader.ApplyTint"/> (gamma-linear blend, matching
    /// PowerPoint's pale tints); the remaining transforms use per-channel arithmetic —
    /// a close approximation of the HSL-space spec (ECMA-376).
    /// <c>a:alpha</c> yields an 8-digit <c>#RRGGBBAA</c> hex (Typst <c>rgb()</c> accepts it).
    /// Saturation transforms are not applied (unused in diagram drawing-part solid fills).
    /// </summary>
    private static string ApplyColorTransforms(string hex, OpenXmlElement colorElement)
    {
        var hexValue = hex.TrimStart('#');
        if (hexValue.Length != 6)
            return NormalizeHexColor(hex);

        double r = Convert.ToInt32(hexValue[..2], 16);
        double g = Convert.ToInt32(hexValue[2..4], 16);
        double b = Convert.ToInt32(hexValue[4..6], 16);
        double alpha = 1.0;

        foreach (var child in colorElement.ChildElements)
        {
            var valAttr = ReadAttribute(child, "val");
            if (string.IsNullOrEmpty(valAttr) ||
                !double.TryParse(valAttr, NumberStyles.Float, CultureInfo.InvariantCulture, out var rawVal))
                continue;

            var f = rawVal / 100000.0;
            switch (child.LocalName)
            {
                case "tint": // mix toward white in linear light (matches PowerPoint's pale tints)
                    var tinted = GradientFillReader.ApplyTint(
                        ((byte)ClampChannel(r), (byte)ClampChannel(g), (byte)ClampChannel(b)),
                        (int)Math.Round(rawVal));
                    r = tinted.R; g = tinted.G; b = tinted.B;
                    break;
                case "shade": // mix toward black
                case "lumMod": // luminance multiply (per-channel approximation)
                    r *= f; g *= f; b *= f;
                    break;
                case "lumOff": // luminance offset (per-channel approximation)
                    r += 255 * f; g += 255 * f; b += 255 * f;
                    break;
                case "alpha": // opacity multiplier
                    alpha = f;
                    break;
            }
        }

        var rgbOut = $"#{ClampChannel(r):X2}{ClampChannel(g):X2}{ClampChannel(b):X2}";
        if (alpha < 1.0)
        {
            rgbOut += $"{ClampChannel(alpha * 255):X2}";
        }
        return rgbOut;
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
        public string PrstName { get; init; } = string.Empty;
        public IReadOnlyDictionary<string, double> Adjustments { get; init; } =
            new Dictionary<string, double>(StringComparer.Ordinal);
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
        DownArrow,
        UpArrow,
        LeftArrow,
        LeftRightArrow,
        UpDownArrow,
        Trapezoid,
        CircularArrow,
        LeftCircularArrow,
        Gear6,
        Gear9
    }
}
