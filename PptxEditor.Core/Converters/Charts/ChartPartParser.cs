using System.Globalization;
using System.Xml.Linq;

namespace PptxEditor.Core.Converters.Charts;

/// <summary>
/// Parses a chart part (<c>c:chartSpace</c>) into a <see cref="ChartModel"/>. Only cached
/// values (<c>c:strCache</c>/<c>c:numCache</c> and literal variants) are read — cell
/// references (<c>c:f</c>) are never evaluated, matching how PowerPoint renders the part.
/// Parsing uses <see cref="XDocument"/> rather than the OpenXML SDK so synthetic chart XML
/// can be unit-tested without a package.
/// </summary>
public static class ChartPartParser
{
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// <summary>
    /// Parses the chart part XML. <paramref name="schemeColorResolver"/> maps scheme color
    /// names (e.g. "accent1") to <c>#RRGGBB</c>; typically
    /// <see cref="StyleResolver.ResolveSchemeColor"/>. Returns null when the XML is not a
    /// recognizable chart space.
    /// </summary>
    public static ChartModel? Parse(string chartXml, Func<string, string?>? schemeColorResolver = null)
    {
        if (string.IsNullOrWhiteSpace(chartXml))
            return null;

        XDocument doc;
        try
        {
            doc = XDocument.Parse(chartXml);
        }
        catch (System.Xml.XmlException)
        {
            return null;
        }

        var chart = doc.Root?.Element(C + "chart");
        var plotArea = chart?.Element(C + "plotArea");
        if (chart == null || plotArea == null)
            return null;

        // The chart-type element is the plotArea child ending in "Chart" (barChart, …).
        var chartTypeElement = plotArea.Elements()
            .FirstOrDefault(e => e.Name.LocalName.EndsWith("Chart", StringComparison.Ordinal));
        if (chartTypeElement == null)
            return null;

        var kind = chartTypeElement.Name.LocalName switch
        {
            "barChart" => ChartKind.Bar,
            "lineChart" => ChartKind.Line,
            "pieChart" or "doughnutChart" => ChartKind.Pie,
            _ => ChartKind.Other
        };

        var model = new ChartModel
        {
            Kind = kind,
            Direction = ParseBarDirection(chartTypeElement),
            Grouping = ParseGrouping(chartTypeElement),
            GapWidthPercent = ParseDoubleAttribute(chartTypeElement.Element(C + "gapWidth")) ?? 150.0,
            OverlapPercent = ParseDoubleAttribute(chartTypeElement.Element(C + "overlap")) ?? 0.0,
            DataLabels = ParseDataLabels(chartTypeElement.Element(C + "dLbls"), schemeColorResolver),
            Legend = ParseLegend(chart.Element(C + "legend"), schemeColorResolver),
            Title = chart.Element(C + "title")?.Descendants(A + "t").FirstOrDefault()?.Value,
            CategoryAxis = ParseCategoryAxis(plotArea.Element(C + "catAx"), schemeColorResolver),
            ValueAxis = ParseValueAxis(plotArea.Element(C + "valAx"), schemeColorResolver),
            FirstSliceAngleDegrees = ParseDoubleAttribute(chartTypeElement.Element(C + "firstSliceAng")) ?? 0.0,
            HoleSizePercent = chartTypeElement.Name.LocalName == "doughnutChart"
                ? ParseDoubleAttribute(chartTypeElement.Element(C + "holeSize")) ?? 50.0
                : null
        };

        var series = new List<ChartSeries>();
        IReadOnlyList<string>? categories = null;
        foreach (var ser in chartTypeElement.Elements(C + "ser"))
        {
            series.Add(ParseSeries(ser, schemeColorResolver));
            categories ??= ParseCategories(ser.Element(C + "cat"));
        }

        return new ChartModel
        {
            Kind = model.Kind,
            Direction = model.Direction,
            Grouping = model.Grouping,
            GapWidthPercent = model.GapWidthPercent,
            OverlapPercent = model.OverlapPercent,
            DataLabels = model.DataLabels,
            Legend = model.Legend,
            Title = model.Title,
            CategoryAxis = model.CategoryAxis,
            ValueAxis = model.ValueAxis,
            FirstSliceAngleDegrees = model.FirstSliceAngleDegrees,
            HoleSizePercent = model.HoleSizePercent,
            Series = series,
            Categories = categories ?? []
        };
    }

    private static BarDirection ParseBarDirection(XElement chartType)
        => Val(chartType.Element(C + "barDir")) == "bar" ? BarDirection.Bar : BarDirection.Column;

    private static BarGrouping ParseGrouping(XElement chartType)
        => Val(chartType.Element(C + "grouping")) switch
        {
            "stacked" => BarGrouping.Stacked,
            "percentStacked" => BarGrouping.PercentStacked,
            _ => BarGrouping.Clustered
        };

    private static ChartSeries ParseSeries(XElement ser, Func<string, string?>? schemeColorResolver)
    {
        var namePoints = GetCachedPoints(ser.Element(C + "tx"));
        var values = GetCachedPoints(ser.Element(C + "val"))
            .Select(p => (p.Index, Value: ParseNumber(p.Text)))
            .ToList();

        var valueCount = values.Count > 0 ? values.Max(v => v.Index) + 1 : 0;
        var valueArray = new double?[valueCount];
        foreach (var (index, value) in values)
        {
            if (index >= 0 && index < valueCount)
                valueArray[index] = value;
        }

        // Pie/doughnut per-data-point styling (c:dPt): fills are indexed by point;
        // the slice outline is taken from the first dPt that declares one.
        var pointFills = new List<(int Index, string? Color)>();
        string? pointLineColor = null;
        double pointLineWidthPt = 0;
        foreach (var dPt in ser.Elements(C + "dPt"))
        {
            var index = ParseInt(dPt.Element(C + "idx")?.Attribute("val")?.Value) ?? pointFills.Count;
            var spPr = dPt.Element(C + "spPr");
            pointFills.Add((index, ParseSolidFill(spPr?.Element(A + "solidFill"), schemeColorResolver)));

            var ln = spPr?.Element(A + "ln");
            if (pointLineColor == null && ln != null)
            {
                pointLineColor = ParseSolidFill(ln.Element(A + "solidFill"), schemeColorResolver);
                if (pointLineColor != null)
                    pointLineWidthPt = (ParseDoubleAttribute(ln, "w") ?? 0) / 12700.0;
            }
        }

        var pointFillCount = pointFills.Count > 0 ? pointFills.Max(p => p.Index) + 1 : 0;
        var pointFillArray = new string?[pointFillCount];
        foreach (var (index, color) in pointFills)
        {
            if (index >= 0 && index < pointFillCount)
                pointFillArray[index] = color;
        }

        return new ChartSeries
        {
            Name = namePoints.Count > 0 ? namePoints[0].Text : string.Empty,
            FillColor = ParseSolidFill(ser.Element(C + "spPr")?.Element(A + "solidFill"), schemeColorResolver),
            Values = valueArray,
            DataLabels = ParseDataLabels(ser.Element(C + "dLbls"), schemeColorResolver),
            PointFillColors = pointFillArray,
            PointLineColor = pointLineColor,
            PointLineWidthPt = pointLineWidthPt
        };
    }

    private static IReadOnlyList<string> ParseCategories(XElement? cat)
    {
        if (cat == null)
            return [];

        var points = GetCachedPoints(cat);
        if (points.Count == 0)
            return [];

        var count = points.Max(p => p.Index) + 1;
        var categories = new string[count];
        foreach (var (index, text) in points)
        {
            if (index >= 0 && index < count)
                categories[index] = text;
        }
        return categories;
    }

    /// <summary>
    /// Reads cached points from a <c>c:tx</c>/<c:c:cat</c>/<c:c:val</c> container: the first
    /// <c>strCache</c>/<c>numCache</c>/<c>strLit</c>/<c>numLit</c> descendant, as
    /// (index, text) pairs in document order.
    /// </summary>
    private static List<(int Index, string Text)> GetCachedPoints(XElement? container)
    {
        var cache = container?
            .Descendants()
            .FirstOrDefault(e => e.Name.LocalName is "strCache" or "numCache" or "strLit" or "numLit");
        if (cache == null)
            return [];

        var points = new List<(int Index, string Text)>();
        foreach (var pt in cache.Elements(C + "pt"))
        {
            var index = ParseInt(pt.Attribute("idx")?.Value) ?? points.Count;
            points.Add((index, pt.Element(C + "v")?.Value ?? string.Empty));
        }
        return points;
    }

    private static ChartDataLabels? ParseDataLabels(XElement? dLbls, Func<string, string?>? schemeColorResolver)
    {
        if (dLbls == null)
            return null;

        // Defaults per OOXML: labels hidden unless explicitly enabled.
        var showValue = Val(dLbls.Element(C + "showVal")) == "1";
        var (fontSize, color, fontFamily) = ParseTextProperties(dLbls.Element(C + "txPr"), schemeColorResolver);
        return new ChartDataLabels
        {
            ShowValue = showValue,
            FontSize = fontSize,
            Color = color,
            FontFamily = fontFamily
        };
    }

    private static ChartLegend? ParseLegend(XElement? legend, Func<string, string?>? schemeColorResolver)
    {
        if (legend == null)
            return null;

        var (fontSize, color, fontFamily) = ParseTextProperties(legend.Element(C + "txPr"), schemeColorResolver);
        return new ChartLegend
        {
            Position = Val(legend.Element(C + "legendPos")) switch
            {
                "l" => ChartLegendPosition.Left,
                "t" => ChartLegendPosition.Top,
                "b" => ChartLegendPosition.Bottom,
                "tr" => ChartLegendPosition.TopRight,
                // PowerPoint treats a missing legendPos as right.
                _ => ChartLegendPosition.Right
            },
            FontSize = fontSize,
            Color = color,
            FontFamily = fontFamily
        };
    }

    private static ChartAxis? ParseCategoryAxis(XElement? catAx, Func<string, string?>? schemeColorResolver)
    {
        if (catAx == null)
            return null;

        var (fontSize, color, fontFamily) = ParseTextProperties(catAx.Element(C + "txPr"), schemeColorResolver);
        return new ChartAxis
        {
            Deleted = Val(catAx.Element(C + "delete")) == "1",
            LineColor = ParseAxisLineColor(catAx, schemeColorResolver),
            LabelFontSize = fontSize,
            LabelColor = color,
            LabelFontFamily = fontFamily
        };
    }

    private static ChartValueAxis? ParseValueAxis(XElement? valAx, Func<string, string?>? schemeColorResolver)
    {
        if (valAx == null)
            return null;

        var (fontSize, color, fontFamily) = ParseTextProperties(valAx.Element(C + "txPr"), schemeColorResolver);
        var scaling = valAx.Element(C + "scaling");
        return new ChartValueAxis
        {
            Deleted = Val(valAx.Element(C + "delete")) == "1",
            LineColor = ParseAxisLineColor(valAx, schemeColorResolver),
            LabelFontSize = fontSize,
            LabelColor = color,
            LabelFontFamily = fontFamily,
            Min = ParseDoubleAttribute(scaling?.Element(C + "min")),
            Max = ParseDoubleAttribute(scaling?.Element(C + "max")),
            MajorUnit = ParseDoubleAttribute(valAx.Element(C + "majorUnit")),
            GridlineColor = ParseAxisLineColor(valAx.Element(C + "majorGridlines"), schemeColorResolver),
            NumberFormat = valAx.Element(C + "numFmt")?.Attribute("formatCode")?.Value
        };
    }

    /// <summary>Axis / gridline stroke colour from <c>c:spPr/a:ln/a:solidFill</c>.</summary>
    private static string? ParseAxisLineColor(XElement? axisLike, Func<string, string?>? schemeColorResolver)
        => ParseSolidFill(
            axisLike?.Element(C + "spPr")?.Element(A + "ln")?.Element(A + "solidFill"),
            schemeColorResolver);

    /// <summary>Extracts size/colour/font from a <c>c:txPr</c> block's first <c>a:defRPr</c>.</summary>
    private static (double? FontSize, string? Color, string? FontFamily) ParseTextProperties(
        XElement? txPr, Func<string, string?>? schemeColorResolver)
    {
        var defRPr = txPr?.Descendants(A + "defRPr").FirstOrDefault();
        if (defRPr == null)
            return (null, null, null);

        // a:defRPr sz is in 1/100ths of a point.
        var fontSize = ParseInt(defRPr.Attribute("sz")?.Value) / 100.0;
        var color = ParseSolidFill(defRPr.Element(A + "solidFill"), schemeColorResolver);
        var fontFamily = defRPr.Element(A + "latin")?.Attribute("typeface")?.Value;
        return (fontSize, color, string.IsNullOrWhiteSpace(fontFamily) ? null : fontFamily);
    }

    /// <summary>
    /// Resolves an <c>a:solidFill</c> to <c>#RRGGBB</c>: srgbClr directly, schemeClr through
    /// the theme resolver. Colour transforms (lumMod/shade/…) are not applied (round 2).
    /// </summary>
    private static string? ParseSolidFill(XElement? solidFill, Func<string, string?>? schemeColorResolver)
    {
        if (solidFill == null)
            return null;

        var srgbElement = solidFill.Element(A + "srgbClr");
        var srgb = srgbElement?.Attribute("val")?.Value;
        if (!string.IsNullOrWhiteSpace(srgb))
            return ApplyColorModifiers(srgb, srgbElement!);

        var schemeElement = solidFill.Element(A + "schemeClr");
        var scheme = schemeElement?.Attribute("val")?.Value;
        if (!string.IsNullOrWhiteSpace(scheme) && schemeColorResolver != null)
        {
            var resolved = schemeColorResolver(scheme);
            return resolved == null ? null : ApplyColorModifiers(resolved, schemeElement!);
        }

        return null;
    }

    private static string ApplyColorModifiers(string color, XElement colorElement)
    {
        var hex = color.TrimStart('#');
        if (hex.Length < 6)
            return color;

        var r = Convert.ToInt32(hex.Substring(0, 2), 16);
        var g = Convert.ToInt32(hex.Substring(2, 2), 16);
        var b = Convert.ToInt32(hex.Substring(4, 2), 16);
        var alpha = 255;
        foreach (var modifier in colorElement.Elements())
        {
            if (!int.TryParse(modifier.Attribute("val")?.Value, NumberStyles.Integer,
                    CultureInfo.InvariantCulture, out var value))
                continue;
            switch (modifier.Name.LocalName)
            {
                case "lumMod":
                case "shade":
                    r = Math.Clamp((int)Math.Round(r * value / 100000.0), 0, 255);
                    g = Math.Clamp((int)Math.Round(g * value / 100000.0), 0, 255);
                    b = Math.Clamp((int)Math.Round(b * value / 100000.0), 0, 255);
                    break;
                case "lumOff":
                    r = Math.Clamp((int)Math.Round(r + 255 * value / 100000.0), 0, 255);
                    g = Math.Clamp((int)Math.Round(g + 255 * value / 100000.0), 0, 255);
                    b = Math.Clamp((int)Math.Round(b + 255 * value / 100000.0), 0, 255);
                    break;
                case "alpha":
                    alpha = Math.Clamp((int)Math.Round(value / 100000.0 * 255), 0, 255);
                    break;
            }
        }
        return alpha == 255
            ? $"#{r:X2}{g:X2}{b:X2}"
            : $"#{r:X2}{g:X2}{b:X2}{alpha:X2}";
    }

    private static string? Val(XElement? element) => element?.Attribute("val")?.Value;

    private static double? ParseDoubleAttribute(XElement? element)
        => ParseDoubleAttribute(element, "val");

    private static double? ParseDoubleAttribute(XElement? element, string attribute)
        => double.TryParse(element?.Attribute(attribute)?.Value, NumberStyles.Float,
            CultureInfo.InvariantCulture, out var value)
            ? value
            : null;

    private static int? ParseInt(string? text)
        => int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value)
            ? value
            : null;

    private static double? ParseNumber(string? text)
        => double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var value)
            ? value
            : null;
}
