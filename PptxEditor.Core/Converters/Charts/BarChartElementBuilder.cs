using PptxEditor.Core.Models;

namespace PptxEditor.Core.Converters.Charts;

/// <summary>
/// Decomposes a clustered bar/column <see cref="ChartModel"/> into existing Typst
/// primitives — rects for bars, gridlines, axis lines and legend swatches; text elements
/// for category, tick, value and legend labels — so no new element types or emission code
/// are needed. Layout happens once here, in slide points, inside the graphicFrame rect.
/// </summary>
public static class BarChartElementBuilder
{
    /// <summary>Default Office series palette, used when a series has no explicit fill.</summary>
    private static readonly string[] DefaultPalette =
        ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"];

    /// <summary>Average glyph width as a fraction of the font size (band estimation).</summary>
    private const double CharWidthFactor = 0.62;

    private const double LabelLineHeight = 1.5;
    private const double AxisLineWidth = 0.9;
    private const double GridlineWidth = 0.7;

    /// <summary>
    /// Builds the chart's elements inside the frame rect (<paramref name="x"/>,
    /// <paramref name="y"/>, <paramref name="width"/>, <paramref name="height"/>) in
    /// points. Returns an empty list for degenerate input (no data, no room).
    /// </summary>
    public static List<TypstElement> Build(ChartModel chart, double x, double y, double width, double height)
    {
        var elements = new List<TypstElement>();
        if (chart.Series.Count == 0 || chart.Categories.Count == 0 || width < 20 || height < 20)
            return elements;

        var isStacked = chart.Grouping is BarGrouping.Stacked or BarGrouping.PercentStacked;
        var values = isStacked
            ? chart.Categories.Select((_, i) =>
                chart.Grouping == BarGrouping.PercentStacked
                    ? 100.0
                    : chart.Series.Select(s => i < s.Values.Count ? s.Values[i] ?? 0 : 0)
                        .Where(v => v > 0).Sum()).ToList()
            : chart.Series.SelectMany(s => s.Values).OfType<double>().ToList();
        if (values.Count == 0)
            return elements;

        var isHorizontal = chart.Direction == BarDirection.Bar;
        var labelStyle = chart.DataLabels;
        var showValues = chart.Series.Any(s => (s.DataLabels ?? labelStyle)?.ShowValue == true);

        // ---- value axis scale ------------------------------------------------------
        var valAxis = chart.ValueAxis;
        var (min, max, unit) = ChartAxisScale.Compute(
            values.Min(), values.Max(), showValues,
            valAxis?.Min, valAxis?.Max, valAxis?.MajorUnit);
        var ticks = ChartAxisScale.Ticks(min, max, unit);
        var numberFormat = valAxis?.NumberFormat;

        // ---- typography ------------------------------------------------------------
        var catFont = AxisFont(chart.CategoryAxis);
        var valFont = AxisFont(valAxis);
        var dataFont = LabelFont(labelStyle);
        var legendFont = chart.Legend != null
            ? new LabelStyle(chart.Legend.FontSize ?? 9.0, chart.Legend.Color ?? "#404040", chart.Legend.FontFamily ?? "Arial")
            : null;

        // ---- bands around the plot area ---------------------------------------------
        var catBand = chart.CategoryAxis?.Deleted == true
            ? 0.0
            : isHorizontal
                ? chart.Categories.Max(c => EstimateWidth(c, catFont.Size)) + 8
                : catFont.Size * LabelLineHeight + 4;
        var valBand = valAxis?.Deleted == true
            ? 0.0
            : isHorizontal
                ? valFont.Size * LabelLineHeight + 4
                : ticks.Max(t => EstimateWidth(ChartAxisScale.FormatValue(t, numberFormat), valFont.Size)) + 8;

        var legend = chart.Legend;
        var legendVertical = legend != null && legend.Position is ChartLegendPosition.Right or ChartLegendPosition.Left;
        var legendBand = 0.0;
        if (legend != null && legendFont != null)
        {
            legendBand = legendVertical
                ? chart.Series.Max(s => EstimateWidth(s.Name, legendFont.Size)) + legendFont.Size + 18
                : legendFont.Size * 1.9;
        }

        double plotX = x, plotY = y, plotW = width, plotH = height;
        if (isHorizontal)
        {
            plotX += catBand;
            plotW -= catBand;
            plotH -= valBand;
        }
        else
        {
            plotX += valBand;
            plotW -= valBand;
            plotH -= catBand;
        }

        double legendX = 0, legendY = 0;
        if (legend != null)
        {
            switch (legend.Position)
            {
                case ChartLegendPosition.Left:
                    legendX = plotX - (isHorizontal ? catBand : valBand);
                    legendY = plotY;
                    plotX += legendBand;
                    plotW -= legendBand;
                    break;
                case ChartLegendPosition.Top:
                    legendX = plotX;
                    legendY = plotY;
                    plotY += legendBand;
                    plotH -= legendBand;
                    break;
                case ChartLegendPosition.Bottom:
                    legendX = plotX;
                    plotH -= legendBand;
                    // The category band is already below the plot.  Put a bottom
                    // legend after that band instead of on top of category labels.
                    legendY = plotY + plotH + catBand;
                    break;
                default: // Right / TopRight
                    plotW -= legendBand;
                    legendX = plotX + plotW + 8;
                    legendY = plotY;
                    break;
            }
        }

        if (plotW < 10 || plotH < 10)
            return elements;

        // ---- coordinate mapping -----------------------------------------------------
        double Span() => max - min;
        double XOf(double v) => plotX + (v - min) / Span() * plotW;
        double YOf(double v) => plotY + (max - v) / Span() * plotH;

        // ---- gridlines + axis lines --------------------------------------------------
        var gridColor = valAxis?.Deleted == true ? null : valAxis?.GridlineColor;
        if (gridColor != null)
        {
            foreach (var tick in ticks)
            {
                if (Math.Abs(tick - min) < unit * 1e-6)
                    continue; // the baseline axis line covers the min tick
                elements.Add(isHorizontal
                    ? Rect(XOf(tick) - GridlineWidth / 2, plotY, GridlineWidth, plotH, gridColor)
                    : Rect(plotX, YOf(tick) - GridlineWidth / 2, plotW, GridlineWidth, gridColor));
            }
        }

        var catLine = AxisLineColor(chart.CategoryAxis, catFont);
        var valLine = AxisLineColor(valAxis, valFont);
        if (chart.CategoryAxis?.Deleted != true)
        {
            // Category axis line runs along the value baseline.
            elements.Add(isHorizontal
                ? Rect(plotX - AxisLineWidth / 2, plotY, AxisLineWidth, plotH, catLine)
                : Rect(plotX, YOf(Math.Max(min, 0.0)) - AxisLineWidth / 2, plotW, AxisLineWidth, catLine));
        }
        if (valAxis?.Deleted != true)
        {
            elements.Add(isHorizontal
                ? Rect(plotX, plotY + plotH - AxisLineWidth / 2, plotW, AxisLineWidth, valLine)
                : Rect(plotX - AxisLineWidth / 2, plotY, AxisLineWidth, plotH, valLine));
        }

        // ---- tick + category labels ---------------------------------------------------
        if (valAxis?.Deleted != true)
        {
            foreach (var tick in ticks)
            {
                var text = ChartAxisScale.FormatValue(tick, numberFormat);
                var h = valFont.Size * LabelLineHeight;
                elements.Add(isHorizontal
                    ? Text(text, XOf(tick) - 25, plotY + plotH + 2, 50, h, valFont, "center")
                    : Text(text, plotX - valBand, YOf(tick) - h / 2, valBand - 4, h, valFont, "right"));
            }
        }

        var categoryCount = chart.Categories.Count;
        var seriesCount = chart.Series.Count;
        var slot = (isHorizontal ? plotH : plotW) / categoryCount;
        var gap = chart.GapWidthPercent / 100.0;
        var overlap = Math.Clamp(chart.OverlapPercent / 100.0, -1.0, 1.0);
        // slot = n·bar − (n−1)·overlap·bar + gap·bar  ⇒  bar thickness per category slot.
        var barsInCluster = isStacked ? 1 : seriesCount;
        var barThickness = slot / (barsInCluster - (barsInCluster - 1) * overlap + gap);
        var clusterThickness = barThickness * (barsInCluster - (barsInCluster - 1) * overlap);

        if (chart.CategoryAxis?.Deleted != true)
        {
            for (var i = 0; i < categoryCount; i++)
            {
                var h = catFont.Size * LabelLineHeight;
                if (isHorizontal)
                {
                    var slotTop = plotY + plotH - (i + 1) * slot;
                    elements.Add(Text(chart.Categories[i], plotX - catBand, slotTop + slot / 2 - h / 2,
                        catBand - 4, h, catFont, "right"));
                }
                else
                {
                    elements.Add(Text(chart.Categories[i], plotX + i * slot, plotY + plotH + 2,
                        slot, h, catFont, "center", verticalAlign: "top"));
                }
            }
        }

        // ---- bars + value labels -------------------------------------------------------
        for (var i = 0; i < categoryCount; i++)
        {
            for (var j = 0; j < seriesCount; j++)
            {
                var series = chart.Series[j];
                if (i >= series.Values.Count || series.Values[i] is not { } value)
                    continue;

                var color = series.FillColor ?? DefaultPalette[j % DefaultPalette.Length];
                var plottedValue = value;
                var stackStart = 0.0;
                if (isStacked)
                {
                    var categoryTotal = chart.Series.Select(s => i < s.Values.Count ? s.Values[i] ?? 0 : 0)
                        .Where(v => v > 0).Sum();
                    if (chart.Grouping == BarGrouping.PercentStacked)
                        plottedValue = categoryTotal > 0 ? value / categoryTotal * 100.0 : 0;
                    stackStart = chart.Series.Take(j)
                        .Select(s => i < s.Values.Count ? s.Values[i] ?? 0 : 0)
                        .Where(v => v > 0)
                        .Sum();
                    if (chart.Grouping == BarGrouping.PercentStacked && categoryTotal > 0)
                        stackStart = stackStart / categoryTotal * 100.0;
                }
                double barX, barY, barW, barH;
                if (isHorizontal)
                {
                    var slotTop = plotY + plotH - (i + 1) * slot;
                    // Horizontal bars plot categories and series bottom-to-top.
                    var clusterBottom = slotTop + slot - (slot - clusterThickness) / 2;
                    barY = clusterBottom - (isStacked ? 0 : j * (1 - overlap)) * barThickness - barThickness;
                    barH = barThickness;
                    var startValue = stackStart;
                    var endValue = stackStart + plottedValue;
                    barX = Math.Min(XOf(startValue), XOf(endValue));
                    barW = Math.Abs(XOf(endValue) - XOf(startValue));
                }
                else
                {
                    var slotLeft = plotX + i * slot;
                    barW = barThickness;
                    var barSlot = slotLeft + (slot - clusterThickness) / 2;
                    barX = isStacked ? barSlot : barSlot + j * barThickness * (1 - overlap);
                    var endValue = stackStart + plottedValue;
                    barY = Math.Min(YOf(stackStart), YOf(endValue));
                    barH = Math.Abs(YOf(endValue) - YOf(stackStart));
                }

                if (barW > 0.2 && barH > 0.2)
                    elements.Add(Rect(barX, barY, barW, barH, color));

                var seriesLabels = series.DataLabels ?? labelStyle;
                if (seriesLabels?.ShowValue == true)
                {
                    var style = LabelFont(seriesLabels);
                    var text = ChartAxisScale.FormatValue(value, numberFormat);
                    var h = style.Size * LabelLineHeight;
                    var w = EstimateWidth(text, style.Size) + 6;
                    if (isHorizontal)
                    {
                        var labelX = value >= 0 ? XOf(value) + 2 : XOf(value) - w - 2;
                        elements.Add(Text(text, labelX, barY + barH / 2 - h / 2, w, h, style,
                            value >= 0 ? "left" : "right"));
                    }
                    else
                    {
                        var labelY = plottedValue >= 0 ? barY - h - 1 : barY + barH + 1;
                        elements.Add(Text(text, barX + barW / 2 - w / 2, labelY, w, h, style, "center"));
                    }
                }
            }
        }

        // ---- legend ---------------------------------------------------------------------
        if (legend != null && legendFont != null)
        {
            // Horizontal bar charts reverse the legend to match the bottom-to-top axis.
            var order = isHorizontal
                ? Enumerable.Range(0, seriesCount).Reverse()
                : Enumerable.Range(0, seriesCount);
            var entryHeight = legendFont.Size * 1.9;
            var swatch = legendFont.Size * 0.8;

            if (legendVertical)
            {
                var totalHeight = seriesCount * entryHeight;
                var entryY = legendY + (plotH - totalHeight) / 2;
                foreach (var j in order)
                {
                    var color = chart.Series[j].FillColor ?? DefaultPalette[j % DefaultPalette.Length];
                    elements.Add(Rect(legendX, entryY + (entryHeight - swatch) / 2, swatch, swatch, color));
                    elements.Add(Text(chart.Series[j].Name, legendX + swatch + 5, entryY,
                        legendBand - swatch - 13, entryHeight, legendFont, "left"));
                    entryY += entryHeight;
                }
            }
            else
            {
                var entryX = legendX;
                foreach (var j in order)
                {
                    var color = chart.Series[j].FillColor ?? DefaultPalette[j % DefaultPalette.Length];
                    var w = EstimateWidth(chart.Series[j].Name, legendFont.Size) + swatch + 14;
                    elements.Add(Rect(entryX, legendY + (entryHeight - swatch) / 2, swatch, swatch, color));
                    elements.Add(Text(chart.Series[j].Name, entryX + swatch + 5, legendY,
                        w - swatch - 5, entryHeight, legendFont, "left"));
                    entryX += w;
                }
            }
        }

        return elements;
    }

    private sealed record LabelStyle(double Size, string Color, string Font);

    private static LabelStyle AxisFont(ChartAxis? axis)
        => new(axis?.LabelFontSize ?? 9.0, axis?.LabelColor ?? "#595959", axis?.LabelFontFamily ?? "Arial");

    private static LabelStyle LabelFont(ChartDataLabels? labels)
        => new(labels?.FontSize ?? 9.0, labels?.Color ?? "#404040", labels?.FontFamily ?? "Arial");

    private static string AxisLineColor(ChartAxis? axis, LabelStyle font)
        => axis?.LineColor ?? axis?.LabelColor ?? "#BFBFBF";

    private static double EstimateWidth(string text, double fontSize)
        => text.Length * fontSize * CharWidthFactor;

    private static TypstElement Rect(double x, double y, double w, double h, string fill)
        => new()
        {
            Type = "Shape",
            X = x,
            Y = y,
            Width = w,
            Height = h,
            Shape = new TypstShapeElement { ShapeType = "rect", FillColor = fill }
        };

    private static TypstElement Text(string content, double x, double y, double w, double h,
        LabelStyle style, string align, string verticalAlign = "center")
    {
        var formatting = new TypstTextFormatting
        {
            FontSize = style.Size,
            Color = style.Color,
            FontFamily = style.Font,
            Align = align
        };
        return new TypstElement
        {
            Type = "Text",
            X = x,
            Y = y,
            Width = w,
            Height = h,
            Text = new TypstTextElement
            {
                Paragraphs =
                [
                    new TypstParagraph
                    {
                        Content = content,
                        Runs = [new TypstTextRun { Content = content, Formatting = formatting }],
                        Formatting = formatting
                    }
                ],
                VerticalAlign = verticalAlign,
                ParagraphCount = 1
            }
        };
    }
}
