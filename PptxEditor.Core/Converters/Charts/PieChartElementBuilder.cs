using PptxEditor.Core.Models;

namespace PptxEditor.Core.Converters.Charts;

/// <summary>
/// Decomposes a pie <see cref="ChartModel"/> into wedge polygons (arc-approximated,
/// one per slice) inside the graphicFrame rect, so pie charts render natively instead
/// of falling back to a placeholder frame. Angles follow OOXML semantics:
/// <c>c:firstSliceAng</c> degrees clockwise from 12 o'clock, slices laid out clockwise.
/// </summary>
public static class PieChartElementBuilder
{
    /// <summary>Default Office per-point palette for varyColors pies without dPt fills.</summary>
    private static readonly string[] DefaultPalette =
        ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"];

    /// <summary>
    /// PowerPoint's auto layout leaves a margin around the pie: the circle's diameter is
    /// ~78% of the smaller plot dimension (measured on the Space deck reference render).
    /// </summary>
    // Small chart frames in editable infographic pins/panels include title/legend
    // space; using the full frame made the pie cover its surrounding marker. Larger
    // standalone pie frames retain the established Office-like 78% sizing.
    private const double PieDiameterFactor = 0.78;

    /// <summary>Maximum arc degrees per polygon segment (smoothness of the wedge rim).</summary>
    private const double ArcStepDegrees = 4.0;

    /// <summary>
    /// Builds the pie's wedge elements inside the frame rect in points. Returns an empty
    /// list for degenerate input (no data, non-positive total, no room).
    /// </summary>
    public static List<TypstElement> Build(ChartModel chart, double x, double y, double width, double height)
    {
        var elements = new List<TypstElement>();
        if (chart.Series.Count == 0 || width < 10 || height < 10)
            return elements;

        var series = chart.Series[0];
        var slices = series.Values
            .Select((value, index) => (Value: value ?? 0, Index: index))
            .Where(s => s.Value > 0)
            .ToList();
        if (slices.Count == 0)
            return elements;

        var total = slices.Sum(s => s.Value);
        if (total <= 0)
            return elements;

        var titleHeight = string.IsNullOrWhiteSpace(chart.Title) ? 0.0 : 18.0;
        if (titleHeight > 0)
        {
            var formatting = new TypstTextFormatting { FontSize = 10, Color = "#FFFFFF", Align = "center" };
            elements.Add(new TypstElement
            {
                Type = "Text", X = x, Y = y, Width = width, Height = titleHeight,
                Text = new TypstTextElement
                {
                    Paragraphs = [new TypstParagraph
                    {
                        Content = chart.Title!,
                        Runs = [new TypstTextRun { Content = chart.Title!, Formatting = formatting }],
                        Formatting = formatting
                    }],
                    VerticalAlign = "center", ParagraphCount = 1
                }
            });
        }
        // Keep the narrowly supported legend path safe: doughnut legends are
        // emitted only for a bottom legend with enough category data. Other
        // positions retain the previous chart geometry.
        var legend = chart.HoleSizePercent is > 0 && chart.Legend is
            { Position: ChartLegendPosition.Bottom } && chart.Categories.Count > 0
            && width >= 80 && height >= 40
            ? chart.Legend
            : null;
        var legendFont = legend == null
            ? null
            : new LabelStyle(legend.FontSize ?? 9.0, legend.Color ?? "#404040", legend.FontFamily ?? "Arial");
        List<(int Index, string Label)> legendEntries = legend == null
            ? []
            : slices.Select(slice =>
                (slice.Index, Label: chart.Categories.ElementAtOrDefault(slice.Index) ?? $"item {slice.Index + 1}"))
                .ToList();
        var legendEntryHeight = legendFont?.Size * 1.9 ?? 0.0;
        var legendRows = legendFont == null
            ? []
            : PackLegendRows(legendEntries, width, legendFont.Size);
        var legendHeight = legendEntryHeight * legendRows.Count;
        var plotHeight = Math.Max(1.0, height - titleHeight - legendHeight);
        var diameterFactor = width < 200 && plotHeight < 200 ? 0.52 : PieDiameterFactor;
        var diameter = Math.Min(width, plotHeight) * diameterFactor;
        var radius = diameter / 2.0;
        var innerRadius = radius * Math.Clamp(chart.HoleSizePercent ?? 0, 0, 100) / 100.0;
        var cx = width / 2.0;
        var cy = titleHeight + plotHeight / 2.0;

        string Fill(int pointIndex)
            => pointIndex < series.PointFillColors.Count && series.PointFillColors[pointIndex] != null
                ? series.PointFillColors[pointIndex]!
                : series.FillColor ?? DefaultPalette[pointIndex % DefaultPalette.Length];

        // A single 100% slice is a plain ellipse unless this is a doughnut.  A
        // doughnut uses two even-odd contours so the hole remains transparent.
        if (slices.Count == 1 && innerRadius <= 0)
        {
            elements.Add(new TypstElement
            {
                Type = "Shape",
                Name = "Pie slice",
                X = x + cx - radius,
                Y = y + cy - radius,
                Width = diameter,
                Height = diameter,
                Shape = new TypstShapeElement
                {
                    ShapeType = "ellipse",
                    FillColor = Fill(slices[0].Index),
                    StrokeColor = series.PointLineColor ?? string.Empty,
                    StrokeWidth = series.PointLineWidthPt
                }
            });
            return elements;
        }

        var angle = chart.FirstSliceAngleDegrees;
        foreach (var (value, pointIndex) in slices)
        {
            var sweep = value / total * 360.0;
            var points = innerRadius > 0
                ? new List<(double X, double Y)>()
                : new List<(double X, double Y)> { Norm(cx, cy) };

            var steps = Math.Max(1, (int)Math.Ceiling(sweep / ArcStepDegrees));
            for (var i = 0; i <= steps; i++)
            {
                var a = (angle + sweep * i / steps) * Math.PI / 180.0;
                points.Add(Norm(cx + radius * Math.Sin(a), cy - radius * Math.Cos(a)));
            }

            List<List<(double X, double Y)>>? subpaths = null;
            if (innerRadius > 0)
            {
                var outer = points.ToList();
                var inner = new List<(double X, double Y)>();
                for (var i = steps; i >= 0; i--)
                {
                    var a = (angle + sweep * i / steps) * Math.PI / 180.0;
                    inner.Add(Norm(cx + innerRadius * Math.Sin(a), cy - innerRadius * Math.Cos(a)));
                }
                points = outer;
                subpaths = new List<List<(double X, double Y)>> { outer, inner };
            }

            elements.Add(new TypstElement
            {
                Type = "Shape",
                Name = $"Pie slice {pointIndex}",
                X = x,
                Y = y,
                Width = width,
                Height = height,
                Shape = new TypstShapeElement
                {
                    ShapeType = "polygon",
                    FillColor = Fill(pointIndex),
                    StrokeColor = series.PointLineColor ?? string.Empty,
                    StrokeWidth = series.PointLineWidthPt,
                    Points = points,
                    Subpaths = subpaths ?? []
                }
            });

            angle += sweep;
        }

        if (legendFont != null && legendRows.Count > 0)
        {
            var swatch = legendFont.Size * 0.8;
            var entryHeight = legendEntryHeight;
            for (var rowIndex = 0; rowIndex < legendRows.Count; rowIndex++)
            {
                var row = legendRows[rowIndex];
                var totalWidth = row.Sum(entry => entry.Width);
                var entryX = x + Math.Max(4, (width - totalWidth) / 2);
                var entryY = y + height - (legendRows.Count - rowIndex) * entryHeight;
                foreach (var entry in row)
                {
                    elements.Add(new TypstElement
                    {
                        Type = "Shape",
                        X = entryX,
                        Y = entryY + (entryHeight - swatch) / 2,
                        Width = swatch,
                        Height = swatch,
                        Shape = new TypstShapeElement { ShapeType = "rect", FillColor = Fill(entry.Index) }
                    });
                    elements.Add(Text(entry.Label, entryX + swatch + 4, entryY,
                        Math.Max(1, entry.Width - swatch - 4), entryHeight, legendFont, "left"));
                    entryX += entry.Width;
                }
            }
        }

        return elements;

        (double X, double Y) Norm(double px, double py) => (px / width, py / height);
    }

    private sealed record LabelStyle(double Size, string Color, string Font);

    private sealed record LegendEntry(int Index, string Label, double Width);

    private static List<List<LegendEntry>> PackLegendRows(
        IReadOnlyList<(int Index, string Label)> entries,
        double width,
        double fontSize)
    {
        var availableWidth = Math.Max(1.0, width - 8);
        var rows = new List<List<LegendEntry>>();
        var currentRow = new List<LegendEntry>();
        var currentWidth = 0.0;

        foreach (var entry in entries)
        {
            // Use this exact packed width for both row-height estimation and
            // emission. A long label occupies one clipped row rather than
            // causing the estimator to undercount rows via a global sum.
            var packedWidth = Math.Min(
                EstimateWidth(entry.Label, fontSize) + fontSize * 0.8 + 8,
                availableWidth);
            var packedEntry = new LegendEntry(entry.Index, entry.Label, packedWidth);
            if (currentRow.Count > 0 && currentWidth + packedWidth > availableWidth)
            {
                rows.Add(currentRow);
                currentRow = [];
                currentWidth = 0;
            }

            currentRow.Add(packedEntry);
            currentWidth += packedWidth;
        }

        if (currentRow.Count > 0)
            rows.Add(currentRow);
        return rows;
    }

    private static double EstimateWidth(string text, double fontSize)
        => text.Length * fontSize * 0.62;

    private static TypstElement Text(string content, double x, double y, double width, double height,
        LabelStyle style, string align)
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
            Width = width,
            Height = height,
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
                VerticalAlign = "center",
                ParagraphCount = 1
            }
        };
    }
}
