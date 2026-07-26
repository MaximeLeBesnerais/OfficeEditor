using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using PptxEditor.Core.Converters.Charts;
using PptxEditor.Core.Models;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace PptxEditor.Core.Converters;

/// <summary>
/// Chart graphicFrame conversion: resolves the <c>c:chart</c> relationship to its chart
/// part, parses it (see <see cref="ChartPartParser"/>) and decomposes clustered bar/column
/// charts into Typst primitives (see <see cref="BarChartElementBuilder"/>) and pie charts
/// into wedge polygons (see <see cref="PieChartElementBuilder"/>). Everything
/// else — line/stacked/doughnut charts, unresolvable parts, unparseable XML — keeps the
/// existing visible-placeholder fallback with an accurate warning.
/// </summary>
public sealed partial class PptxToTypstConverter
{
    private static readonly Regex ChartRelationshipIdPattern =
        new(@"\br:id\s*=\s*""([^""]*)""", RegexOptions.Compiled);

    private IEnumerable<TypstElement> ConvertChartGraphicFrame(
        SlidePart slidePart, Drawing.GraphicData graphicData,
        (double X, double Y, double Width, double Height) position,
        double offX, double offY, double scaleX, double scaleY,
        StyleResolver styleResolver, string name)
    {
        var (elements, warning) = BuildChartElements(
            slidePart, graphicData, position, offX, offY, scaleX, scaleY, styleResolver, name);

        if (warning != null)
            AddSlideWarning(warning);
        foreach (var element in elements)
            yield return element;
    }

    private (List<TypstElement> Elements, string? Warning) BuildChartElements(
        SlidePart slidePart, Drawing.GraphicData graphicData,
        (double X, double Y, double Width, double Height) position,
        double offX, double offY, double scaleX, double scaleY,
        StyleResolver styleResolver, string name)
    {
        // The c:chart child carries the relationship id; regex on OuterXml per repo
        // convention for unreliable OOXML attribute access.
        var relIdMatch = ChartRelationshipIdPattern.Match(graphicData.OuterXml);
        var chartXml = relIdMatch.Success
            ? TryReadChartPartXml(slidePart, relIdMatch.Groups[1].Value)
            : null;

        if (chartXml == null)
        {
            return (Placeholder(position, offX, offY, scaleX, scaleY),
                $"Chart '{name}' is not supported and was replaced by a placeholder.");
        }

        var model = ChartPartParser.Parse(chartXml, styleResolver.ResolveSchemeColor);
        if (model == null)
        {
            return (Placeholder(position, offX, offY, scaleX, scaleY),
                $"Chart '{name}' could not be parsed and was replaced by a placeholder.");
        }

        if (model.Kind == ChartKind.Bar && model.Grouping == BarGrouping.Clustered)
        {
            var elements = BarChartElementBuilder.Build(model,
                offX + position.X * scaleX, offY + position.Y * scaleY,
                position.Width * scaleX, position.Height * scaleY);
            if (elements.Count > 0)
                return (elements, null);

            return (Placeholder(position, offX, offY, scaleX, scaleY),
                $"Chart '{name}' has no plottable data and was replaced by a placeholder.");
        }

        if (model.Kind == ChartKind.Pie)
        {
            var elements = PieChartElementBuilder.Build(model,
                offX + position.X * scaleX, offY + position.Y * scaleY,
                position.Width * scaleX, position.Height * scaleY);
            if (elements.Count > 0)
                return (elements, null);

            return (Placeholder(position, offX, offY, scaleX, scaleY),
                $"Chart '{name}' has no plottable data and was replaced by a placeholder.");
        }

        var description = DescribeUnsupportedChart(model);
        return (Placeholder(position, offX, offY, scaleX, scaleY),
            $"Chart '{name}' ({description}) is not supported yet and was replaced by a placeholder.");
    }

    private List<TypstElement> Placeholder(
        (double X, double Y, double Width, double Height) position,
        double offX, double offY, double scaleX, double scaleY)
        => CreateUnsupportedFramePlaceholders(position, offX, offY, scaleX, scaleY, "Chart").ToList();

    private static string DescribeUnsupportedChart(ChartModel model)
    {
        var kind = model.Kind switch
        {
            ChartKind.Line => "line chart",
            ChartKind.Pie => "pie chart",
            ChartKind.Bar => "bar chart",
            _ => "unsupported chart type"
        };
        return model.Grouping == BarGrouping.Clustered
            ? kind
            : $"{model.Grouping.ToString().ToLowerInvariant()} {kind}";
    }

    private static string? TryReadChartPartXml(SlidePart slidePart, string relationshipId)
    {
        try
        {
            var part = slidePart.GetPartById(relationshipId);
            using var stream = part.GetStream();
            using var reader = new StreamReader(stream);
            return reader.ReadToEnd();
        }
        catch (Exception ex) when (ex is ArgumentOutOfRangeException or KeyNotFoundException or InvalidOperationException or IOException)
        {
            return null;
        }
    }
}
