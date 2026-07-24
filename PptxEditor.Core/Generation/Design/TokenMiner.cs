using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PptxEditor.Core.Generation.Model;
using PptxEditor.Core.Models;
using PptxEditor.Core.Services;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxEditor.Core.Generation.Design;

public static class TokenMiner
{
    public static DesignTokens Mine(string pptxPath)
    {
        if (!File.Exists(pptxPath))
            throw new FileNotFoundException("PPTX file not found.", pptxPath);

        using var document = PresentationDocument.Open(pptxPath, false);
        var brandProfile = new BrandProfileExtractor().Extract(document);

        var (shapeColors, shapeFonts) = MineShapeLevel(document);

        var palette = BuildPalette(brandProfile, shapeColors);
        var fonts = BuildFonts(brandProfile, shapeFonts);
        var metrics = BuildMetrics(brandProfile);

        return new DesignTokens
        {
            Palette = palette,
            Fonts = fonts,
            Shape = new ShapeTokens { CornerRadius = 8, CardStyle = CardStyle.Flat },
            Metrics = metrics
        };
    }

    private static (Dictionary<string, int> Colors, Dictionary<string, int> Fonts) MineShapeLevel(PresentationDocument document)
    {
        var paletteCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var fontCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        var presentationPart = document.PresentationPart;
        if (presentationPart?.Presentation?.SlideIdList == null)
            return (paletteCounts, fontCounts);

        foreach (var slideId in presentationPart.Presentation.SlideIdList.ChildElements.OfType<SlideId>())
        {
            var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!);
            var shapeTree = slidePart.Slide?.CommonSlideData?.ShapeTree;
            if (shapeTree == null)
                continue;

            CountColorsInElements(shapeTree.ChildElements, paletteCounts, fontCounts);
        }

        return (paletteCounts, fontCounts);
    }

    private static void CountColorsInElements(OpenXmlElementList elements, Dictionary<string, int> palette, Dictionary<string, int> fonts)
    {
        foreach (var element in elements)
        {
            switch (element)
            {
                case P.Shape shape:
                    CountShapeColors(shape, palette, fonts);
                    break;
                case P.GraphicFrame graphicFrame:
                    CountGraphicFrameColors(graphicFrame, palette, fonts);
                    break;
                case P.GroupShape groupShape:
                    CountColorsInElements(groupShape.ChildElements, palette, fonts);
                    break;
                case P.Picture picture:
                    break;
            }
        }
    }

    private static void CountShapeColors(P.Shape shape, Dictionary<string, int> palette, Dictionary<string, int> fonts)
    {
        if (shape.ShapeProperties != null)
        {
            CountFillColors(shape.ShapeProperties.Elements<Drawing.SolidFill>(), palette);
            CountStrokeColors(shape.ShapeProperties.Elements<Drawing.Outline>(), palette);
        }

        CountTextColors(shape.TextBody, palette, fonts);
    }

    private static void CountGraphicFrameColors(P.GraphicFrame graphicFrame, Dictionary<string, int> palette, Dictionary<string, int> fonts)
    {
        var table = graphicFrame.Descendants<DocumentFormat.OpenXml.Drawing.Table>().FirstOrDefault();
        if (table != null)
        {
            foreach (var gridCol in table.Descendants<Drawing.GridColumn>())
            {
                if (gridCol.Width?.Value != null)
                {
                }
            }

            foreach (var row in table.Elements<Drawing.TableRow>())
            {
                foreach (var cell in row.Elements<Drawing.TableCell>())
                {
                    var cellProps = cell.TableCellProperties;
                    if (cellProps != null)
                    {
                        CountFillColors(cellProps.Elements<Drawing.SolidFill>(), palette);
                    }

                    foreach (var paragraph in cell.Elements<Drawing.Paragraph>())
                    {
                        foreach (var run in paragraph.Elements<Drawing.Run>())
                        {
                            CountRunText(run, palette, fonts);
                        }
                    }
                }
            }

            return;
        }

        foreach (var paragraph in graphicFrame.Descendants<Drawing.Paragraph>())
        {
            foreach (var run in paragraph.Elements<Drawing.Run>())
            {
                CountRunText(run, palette, fonts);
            }
        }
    }

    private static void CountFillColors(IEnumerable<Drawing.SolidFill> fills, Dictionary<string, int> palette)
    {
        foreach (var fill in fills)
        {
            if (fill.RgbColorModelHex?.Val?.Value is { } hex)
                IncrementKey(palette, $"#{hex}");

            if (fill.SchemeColor != null && TryReadSchemeColor(fill.SchemeColor, out var scheme))
                IncrementKey(palette, $"%{scheme}%");
        }
    }

    private static void CountStrokeColors(IEnumerable<Drawing.Outline> outlines, Dictionary<string, int> palette)
    {
        foreach (var outline in outlines)
        {
            var fill = outline.Elements<Drawing.SolidFill>().FirstOrDefault();
            if (fill == null)
                continue;

            if (fill.RgbColorModelHex?.Val?.Value is { } hex)
                IncrementKey(palette, $"#{hex}");

            if (fill.SchemeColor != null && TryReadSchemeColor(fill.SchemeColor, out var scheme))
                IncrementKey(palette, $"%{scheme}%");
        }
    }

    private static void CountTextColors(OpenXmlElement? textBody, Dictionary<string, int> palette, Dictionary<string, int> fonts)
    {
        if (textBody == null)
            return;

        foreach (var paragraph in textBody.Elements<Drawing.Paragraph>())
        {
            foreach (var run in paragraph.Elements<Drawing.Run>())
            {
                CountRunText(run, palette, fonts);
            }
        }
    }

    private static void CountRunText(Drawing.Run run, Dictionary<string, int> palette, Dictionary<string, int> fonts)
    {
        var rpr = run.RunProperties;
        if (rpr == null)
            return;

        foreach (var fill in rpr.Elements<Drawing.SolidFill>())
        {
            if (fill.RgbColorModelHex?.Val?.Value is { } hex)
                IncrementKey(palette, $"#{hex}");

            if (fill.SchemeColor != null && TryReadSchemeColor(fill.SchemeColor, out var scheme))
                IncrementKey(palette, $"%{scheme}%");
        }

        for (var i = 0; i < rpr.ChildElements.Count; i++)
        {
            var child = rpr.ChildElements[i];
            if (child is Drawing.LatinFont latin && !string.IsNullOrWhiteSpace(latin.Typeface?.Value))
                IncrementKey(fonts, latin.Typeface!.Value!);
        }
    }

    private static bool TryReadSchemeColor(Drawing.SchemeColor schemeColor, out string value)
    {
        value = string.Empty;
        if (schemeColor.Val?.Value != null)
        {
            value = schemeColor.Val.Value.ToString();
            return !string.IsNullOrWhiteSpace(value);
        }

        var match = Regex.Match(schemeColor.OuterXml, @"\bval\s*=\s*""([^""]*)""");
        if (match.Success)
        {
            value = match.Groups[1].Value;
            return !string.IsNullOrWhiteSpace(value);
        }

        return false;
    }

    private static void IncrementKey(Dictionary<string, int> dict, string key)
    {
        dict.TryGetValue(key, out var count);
        dict[key] = count + 1;
    }

    private static IReadOnlyDictionary<string, string> BuildPalette(BrandProfile brandProfile, Dictionary<string, int> shapeColors)
    {
        var theme = brandProfile.ThemeColors;

        var ink = theme.GetValueOrDefault("dk1", "#000000");
        var paper = theme.GetValueOrDefault("lt1", "#FFFFFF");
        var muted = theme.GetValueOrDefault("lt2", theme.GetValueOrDefault("dk2", "#888888"));

        var accentCandidates = new[] { "accent1", "accent2", "accent3", "accent4", "accent5", "accent6" };
        var primary = ResolvePrimaryColor(theme, accentCandidates, shapeColors);
        var accent = ResolveAccentColor(theme, primary, accentCandidates, shapeColors);

        return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["primary"] = primary,
            ["accent"] = accent,
            ["ink"] = ink,
            ["paper"] = paper,
            ["muted"] = muted
        };
    }

    private static string ResolvePrimaryColor(
        IReadOnlyDictionary<string, string> theme,
        string[] accentCandidates,
        Dictionary<string, int> shapeColors)
    {
        string? best = null;
        var bestCount = -1;

        foreach (var slot in accentCandidates)
        {
            if (!theme.TryGetValue(slot, out var hex))
                continue;

            var freq = shapeColors.GetValueOrDefault($"%{slot}%", 0);
            freq += shapeColors.GetValueOrDefault(hex, 0);

            if (freq > bestCount)
            {
                bestCount = freq;
                best = hex;
            }
        }

        return best ?? theme.GetValueOrDefault("accent1", "#000000");
    }

    private static string ResolveAccentColor(
        IReadOnlyDictionary<string, string> theme,
        string primary,
        string[] accentCandidates,
        Dictionary<string, int> shapeColors)
    {
        string? best = null;
        var bestCount = -1;

        foreach (var slot in accentCandidates)
        {
            if (!theme.TryGetValue(slot, out var hex))
                continue;

            if (string.Equals(hex, primary, StringComparison.OrdinalIgnoreCase))
                continue;

            var freq = shapeColors.GetValueOrDefault($"%{slot}%", 0);
            freq += shapeColors.GetValueOrDefault(hex, 0);

            if (freq > bestCount)
            {
                bestCount = freq;
                best = hex;
            }
        }

        return best ?? theme.GetValueOrDefault("accent2", "#000000");
    }

    private static FontTokens BuildFonts(BrandProfile brandProfile, Dictionary<string, int> shapeFonts)
    {
        var display = brandProfile.MajorFont ?? PickTopFont(shapeFonts);
        var body = brandProfile.MinorFont ?? display;

        return new FontTokens { Display = display, Body = body };
    }

    private static string? PickTopFont(Dictionary<string, int> fonts)
    {
        string? best = null;
        var bestCount = -1;

        foreach (var (family, count) in fonts)
        {
            if (count > bestCount)
            {
                bestCount = count;
                best = family;
            }
        }

        return best;
    }

    private static MetricTokens BuildMetrics(BrandProfile brandProfile)
    {
        return new MetricTokens
        {
            MarginPt = 48,
            GutterPt = 18,
            TitleSizePt = brandProfile.Title?.FontSizePt ?? 32,
            BodySizePt = brandProfile.Body?.FontSizePt is { } bodySize and > 0 ? bodySize : 14
        };
    }
}
