using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Generation.Design;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Design;

/// <summary>
/// Builds the baseline <c>w:style</c> elements for <see cref="BaselineStyleKind"/>. All values are
/// deterministic and derive from the resolved design: heading/title/callout use the display font and
/// the palette's <c>primary</c> color when defined, body-derived styles use the body font, and
/// nothing forces a font the design did not specify (so document defaults stay authoritative).
/// </summary>
internal static class BaselineStyles
{
    /// <summary>
    /// Creates the baseline style element for <paramref name="kind"/>. <paramref name="bodyStyleId"/>
    /// is the (existing or generated) body style every paragraph style derives from. The Body style
    /// itself is created standalone — it is the root of the baseline chain.
    /// </summary>
    public static Style Create(BaselineStyleKind kind, ResolvedDesign design, string styleId, string styleName, string bodyStyleId)
    {
        if (kind == BaselineStyleKind.Body)
        {
            return CreateBody(design, styleId, styleName);
        }
        if (kind == BaselineStyleKind.Table)
        {
            return CreateTable(styleId, styleName);
        }

        var style = new Style { Type = StyleValues.Paragraph, StyleId = styleId };
        style.Append(new StyleName { Val = styleName });
        style.Append(new BasedOn { Val = bodyStyleId });
        style.Append(new NextParagraphStyle { Val = bodyStyleId });

        switch (kind)
        {
            case BaselineStyleKind.Title:
                // w:pPr must precede w:rPr in CT_Style.
                style.Append(ParagraphProperties(beforePt: null, afterPt: 15, keepNext: false));
                style.Append(RunProperties(design, DocxDesignDefaults.TitleFontSizePt, bold: null, color: Primary(design), display: true));
                break;
            case BaselineStyleKind.Callout:
                // w:pPr child order: pBdr → shd → spacing → ind.
                style.Append(new StyleParagraphProperties(
                    new ParagraphBorders(new LeftBorder { Val = BorderValues.Single, Color = ToColor(Primary(design)) ?? "808080", Size = 16 }),
                    new Shading { Fill = "F2F2F2" },
                    new SpacingBetweenLines { Before = "120", After = "120" },
                    new Indentation { Left = "288", Right = "288" }));
                style.Append(RunProperties(design, sizePt: null, bold: null, color: null, display: false));
                break;
            case BaselineStyleKind.TableHeader:
                style.Append(RunProperties(design, sizePt: null, bold: true, color: Primary(design) is null ? null : "#FFFFFF", display: false));
                break;
            case BaselineStyleKind.Code:
                style.Append(new StyleParagraphProperties(new Shading { Fill = "F5F5F5" }));
                style.Append(new StyleRunProperties(
                    new RunFonts { Ascii = DocxDesignDefaults.CodeFont, HighAnsi = DocxDesignDefaults.CodeFont, ComplexScript = DocxDesignDefaults.CodeFont },
                    new FontSize { Val = DocxFormattingHelpers.HalfPoints(DocxDesignDefaults.CodeFontSizePt) },
                    new FontSizeComplexScript { Val = DocxFormattingHelpers.HalfPoints(DocxDesignDefaults.CodeFontSizePt) }));
                break;
            default:
                // Heading1..6.
                var level = kind.HeadingLevel();
                var sizePt = DocxDesignDefaults.HeadingFontSizePt(level);
                var beforeTwentieths = Math.Max(240 - (level - 1) * 30, 120);
                style.Append(ParagraphProperties(beforePt: beforeTwentieths / 20.0, afterPt: 3, keepNext: true, outlineLevel: level - 1));
                style.Append(RunProperties(design, sizePt, bold: true, color: Primary(design), display: true));
                break;
        }
        return style;
    }

    private static Style CreateBody(ResolvedDesign design, string styleId, string styleName)
    {
        var style = new Style { Type = StyleValues.Paragraph, StyleId = styleId };
        style.Append(new StyleName { Val = styleName });
        if (design.BodyFontFamily is { } font)
        {
            style.Append(new StyleRunProperties(
                new RunFonts { Ascii = font, HighAnsi = font, ComplexScript = font },
                new FontSize { Val = DocxFormattingHelpers.HalfPoints(DocxDesignDefaults.BodyFontSizePt) },
                new FontSizeComplexScript { Val = DocxFormattingHelpers.HalfPoints(DocxDesignDefaults.BodyFontSizePt) }));
        }
        return style;
    }

    private static Style CreateTable(string styleId, string styleName)
    {
        var borders = new TableBorders(
            new TopBorder { Val = BorderValues.Single, Size = 4, Color = "auto" },
            new LeftBorder { Val = BorderValues.Single, Size = 4, Color = "auto" },
            new BottomBorder { Val = BorderValues.Single, Size = 4, Color = "auto" },
            new RightBorder { Val = BorderValues.Single, Size = 4, Color = "auto" },
            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "auto" },
            new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = "auto" });
        return new Style
        {
            Type = StyleValues.Table,
            StyleId = styleId,
            StyleName = new StyleName { Val = styleName },
            StyleTableProperties = new StyleTableProperties(new TableProperties(borders))
        };
    }

    private static StyleRunProperties RunProperties(ResolvedDesign design, double? sizePt, bool? bold, string? color, bool display)
    {
        var font = display ? design.DisplayFontFamily : design.BodyFontFamily;
        var properties = new StyleRunProperties();
        if (font is not null)
        {
            properties.Append(new RunFonts { Ascii = font, HighAnsi = font, ComplexScript = font });
        }
        if (bold == true)
        {
            properties.Append(new Bold());
            properties.Append(new BoldComplexScript());
        }
        if (color is not null && ToColor(color) is { } hex)
        {
            properties.Append(new Color { Val = hex });
        }
        if (sizePt is { } size)
        {
            var halfPoints = DocxFormattingHelpers.HalfPoints(size);
            properties.Append(new FontSize { Val = halfPoints });
            properties.Append(new FontSizeComplexScript { Val = halfPoints });
        }
        return properties;
    }

    private static StyleParagraphProperties ParagraphProperties(double? beforePt, double? afterPt, bool keepNext, int? outlineLevel = null)
    {
        var properties = new StyleParagraphProperties();
        if (keepNext)
        {
            properties.Append(new KeepNext());
        }
        if (beforePt is not null || afterPt is not null)
        {
            properties.Append(new SpacingBetweenLines
            {
                Before = beforePt is { } b ? DocxFormattingHelpers.Twentieths(b) : null,
                After = afterPt is { } a ? DocxFormattingHelpers.Twentieths(a) : null
            });
        }
        if (outlineLevel is { } level)
        {
            properties.Append(new OutlineLevel { Val = level });
        }
        return properties;
    }

    /// <summary>The palette's "primary" color as #RRGGBB, when the design defines it.</summary>
    private static string? Primary(ResolvedDesign design) =>
        design.Palette.TryGetValue("primary", out var hex) ? hex : null;

    /// <summary>Normalizes a #RRGGBB to the OOXML color form (no leading '#').</summary>
    private static string? ToColor(string? hex) => DocxFormattingHelpers.ToOoxmlColor(hex);
}
