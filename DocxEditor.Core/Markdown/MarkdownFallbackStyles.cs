using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxEditor.Core.Markdown;

/// <summary>
/// Builds the generated fallback <see cref="Style"/> definitions that
/// <see cref="MarkdownStyleResolver"/> appends when a requested style reference is missing or
/// resolves to the wrong kind. Formatting is derived from the semantic markdown element
/// (heading1..heading6, paragraph, blockquote, codeBlock, codeInline, hyperlink,
/// definitionTerm, definitionDescription, footnoteText, list, tableHeader, table, …) so a
/// standalone markdown→DOCX render stays readable even against a style-less template; a
/// well-known Word style reference (Heading1..6, Normal, Quote, Code, CodeChar, Hyperlink,
/// TableGrid, …) is used as a fallback when no element key is available. Values are
/// deterministic and the generated style never mutates existing template styles.
/// </summary>
internal static class MarkdownFallbackStyles
{
    private const string BodyFont = "Calibri";
    private const string MonoFont = "Consolas";

    /// <summary>
    /// Creates the fallback style for <paramref name="styleId"/>/<paramref name="styleName"/> of
    /// the expected <paramref name="kind"/>. <paramref name="element"/> is the semantic markdown
    /// key; when null, the style name is consulted.
    /// </summary>
    public static Style Create(string styleId, string styleName, MarkdownStyleKind kind, string? element)
    {
        var elementKey = element?.ToLowerInvariant();
        switch (kind)
        {
            case MarkdownStyleKind.Character:
                if (elementKey is "codeinline" or "code" || styleName.Equals("CodeChar", StringComparison.OrdinalIgnoreCase))
                {
                    return CodeChar(styleId, styleName);
                }
                if (elementKey == "hyperlink" || styleName.Equals("Hyperlink", StringComparison.OrdinalIgnoreCase))
                {
                    return Hyperlink(styleId, styleName);
                }
                return new Style(new StyleName { Val = styleName }) { Type = StyleValues.Character, StyleId = styleId };

            case MarkdownStyleKind.Table:
                return TableGrid(styleId, styleName);

            default:
                if (TryHeadingLevel(elementKey ?? styleName, out var level))
                {
                    return Heading(styleId, styleName, level);
                }
                if (elementKey is "paragraph" or "normal" || styleName.Equals("Normal", StringComparison.OrdinalIgnoreCase))
                {
                    return Body(styleId, styleName);
                }
                if (elementKey is "blockquote" or "quote" || styleName.Equals("Quote", StringComparison.OrdinalIgnoreCase))
                {
                    return Quote(styleId, styleName);
                }
                if (elementKey is "codeblock" or "code" || styleName.Equals("Code", StringComparison.OrdinalIgnoreCase))
                {
                    return CodeBlock(styleId, styleName);
                }
                if (elementKey == "definitionterm")
                {
                    return DefinitionTerm(styleId, styleName);
                }
                if (elementKey == "definitiondescription")
                {
                    return DefinitionDescription(styleId, styleName);
                }
                if (elementKey == "footnotetext")
                {
                    return FootnoteText(styleId, styleName);
                }
                if (elementKey == "list")
                {
                    return ListParagraph(styleId, styleName);
                }
                if (elementKey == "tableheader")
                {
                    return TableHeader(styleId, styleName);
                }
                if (elementKey == "thematicbreak")
                {
                    return ThematicBreak(styleId, styleName);
                }
                if (elementKey is "imagecaption" or "caption")
                {
                    return Caption(styleId, styleName);
                }
                return GenericParagraph(styleId, styleName);
        }
    }

    /// <summary>
    /// Heading fallback: distinct size per level, bold, an outline level (0-based), keep-with-next
    /// and decreasing space-before so the document's section structure is visually obvious.
    /// </summary>
    private static Style Heading(string styleId, string styleName, int level)
    {
        var size = level switch
        {
            1 => "48",
            2 => "40",
            3 => "32",
            4 => "28",
            5 => "24",
            _ => "22"
        };
        var before = (240 - (level - 1) * 30).ToString();
        return new Style(
            new StyleName { Val = styleName },
            new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new KeepNext(),
                new SpacingBetweenLines { Before = before, After = "60" },
                new OutlineLevel { Val = level - 1 }),
            new StyleRunProperties(
                new Bold(),
                new BoldComplexScript(),
                new FontSize { Val = size },
                new FontSizeComplexScript { Val = size }))
        {
            Type = StyleValues.Paragraph,
            StyleId = styleId
        };
    }

    /// <summary>
    /// Body fallback: a readable 11pt body with 8pt space-after and 1.15 line spacing, matching
    /// Word's modern default paragraph cadence.
    /// </summary>
    private static Style Body(string styleId, string styleName)
    {
        return new Style(
            new StyleName { Val = styleName },
            new StyleParagraphProperties(
                new SpacingBetweenLines { After = "160", Line = "276", LineRule = LineSpacingRuleValues.Auto }),
            new StyleRunProperties(
                new RunFonts { Ascii = BodyFont, HighAnsi = BodyFont, ComplexScript = BodyFont },
                new FontSize { Val = "22" },
                new FontSizeComplexScript { Val = "22" }))
        {
            Type = StyleValues.Paragraph,
            StyleId = styleId
        };
    }

    /// <summary>
    /// Quote fallback: a left-indented, italic block with a subtle left border that reads as a
    /// quotation without touching paragraph content.
    /// </summary>
    private static Style Quote(string styleId, string styleName)
    {
        return new Style(
            new StyleName { Val = styleName },
            new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new ParagraphBorders(new LeftBorder { Val = BorderValues.Single, Size = 8, Space = 4, Color = "7F7F7F" }),
                new Indentation { Left = "720" }),
            new StyleRunProperties(new Italic()))
        {
            Type = StyleValues.Paragraph,
            StyleId = styleId
        };
    }

    /// <summary>
    /// Code block fallback: monospace runs over a light shading, with single line spacing and no
    /// space-after so consecutive code lines stay visually contiguous.
    /// </summary>
    private static Style CodeBlock(string styleId, string styleName)
    {
        return new Style(
            new StyleName { Val = styleName },
            new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new Shading { Val = ShadingPatternValues.Clear, Fill = "F2F2F2" },
                new SpacingBetweenLines { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }),
            new StyleRunProperties(
                new RunFonts { Ascii = MonoFont, HighAnsi = MonoFont, ComplexScript = MonoFont },
                new FontSize { Val = "20" },
                new FontSizeComplexScript { Val = "20" }))
        {
            Type = StyleValues.Paragraph,
            StyleId = styleId
        };
    }

    /// <summary>
    /// Inline code character fallback: monospace runs over a light shading.
    /// </summary>
    private static Style CodeChar(string styleId, string styleName)
    {
        return new Style(
            new StyleName { Val = styleName },
            new StyleRunProperties(
                new RunFonts { Ascii = MonoFont, HighAnsi = MonoFont, ComplexScript = MonoFont },
                new FontSize { Val = "20" },
                new FontSizeComplexScript { Val = "20" },
                new Shading { Val = ShadingPatternValues.Clear, Fill = "F2F2F2" }))
        {
            Type = StyleValues.Character,
            StyleId = styleId
        };
    }

    /// <summary>
    /// Hyperlink character fallback: Word's theme hyperlink blue with a single underline.
    /// </summary>
    private static Style Hyperlink(string styleId, string styleName)
    {
        return new Style(
            new StyleName { Val = styleName },
            new StyleRunProperties(
                new Color { Val = "0563C1", ThemeColor = ThemeColorValues.Hyperlink },
                new Underline { Val = UnderlineValues.Single }))
        {
            Type = StyleValues.Character,
            StyleId = styleId
        };
    }

    /// <summary>
    /// Table fallback: a full single-line grid so every table cell is visibly framed.
    /// </summary>
    private static Style TableGrid(string styleId, string styleName)
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
            StyleTableProperties = new StyleTableProperties(borders)
        };
    }

    /// <summary>Definition term fallback: bold, kept with the definition that follows it.</summary>
    private static Style DefinitionTerm(string styleId, string styleName)
    {
        return new Style(
            new StyleName { Val = styleName },
            new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new KeepNext(),
                new SpacingBetweenLines { After = "0" }),
            new StyleRunProperties(
                new Bold(),
                new BoldComplexScript()))
        {
            Type = StyleValues.Paragraph,
            StyleId = styleId
        };
    }

    /// <summary>Definition description fallback: indented beneath its term.</summary>
    private static Style DefinitionDescription(string styleId, string styleName)
    {
        return new Style(
            new StyleName { Val = styleName },
            new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new SpacingBetweenLines { After = "80" },
                new Indentation { Left = "720" }))
        {
            Type = StyleValues.Paragraph,
            StyleId = styleId
        };
    }

    /// <summary>
    /// Footnote body fallback: compact 9pt text with single line spacing and no space-after, the
    /// standard footnote treatment.
    /// </summary>
    private static Style FootnoteText(string styleId, string styleName)
    {
        return new Style(
            new StyleName { Val = styleName },
            new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new SpacingBetweenLines { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }),
            new StyleRunProperties(
                new FontSize { Val = "18" },
                new FontSizeComplexScript { Val = "18" }))
        {
            Type = StyleValues.Paragraph,
            StyleId = styleId
        };
    }

    /// <summary>List item fallback: body text with a small space-after per item.</summary>
    private static Style ListParagraph(string styleId, string styleName)
    {
        return new Style(
            new StyleName { Val = styleName },
            new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new KeepLines(),
                new SpacingBetweenLines { After = "80" }))
        {
            Type = StyleValues.Paragraph,
            StyleId = styleId
        };
    }

    /// <summary>Table header fallback: bold, centered, kept with the table body.</summary>
    private static Style TableHeader(string styleId, string styleName)
    {
        return new Style(
            new StyleName { Val = styleName },
            new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new KeepLines(),
                new SpacingBetweenLines { After = "0" },
                new Justification { Val = JustificationValues.Center }),
            new StyleRunProperties(
                new Bold(),
                new BoldComplexScript()))
        {
            Type = StyleValues.Paragraph,
            StyleId = styleId
        };
    }

    /// <summary>Thematic-break fallback: a short, spaced paragraph before and after.</summary>
    private static Style ThematicBreak(string styleId, string styleName)
    {
        return new Style(
            new StyleName { Val = styleName },
            new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new KeepLines(),
                new SpacingBetweenLines { Before = "60", After = "60" }))
        {
            Type = StyleValues.Paragraph,
            StyleId = styleId
        };
    }

    /// <summary>Image-caption fallback: small, italic, centered.</summary>
    private static Style Caption(string styleId, string styleName)
    {
        return new Style(
            new StyleName { Val = styleName },
            new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new SpacingBetweenLines { After = "80" },
                new Justification { Val = JustificationValues.Center }),
            new StyleRunProperties(
                new Italic(),
                new FontSize { Val = "18" },
                new FontSizeComplexScript { Val = "18" }))
        {
            Type = StyleValues.Paragraph,
            StyleId = styleId
        };
    }

    /// <summary>Generic paragraph fallback: body-cadence spacing only.</summary>
    private static Style GenericParagraph(string styleId, string styleName)
    {
        return new Style(
            new StyleName { Val = styleName },
            new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(new SpacingBetweenLines { After = "80" }))
        {
            Type = StyleValues.Paragraph,
            StyleId = styleId
        };
    }

    /// <summary>True when <paramref name="value"/> is a heading key or name carrying a level 1–6.</summary>
    private static bool TryHeadingLevel(string value, out int level)
    {
        level = 0;
        if (string.IsNullOrEmpty(value)
            || !value.StartsWith("heading", StringComparison.OrdinalIgnoreCase)
            || !int.TryParse(value.AsSpan(7), out var parsed)
            || parsed is < 1 or > 6)
        {
            return false;
        }
        level = parsed;
        return true;
    }
}
