using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Generation.Design;
using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Design;

/// <summary>
/// Builds the baseline <c>w:style</c> elements for <see cref="BaselineStyleKind"/>. All values are
/// deterministic and derive from the resolved design: paragraph-style kinds (title, subtitle,
/// eyebrow, headings, muted body, label, metric, metric label, table header/body, footer) are
/// generated from the active theme's semantic-role formatting (<see cref="ResolvedDesign.Roles"/>),
/// while Body, Table, Code and Callout keep dedicated treatments (callout derives its fill/accent
/// from the theme palette). Nothing forces a font the design did not specify, so document defaults
/// stay authoritative.
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
        if (kind == BaselineStyleKind.Code)
        {
            return CreateCode(styleId, styleName);
        }
        if (kind == BaselineStyleKind.Callout)
        {
            return CreateCallout(design, styleId, styleName, bodyStyleId);
        }

        return CreateRoleStyle(design, kind, styleId, styleName, bodyStyleId);
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
        // StyleTableProperties IS w:tblPr; its children are the CT_TblPrBase elements directly.
        return new Style
        {
            Type = StyleValues.Table,
            StyleId = styleId,
            StyleName = new StyleName { Val = styleName },
            StyleTableProperties = new StyleTableProperties(borders)
        };
    }

    private static Style CreateCode(string styleId, string styleName)
    {
        var style = new Style { Type = StyleValues.Paragraph, StyleId = styleId };
        style.Append(new StyleName { Val = styleName });
        style.Append(new StyleParagraphProperties(new Shading { Fill = "F5F5F5" }));
        style.Append(new StyleRunProperties(
            new RunFonts { Ascii = DocxDesignDefaults.CodeFont, HighAnsi = DocxDesignDefaults.CodeFont, ComplexScript = DocxDesignDefaults.CodeFont },
            new FontSize { Val = DocxFormattingHelpers.HalfPoints(DocxDesignDefaults.CodeFontSizePt) },
            new FontSizeComplexScript { Val = DocxFormattingHelpers.HalfPoints(DocxDesignDefaults.CodeFontSizePt) }));
        return style;
    }

    private static Style CreateCallout(ResolvedDesign design, string styleId, string styleName, string bodyStyleId)
    {
        var style = new Style { Type = StyleValues.Paragraph, StyleId = styleId };
        style.Append(new StyleName { Val = styleName });
        style.Append(new BasedOn { Val = bodyStyleId });
        style.Append(new NextParagraphStyle { Val = bodyStyleId });

        // w:pPr child order: pBdr → shd → spacing → ind → jc.
        var role = design.RoleOrDefault(TextRole.Callout);
        var properties = new StyleParagraphProperties();
        properties.Append(new ParagraphBorders(new LeftBorder
        {
            Val = BorderValues.Single,
            Color = ToColor(Primary(design)) ?? "808080",
            Size = 16
        }));
        properties.Append(new Shading { Fill = ToColor(PaletteColor(design, "pale")) ?? "F2F2F2" });
        if (BuildSpacing(role) is { } spacing)
        {
            properties.Append(spacing);
        }
        properties.Append(new Indentation { Left = "288", Right = "288" });
        style.Append(properties);
        style.Append(BuildRoleRunProperties(role));
        return style;
    }

    /// <summary>
    /// Generates a semantic paragraph style from the theme role formatting: paragraph properties
    /// (keep-next/lines, spacing, alignment, outline level) plus run properties (font, size, color,
    /// emphasis, caps).
    /// </summary>
    private static Style CreateRoleStyle(ResolvedDesign design, BaselineStyleKind kind, string styleId, string styleName, string bodyStyleId)
    {
        var style = new Style { Type = StyleValues.Paragraph, StyleId = styleId };
        style.Append(new StyleName { Val = styleName });
        style.Append(new BasedOn { Val = bodyStyleId });
        style.Append(new NextParagraphStyle { Val = bodyStyleId });

        var role = design.RoleOrDefault(kind.ToTextRole());
        style.Append(BuildRoleParagraphProperties(role, kind));
        style.Append(BuildRoleRunProperties(role));
        return style;
    }

    private static StyleParagraphProperties BuildRoleParagraphProperties(ResolvedRoleFormatting role, BaselineStyleKind kind)
    {
        // CT_PPr order: keepNext → keepLines → … → spacing → … → jc → … → outlineLvl.
        var properties = new StyleParagraphProperties();
        if (role.KeepNext)
        {
            properties.Append(new KeepNext());
        }
        if (role.KeepLines)
        {
            properties.Append(new KeepLines());
        }
        if (BuildSpacing(role) is { } spacing)
        {
            properties.Append(spacing);
        }
        if (role.Paragraph.Alignment is { } alignment)
        {
            properties.Append(new Justification { Val = DocxFormattingHelpers.ToJustification(alignment) });
        }
        if (kind.IsHeading())
        {
            properties.Append(new OutlineLevel { Val = kind.HeadingLevel() - 1 });
        }
        return properties;
    }

    private static StyleRunProperties BuildRoleRunProperties(ResolvedRoleFormatting role)
    {
        var run = role.Run;
        var properties = new StyleRunProperties();
        if (run.FontFamily is { } font)
        {
            properties.Append(new RunFonts { Ascii = font, HighAnsi = font, ComplexScript = font });
        }
        if (run.Bold is true)
        {
            properties.Append(new Bold());
            properties.Append(new BoldComplexScript());
        }
        if (run.Italic is true)
        {
            properties.Append(new Italic());
            properties.Append(new ItalicComplexScript());
        }
        if (run.Underline is true)
        {
            properties.Append(new Underline { Val = UnderlineValues.Single });
        }
        if (run.AllCaps is true)
        {
            properties.Append(new Caps());
        }
        if (run.ColorHex is { } color)
        {
            properties.Append(new Color { Val = ToColor(color) });
        }
        if (run.FontSizePt is { } size)
        {
            var halfPoints = DocxFormattingHelpers.HalfPoints(size);
            properties.Append(new FontSize { Val = halfPoints });
            properties.Append(new FontSizeComplexScript { Val = halfPoints });
        }
        return properties;
    }

    /// <summary>
    /// Builds the <c>w:spacing</c> element for a role, emitting only the fields that are set —
    /// assigning a null string to a StringValue property serializes an empty attribute, which the
    /// schema rejects. Returns null when nothing is set.
    /// </summary>
    private static SpacingBetweenLines? BuildSpacing(ResolvedRoleFormatting role)
    {
        var before = role.Paragraph.SpaceBeforePt;
        var after = role.Paragraph.SpaceAfterPt;
        var line = role.Paragraph.LineSpacingMultiple;
        if (before is null && after is null && line is null)
        {
            return null;
        }
        var spacing = new SpacingBetweenLines();
        if (before is { } beforePt)
        {
            spacing.Before = DocxFormattingHelpers.Twentieths(beforePt);
        }
        if (after is { } afterPt)
        {
            spacing.After = DocxFormattingHelpers.Twentieths(afterPt);
        }
        if (line is { } lineMultiple)
        {
            spacing.Line = DocxFormattingHelpers.LineMultiple(lineMultiple);
            spacing.LineRule = LineSpacingRuleValues.Auto;
        }
        return spacing;
    }

    /// <summary>The palette's "primary" color as #RRGGBB, when the design defines it.</summary>
    private static string? Primary(ResolvedDesign design) =>
        design.Palette.TryGetValue("primary", out var hex) ? hex : null;

    /// <summary>A palette token's #RRGGBB, when defined.</summary>
    private static string? PaletteColor(ResolvedDesign design, string token) =>
        design.Palette.TryGetValue(token, out var hex) ? hex : null;

    /// <summary>Normalizes a #RRGGBB to the OOXML color form (no leading '#').</summary>
    private static string? ToColor(string? hex) => DocxFormattingHelpers.ToOoxmlColor(hex);
}
