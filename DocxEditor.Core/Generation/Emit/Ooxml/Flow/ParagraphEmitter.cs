using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Flow;

/// <summary>
/// Emits flow paragraphs and headings: paragraph properties (style ID, outline level for
/// headings, alignment, spacing) plus runs with direct character formatting resolved from
/// typography tokens and individual run properties. Style IDs reference existing styles only.
/// </summary>
internal static class ParagraphEmitter
{
    /// <summary>Emits a body paragraph with the given content/style/heading level.</summary>
    public static void EmitParagraph(
        OoxmlEmitContext context,
        OpenXmlCompositeElement container,
        TextModel content,
        string? styleId,
        int? headingLevel,
        string path)
    {
        container.Append(BuildParagraph(context, content, styleId, headingLevel, path));
    }

    /// <summary>Emits a heading paragraph.</summary>
    public static void EmitHeading(OoxmlEmitContext context, OpenXmlCompositeElement container, HeadingBlock heading, string path)
    {
        var content = heading.Content;
        var styleId = heading.Style ?? $"Heading{heading.Level}";
        container.Append(BuildParagraph(context, content, styleId, heading.Level, path));
    }

    /// <summary>
    /// Builds a paragraph for a text model. Default run formatting comes from the content's
    /// typography token; headings without a token fall back to heading-level direct formatting.
    /// </summary>
    public static Paragraph BuildParagraph(
        OoxmlEmitContext context,
        TextModel content,
        string? styleId,
        int? headingLevel,
        string path)
    {
        var paragraph = new Paragraph();
        var paragraphProperties = new ParagraphProperties();

        if (styleId is not null && context.TryEnsureStyle(styleId, StyleValues.Paragraph, path))
        {
            paragraphProperties.ParagraphStyleId = new ParagraphStyleId { Val = styleId };
        }

        if (headingLevel is { } level)
        {
            paragraphProperties.OutlineLevel = new OutlineLevel { Val = level - 1 };
        }

        if (content.Alignment is { } alignment && FormattingHelpers.Justification(alignment) is { } justification)
        {
            paragraphProperties.Justification = new Justification { Val = justification };
        }

        if (content.Spacing is { } spacing)
        {
            var between = new SpacingBetweenLines();
            if (spacing.BeforePt is { } before)
            {
                between.Before = FormattingHelpers.Twips(before);
            }
            if (spacing.AfterPt is { } after)
            {
                between.After = FormattingHelpers.Twips(after);
            }
            if (spacing.LineMultiple is { } line)
            {
                between.Line = FormattingHelpers.LineMultiple(line);
                between.LineRule = LineSpacingRuleValues.Auto;
            }
            paragraphProperties.SpacingBetweenLines = between;
        }

        if (paragraphProperties.HasChildren)
        {
            paragraph.ParagraphProperties = paragraphProperties;
        }

        var defaults = content.Token is not null
            ? FormattingHelpers.ResolveTypographyToken(context, content.Token, path)
            : headingLevel is { } headingLevelValue ? FormattingHelpers.HeadingDefaults(context, headingLevelValue) : null;

        if (content.Runs is { Count: > 0 } runs)
        {
            foreach (var run in runs)
            {
                paragraph.Append(FormattingHelpers.BuildRun(context, run, defaults, path));
            }
        }
        else if (content.Text is { } text)
        {
            paragraph.Append(FormattingHelpers.BuildTextRun(context, text, defaults, path));
        }

        return paragraph;
    }
}
