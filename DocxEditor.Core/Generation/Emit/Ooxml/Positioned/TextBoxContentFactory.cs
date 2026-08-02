using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Model = DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Positioned;

/// <summary>
/// Builds the <c>w:txbxContent</c> body of a positioned text box / callout from the
/// shared <see cref="TextModel"/>. Text boxes store WML paragraphs (not DrawingML), so
/// this produces a single <c>w:p</c> whose runs carry WML run properties resolved from
/// the model's typography token + per-run overrides and the design tokens. Newlines
/// inside a run become <c>w:br</c> line breaks (single-paragraph semantics, matching the
/// flow tier's treatment of text as one paragraph). Paragraph alignment and spacing are
/// emitted onto <c>w:jc</c>/<c>w:spacing</c> where represented.
/// </summary>
internal sealed class TextBoxContentFactory
{
    private readonly DesignTokenResolver _design;

    public TextBoxContentFactory(DesignTokenResolver design)
    {
        _design = design;
    }

    /// <summary>Builds the <c>w:txbxContent</c> element for a text model.</summary>
    public W.TextBoxContent Build(Model.TextModel content)
    {
        ArgumentNullException.ThrowIfNull(content);

        W.TextBoxContent textBoxContent = new();
        W.Paragraph paragraph = BuildParagraph(content);
        textBoxContent.Append(paragraph);
        return textBoxContent;
    }

    private W.Paragraph BuildParagraph(Model.TextModel content)
    {
        W.Paragraph paragraph = new();

        W.ParagraphProperties? properties = BuildParagraphProperties(content);
        if (properties is not null)
        {
            paragraph.Append(properties);
        }

        if (content.Text is not null)
        {
            paragraph.Append(BuildRun(content.Text, baseFormatting: ResolveToken(content.Token), overrides: null));
            return paragraph;
        }

        if (content.Runs is not null)
        {
            foreach (Model.Run run in content.Runs)
            {
                paragraph.Append(BuildRun(run.Text, ResolveToken(content.Token), run));
            }

            return paragraph;
        }

        // Blank text box (parser allows neither text nor runs only for table cells, so
        // this is defensive): a placeholder paragraph with a single empty run.
        paragraph.Append(BuildRun(string.Empty, ResolveToken(content.Token), null));
        return paragraph;
    }

    private W.ParagraphProperties? BuildParagraphProperties(Model.TextModel content)
    {
        W.ParagraphProperties? properties = null;

        if (content.Alignment is { } alignment)
        {
            properties ??= new W.ParagraphProperties();
            properties.Append(new W.Justification { Val = MapAlignment(alignment) });
        }

        if (content.Spacing is { } spacing)
        {
            properties ??= new W.ParagraphProperties();
            W.SpacingBetweenLines spacingBetweenLines = new();
            if (spacing.BeforePt is { } before)
            {
                spacingBetweenLines.Before = PtToTwips(before);
            }

            if (spacing.AfterPt is { } after)
            {
                spacingBetweenLines.After = PtToTwips(after);
            }

            if (spacing.LineMultiple is { } multiple)
            {
                spacingBetweenLines.Line = Invariant(((int)Math.Round(multiple * 240.0, MidpointRounding.AwayFromZero)));
                spacingBetweenLines.LineRule = W.LineSpacingRuleValues.Auto;
            }

            properties.Append(spacingBetweenLines);
        }

        return properties;
    }

    private W.Run BuildRun(string text, Model.TypographyToken? baseFormatting, Model.Run? overrides)
    {
        W.RunProperties runProperties = new();

        string? fontFamily = _design.ResolveFontFamily(overrides?.FontFamily)
            ?? _design.ResolveFontFamily(baseFormatting?.FontFamily)
            ?? _design.DefaultBodyFontFamily;
        if (fontFamily is not null)
        {
            runProperties.Append(new W.RunFonts { Ascii = fontFamily, HighAnsi = fontFamily, ComplexScript = fontFamily });
        }

        if (overrides?.FontSizePt is { } runSize)
        {
            runProperties.Append(new W.FontSize { Val = HalfPoints(runSize) });
            runProperties.Append(new W.FontSizeComplexScript { Val = HalfPoints(runSize) });
        }
        else if (baseFormatting?.SizePt is { } tokenSize)
        {
            runProperties.Append(new W.FontSize { Val = HalfPoints(tokenSize) });
            runProperties.Append(new W.FontSizeComplexScript { Val = HalfPoints(tokenSize) });
        }

        bool bold = overrides?.Bold == true || baseFormatting?.Bold == true;
        bool italic = overrides?.Italic == true || baseFormatting?.Italic == true;
        bool underline = overrides?.Underline == true || baseFormatting?.Underline == true;
        if (bold)
        {
            runProperties.Append(new W.Bold());
        }

        if (italic)
        {
            runProperties.Append(new W.Italic());
        }

        if (underline)
        {
            runProperties.Append(new W.Underline { Val = W.UnderlineValues.Single });
        }

        string colorHex = _design.ResolveHex(
            overrides?.Color ?? baseFormatting?.Color,
            _design.DefaultTextHex);
        runProperties.Append(new W.Color { Val = colorHex });

        // Split the run text on newlines into w:t segments separated by w:br.
        W.Run wordprocessingRun = new(runProperties);
        string[] segments = text.Split('\n');
        for (int i = 0; i < segments.Length; i++)
        {
            if (i > 0)
            {
                wordprocessingRun.Append(new W.Break());
            }

            wordprocessingRun.Append(CreateTextElement(segments[i]));
        }

        return wordprocessingRun;
    }

    private static W.Text CreateTextElement(string value)
    {
        W.Text text = new(value);
        if (value.Length > 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[^1])))
        {
            text.Space = SpaceProcessingModeValues.Preserve;
        }

        return text;
    }

    private Model.TypographyToken? ResolveToken(string? name) => _design.TryGetTypographyToken(name);

    private static W.JustificationValues MapAlignment(Model.TextAlignment alignment) => alignment switch
    {
        Model.TextAlignment.Left => W.JustificationValues.Left,
        Model.TextAlignment.Center => W.JustificationValues.Center,
        Model.TextAlignment.Right => W.JustificationValues.Right,
        Model.TextAlignment.Justify => W.JustificationValues.Both,
        _ => W.JustificationValues.Left
    };

    private static string PtToTwips(double points) => Invariant((int)Math.Round(points * 20.0, MidpointRounding.AwayFromZero));

    private static string HalfPoints(double points) => Invariant((int)Math.Round(points * 2.0, MidpointRounding.AwayFromZero));

    private static string Invariant(int value) => value.ToString(System.Globalization.CultureInfo.InvariantCulture);
}
