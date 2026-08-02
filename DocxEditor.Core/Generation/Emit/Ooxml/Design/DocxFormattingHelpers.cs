using System.Globalization;
using DocxEditor.Core.Generation.Design;
using DocxEditor.Core.Generation.Model;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Design;

/// <summary>
/// Conversion + construction helpers that translate <see cref="DocxEditor.Core.Generation.Design"/>
/// resolved values into OpenXML WordprocessingML elements. Emitters use these to apply run and
/// paragraph formatting as direct properties when a style is not appropriate (inline emphasis,
/// table cells, positioned text boxes, one-off overrides), and to build paragraphs/runs from
/// scratch. All output is deterministic: null/empty resolved values emit nothing.
/// </summary>
public static class DocxFormattingHelpers
{
    // ---- unit conversions ----

    /// <summary>Converts points to OOXML half-points (FontSize).</summary>
    public static string HalfPoints(double pt) => (pt * 2).ToString("0.###", CultureInfo.InvariantCulture);

    /// <summary>Converts points to OOXML twentieths of a point (spacing, indents).</summary>
    public static string Twentieths(double pt) => (pt * 20).ToString("0.###", CultureInfo.InvariantCulture);

    /// <summary>Converts a line-spacing multiple to the OOXML Line value for LineRule=Auto (240 = single).</summary>
    public static string LineMultiple(double multiple) =>
        ((int)Math.Round(multiple * 240, MidpointRounding.AwayFromZero)).ToString(CultureInfo.InvariantCulture);

    /// <summary>Normalizes a #RRGGBB color to the OOXML form (no leading '#', upper case).</summary>
    public static string? ToOoxmlColor(string? hex) =>
        hex is null ? null : hex.TrimStart('#').ToUpperInvariant();

    // ---- run properties ----

    /// <summary>
    /// Builds the <c>w:rPr</c> element for a resolved run formatting (font family, size, color,
    /// bold/italic/underline). Returns null when nothing is set, so emitters can omit properties.
    /// </summary>
    public static Wp.RunProperties? BuildRunProperties(ResolvedRunFormatting? formatting)
    {
        if (formatting is null || formatting.IsEmpty)
        {
            return null;
        }
        var properties = new Wp.RunProperties();
        if (formatting.FontFamily is { } font)
        {
            properties.Append(new Wp.RunFonts { Ascii = font, HighAnsi = font, ComplexScript = font });
        }
        if (formatting.Bold == true)
        {
            properties.Append(new Wp.Bold());
            properties.Append(new Wp.BoldComplexScript());
        }
        if (formatting.Italic == true)
        {
            properties.Append(new Wp.Italic());
            properties.Append(new Wp.ItalicComplexScript());
        }
        if (formatting.ColorHex is { } color)
        {
            properties.Append(new Wp.Color { Val = ToOoxmlColor(color) });
        }
        if (formatting.FontSizePt is { } size)
        {
            var halfPoints = HalfPoints(size);
            properties.Append(new Wp.FontSize { Val = halfPoints });
            properties.Append(new Wp.FontSizeComplexScript { Val = halfPoints });
        }
        if (formatting.Underline == true)
        {
            properties.Append(new Wp.Underline { Val = Wp.UnderlineValues.Single });
        }
        if (formatting.AllCaps == true)
        {
            properties.Append(new Wp.Caps());
        }
        return properties;
    }

    /// <summary>Builds a single run with the resolved formatting applied as direct properties.</summary>
    public static Wp.Run BuildRun(string text, ResolvedRunFormatting? formatting)
    {
        var run = new Wp.Run();
        if (BuildRunProperties(formatting) is { } properties)
        {
            run.Append(properties);
        }
        run.Append(new Wp.Text(text) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve });
        return run;
    }

    // ---- paragraph properties ----

    /// <summary>
    /// Builds the <c>w:pPr</c> element for a resolved paragraph formatting (alignment, space
    /// before/after, line spacing multiple). Returns null when nothing is set.
    /// </summary>
    public static Wp.ParagraphProperties? BuildParagraphProperties(ResolvedParagraphFormatting? formatting)
    {
        if (formatting is null || formatting.IsEmpty)
        {
            return null;
        }
        var properties = new Wp.ParagraphProperties();
        var before = formatting.SpaceBeforePt;
        var after = formatting.SpaceAfterPt;
        var line = formatting.LineSpacingMultiple;
        if (before is not null || after is not null || line is not null)
        {
            var spacing = new Wp.SpacingBetweenLines();
            if (before is not null)
            {
                spacing.Before = Twentieths(before.Value);
            }
            if (after is not null)
            {
                spacing.After = Twentieths(after.Value);
            }
            if (line is not null)
            {
                spacing.Line = LineMultiple(line.Value);
                spacing.LineRule = Wp.LineSpacingRuleValues.Auto;
            }
            properties.Append(spacing);
        }
        if (formatting.Alignment is { } alignment)
        {
            properties.Append(new Wp.Justification { Val = ToJustification(alignment) });
        }
        return properties;
    }

    /// <summary>
    /// Builds a paragraph whose properties carry an optional style reference plus resolved direct
    /// paragraph formatting, containing one run with the resolved run formatting. This is the
    /// "direct properties instead of a style" helper for body text, cells and text boxes.
    /// </summary>
    public static Wp.Paragraph BuildParagraph(
        string text,
        ResolvedRunFormatting? runFormatting,
        ResolvedParagraphFormatting? paragraphFormatting,
        string? styleId = null)
    {
        var paragraph = new Wp.Paragraph();
        var hasParagraphFormatting = paragraphFormatting is { IsEmpty: false };
        if (styleId is not null || hasParagraphFormatting)
        {
            var properties = BuildParagraphProperties(paragraphFormatting) ?? new Wp.ParagraphProperties();
            if (styleId is not null)
            {
                // w:pStyle must be the first child of w:pPr; the property setter inserts it in schema position.
                properties.ParagraphStyleId = new Wp.ParagraphStyleId { Val = styleId };
            }
            paragraph.Append(properties);
        }
        paragraph.Append(BuildRun(text, runFormatting));
        return paragraph;
    }

    /// <summary>Maps a vocabulary alignment to its WordprocessingML justification value.</summary>
    public static Wp.JustificationValues ToJustification(TextAlignment alignment) => alignment switch
    {
        TextAlignment.Center => Wp.JustificationValues.Center,
        TextAlignment.Right => Wp.JustificationValues.Right,
        TextAlignment.Justify => Wp.JustificationValues.Both,
        _ => Wp.JustificationValues.Left
    };

    // ---- table cell helpers ----

    /// <summary>
    /// Applies a solid fill to a table cell via <c>w:tcPr/w:shd</c> (used for header rows and
    /// emphasized cells). No-op when <paramref name="fillHex"/> is null; any existing shading on the
    /// cell is replaced so repeated calls are idempotent and deterministic.
    /// </summary>
    public static void ApplyCellShading(Wp.TableCell cell, string? fillHex)
    {
        if (fillHex is null)
        {
            return;
        }
        var cellProperties = cell.TableCellProperties ??= new Wp.TableCellProperties();
        cellProperties.RemoveAllChildren<Wp.Shading>();
        cellProperties.Append(new Wp.Shading { Val = Wp.ShadingPatternValues.Clear, Color = "auto", Fill = ToOoxmlColor(fillHex) });
    }
}
