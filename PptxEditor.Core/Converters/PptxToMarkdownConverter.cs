using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxEditor.Core.Converters;

/// <summary>
/// Extracts a structured per-slide Markdown text outline from a PPTX deck. A slide is a
/// POSITIONED document, so the output is NOT a visual reproduction — it is a content
/// outline (slide headings, titles, bullet lists with nesting, tables, free-floating text
/// paragraphs, speaker notes, image references and labeled placeholders for charts and
/// SmartArt) suited for review and LLM consumption.
///
/// Slide read patterns mirror <see cref="PptxToTypstConverter"/> (slide iteration via
/// <c>SlideIdList</c>/<c>GetPartById</c>; unreliable attributes read via regex on
/// OuterXml per the PPTX converter conventions). Only the slide's own shape tree is read — layout
/// and master shapes are deliberately skipped so deck furniture (logos, footers) does not
/// pollute the outline.
///
/// Images referenced as <c>assets/image_N.ext</c> are extracted to the directory given to
/// <see cref="Convert(DocumentFormat.OpenXml.Packaging.PresentationDocument,string)"/> or
/// created by <see cref="ConvertToFile"/> next to the output file. <see cref="Convert(DocumentFormat.OpenXml.Packaging.PresentationDocument)"/>
/// returns the same outline without writing media.
/// </summary>
public sealed class PptxToMarkdownConverter
{
    /// <summary>
    /// Converts the deck to a Markdown outline. Images are referenced as
    /// <c>assets/image_N.ext</c> but NOT written; use
    /// <see cref="Convert(DocumentFormat.OpenXml.Packaging.PresentationDocument,string)"/>
    /// or <see cref="ConvertToFile"/> to extract the media next to the output.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="document"/> is <see langword="null"/>.</exception>
    public static string Convert(PresentationDocument document)
        => new PptxToMarkdownConverter().ConvertDocument(document, assetsDirectory: null);

    /// <summary>
    /// Converts the deck to a Markdown outline and extracts every image into
    /// <paramref name="assetsDirectory"/>, referenced from the outline as
    /// <c>assets/image_N.ext</c>.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="document"/> is <see langword="null"/>.</exception>
    public static string Convert(PresentationDocument document, string assetsDirectory)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(assetsDirectory);
        return new PptxToMarkdownConverter().ConvertDocument(document, assetsDirectory);
    }

    /// <summary>
    /// Opens <paramref name="sourcePath"/> (read-only), converts it to a Markdown outline,
    /// extracts every image into an <c>assets/</c> directory created next to
    /// <paramref name="outputPath"/>, and writes the Markdown file.
    /// </summary>
    /// <exception cref="ArgumentException">A path is null or whitespace.</exception>
    /// <exception cref="FileNotFoundException"><paramref name="sourcePath"/> does not exist.</exception>
    public static void ConvertToFile(string sourcePath, string outputPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sourcePath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        if (!File.Exists(sourcePath))
            throw new FileNotFoundException($"Source PPTX file not found: {sourcePath}", sourcePath);

        var fullOutputPath = Path.GetFullPath(outputPath);
        var outputDirectory = Path.GetDirectoryName(fullOutputPath);
        if (!string.IsNullOrEmpty(outputDirectory))
            Directory.CreateDirectory(outputDirectory);

        var assetsDirectory = Path.Combine(outputDirectory ?? ".", "assets");
        Directory.CreateDirectory(assetsDirectory);

        using var document = PresentationDocument.Open(sourcePath, false);
        var markdown = Convert(document, assetsDirectory);
        File.WriteAllText(fullOutputPath, markdown);
    }

    private string? _assetsDirectory;
    private int _imageCounter;

    private string ConvertDocument(PresentationDocument document, string? assetsDirectory)
    {
        ArgumentNullException.ThrowIfNull(document);

        _assetsDirectory = assetsDirectory;
        _imageCounter = 0;

        var slideIdList = document.PresentationPart?.Presentation?.SlideIdList;
        if (slideIdList == null) return string.Empty;

        var markdown = new StringBuilder();
        int slideIndex = 1;
        foreach (var slideId in slideIdList.ChildElements.OfType<SlideId>())
        {
            var slidePart = (SlidePart)document.PresentationPart!.GetPartById(slideId.RelationshipId!);
            var slideText = ConvertSlide(slidePart, slideIndex, isFirstSlide: slideIndex == 1).ToString().TrimEnd();
            if (slideText.Length > 0)
            {
                if (markdown.Length > 0)
                    markdown.AppendLine().AppendLine("---").AppendLine();
                markdown.Append(slideText);
            }
            slideIndex++;
        }

        return markdown.ToString().TrimEnd();
    }

    private StringBuilder ConvertSlide(SlidePart slidePart, int slideIndex, bool isFirstSlide)
    {
        var sb = new StringBuilder();
        var shapeTree = slidePart.Slide?.CommonSlideData?.ShapeTree;
        if (shapeTree == null) return sb;

        // The slide title comes from its title placeholder (top-level shapes only). Its
        // text becomes the heading; the shape is then skipped during content iteration.
        P.Shape? titleShape = null;
        string? titleText = null;
        foreach (var shape in shapeTree.Elements<P.Shape>())
        {
            if (IsHidden(shape)) continue;
            if (GetPlaceholderType(shape) is not ("title" or "ctrTitle")) continue;
            var text = ExtractShapeText(shape, slideIndex);
            if (text.Length == 0) continue;
            titleShape = shape;
            titleText = text;
            break;
        }

        if (titleShape != null && isFirstSlide)
        {
            sb.AppendLine($"# {EscapeLine(titleText!)}");
        }
        else
        {
            sb.AppendLine($"## Slide {slideIndex}");
            if (titleText != null)
                sb.AppendLine($"### {EscapeLine(titleText)}");
        }
        sb.AppendLine();

        foreach (var element in shapeTree.ChildElements)
        {
            if (ReferenceEquals(element, titleShape)) continue;
            if (IsHidden(element)) continue;
            var before = sb.Length;
            ConvertElement(slidePart, element, slideIndex, sb);
            if (sb.Length > before)
                sb.AppendLine();
        }

        AppendNotes(slidePart, slideIndex, sb);

        return sb;
    }

    private void ConvertElement(SlidePart slidePart, OpenXmlElement element, int slideIndex, StringBuilder sb)
    {
        switch (element)
        {
            case P.Shape shape:
                ConvertShape(slidePart, shape, slideIndex, sb);
                break;
            case P.Picture picture:
                ConvertPicture(slidePart, picture, sb);
                break;
            case P.GraphicFrame frame:
                ConvertGraphicFrame(slidePart, frame, slideIndex, sb);
                break;
            case P.GroupShape group:
                foreach (var child in group.ChildElements)
                {
                    if (IsHidden(child)) continue;
                    ConvertElement(slidePart, child, slideIndex, sb);
                }
                break;
            case P.ConnectionShape connection:
                AppendTextBody(connection.GetFirstChild<P.TextBody>(), slideIndex, sb);
                break;
        }
    }

    private void ConvertShape(SlidePart slidePart, P.Shape shape, int slideIndex, StringBuilder sb)
    {
        var blipFill = shape.ShapeProperties?.Elements<Drawing.BlipFill>().FirstOrDefault();
        var paragraphs = ExtractTextParagraphs(shape.TextBody, slideIndex);

        // An image-filled shape with no text is treated as a picture.
        if (blipFill != null && paragraphs.Count == 0)
        {
            ConvertBlipFillImage(slidePart, blipFill, GetShapeName(shape) ?? "image", sb);
            return;
        }

        if (paragraphs.Count == 0) return;

        var isBodyPlaceholder = GetPlaceholderType(shape) == "body";
        foreach (var paragraph in paragraphs)
        {
            var isBullet = (isBodyPlaceholder && !paragraph.HasNoBullet) || paragraph.HasBullet;
            if (isBullet)
            {
                sb.Append(' ', paragraph.Level * 2)
                    .Append("- ")
                    .AppendLine(EscapeLine(paragraph.Text));
            }
            else
            {
                sb.AppendLine(EscapeLine(paragraph.Text));
            }
        }
    }

    private void ConvertPicture(SlidePart slidePart, P.Picture picture, StringBuilder sb)
    {
        var embed = picture.BlipFill?.Blip?.Embed?.Value;
        if (string.IsNullOrEmpty(embed))
        {
            sb.AppendLine($"<!-- image: {EscapeComment(GetShapeName(picture) ?? "picture")} (no embed reference) -->");
            return;
        }

        var imagePart = TryGetImagePart(slidePart, embed)
            ?? TryGetImagePart(slidePart.SlideLayoutPart, embed);
        if (imagePart == null)
        {
            sb.AppendLine($"<!-- image: {EscapeComment(GetShapeName(picture) ?? "picture")} (media part not found) -->");
            return;
        }

        var alt = GetPictureAltText(picture) ?? $"image {_imageCounter + 1}";
        var fileName = ExtractImage(imagePart);
        sb.AppendLine($"![{EscapeAltText(alt)}](assets/{fileName})");
    }

    private void ConvertBlipFillImage(SlidePart slidePart, Drawing.BlipFill blipFill, string alt, StringBuilder sb)
    {
        var embed = blipFill.Blip?.Embed?.Value;
        if (string.IsNullOrEmpty(embed)) return;

        var imagePart = TryGetImagePart(slidePart, embed)
            ?? TryGetImagePart(slidePart.SlideLayoutPart, embed);
        if (imagePart == null)
        {
            sb.AppendLine($"<!-- image: {EscapeComment(alt)} (media part not found) -->");
            return;
        }

        var fileName = ExtractImage(imagePart);
        sb.AppendLine($"![{EscapeAltText(alt)}](assets/{fileName})");
    }

    private void ConvertGraphicFrame(SlidePart slidePart, P.GraphicFrame frame, int slideIndex, StringBuilder sb)
    {
        var graphicData = frame.Graphic?.GraphicData;
        if (graphicData == null) return;

        var table = graphicData.Elements<Drawing.Table>().FirstOrDefault();
        if (table != null)
        {
            AppendTable(table, slideIndex, sb);
            return;
        }

        var uri = graphicData.Uri?.Value ?? string.Empty;
        if (uri.Contains("/drawingml/2006/chart", StringComparison.Ordinal))
        {
            var chartType = ReadChartTypeName(slidePart, graphicData);
            sb.AppendLine(chartType != null
                ? $"<!-- chart: {Capitalize(chartType)} -->"
                : "<!-- chart -->");
            return;
        }

        if (uri.Contains("/drawingml/2006/diagram", StringComparison.Ordinal))
        {
            var name = GetShapeName(frame);
            sb.AppendLine(name != null
                ? $"<!-- smartart: {EscapeComment(name)} -->"
                : "<!-- smartart -->");
            return;
        }

        var kind = ClassifyOtherGraphicFrame(uri);
        sb.AppendLine($"<!-- {kind} -->");
    }

    private static void AppendTable(Drawing.Table table, int slideIndex, StringBuilder sb)
    {
        var rows = table.Elements<Drawing.TableRow>().ToList();
        if (rows.Count == 0) return;

        var columnCount = table.TableGrid?.Elements<Drawing.GridColumn>().Count()
            ?? rows.Max(r => r.Elements<Drawing.TableCell>().Count());

        // The first row is a header when the table opts in via tblPr firstRow="1".
        var hasHeader = table.TableProperties?.FirstRow?.Value == true;

        for (int i = 0; i < rows.Count; i++)
        {
            var cells = rows[i].Elements<Drawing.TableCell>()
                .Select(cell => EscapeTableCell(ExtractCellText(cell, slideIndex)))
                .ToList();
            while (cells.Count < columnCount) cells.Add(string.Empty);

            sb.Append("| ").Append(string.Join(" | ", cells)).AppendLine(" |");
            if (hasHeader && i == 0)
            {
                sb.Append("| ")
                    .Append(string.Join(" | ", Enumerable.Repeat("---", columnCount)))
                    .AppendLine(" |");
            }
        }
    }

    private void AppendNotes(SlidePart slidePart, int slideIndex, StringBuilder sb)
    {
        var shapeTree = slidePart.NotesSlidePart?.NotesSlide?.CommonSlideData?.ShapeTree;
        if (shapeTree == null) return;

        var notes = new List<string>();
        foreach (var shape in shapeTree.Elements<P.Shape>())
        {
            if (IsHidden(shape)) continue;
            foreach (var paragraph in ExtractTextParagraphs(shape.TextBody, slideIndex))
            {
                if (paragraph.Text.Length > 0) notes.Add(paragraph.Text);
            }
        }

        if (notes.Count == 0) return;

        sb.AppendLine();
        sb.AppendLine("> [!note]");
        foreach (var note in notes)
            sb.AppendLine($"> {EscapeLine(note)}");
    }

    private static string ExtractShapeText(P.Shape shape, int slideIndex)
    {
        var paragraphs = ExtractTextParagraphs(shape.TextBody, slideIndex);
        return string.Join(" ", paragraphs.Select(p => p.Text));
    }

    private static List<TextParagraph> ExtractTextParagraphs(P.TextBody? textBody, int slideIndex)
    {
        var result = new List<TextParagraph>();
        if (textBody == null) return result;

        foreach (var paragraph in textBody.Elements<Drawing.Paragraph>())
        {
            var text = ExtractParagraphText(paragraph, slideIndex);
            if (text.Length == 0) continue;

            var pPr = paragraph.ParagraphProperties;
            result.Add(new TextParagraph(
                text,
                GetLevel(pPr),
                HasBulletMarker(pPr),
                HasNoBullet(pPr)));
        }

        return result;
    }

    /// <summary>Appends a text body's paragraphs as plain paragraphs (non-placeholder
    /// shapes such as connection labels).</summary>
    private static void AppendTextBody(P.TextBody? textBody, int slideIndex, StringBuilder sb)
    {
        foreach (var paragraph in ExtractTextParagraphs(textBody, slideIndex))
            sb.AppendLine(EscapeLine(paragraph.Text));
    }

    private static string ExtractParagraphText(Drawing.Paragraph paragraph, int slideIndex)
    {
        var sb = new StringBuilder();
        foreach (var child in paragraph.ChildElements)
        {
            switch (child)
            {
                case Drawing.Run run when run.Text?.Text != null:
                    sb.Append(run.Text.Text);
                    break;
                case Drawing.Field field:
                    sb.Append(ResolveFieldText(field, slideIndex));
                    break;
                case Drawing.Break:
                    sb.Append(' ');
                    break;
            }
        }

        // Soft breaks (a:br) and literal newlines inside a:t collapse to a space: the
        // Markdown outline treats a paragraph as a single line.
        return sb.ToString().Replace('\r', ' ').Replace('\n', ' ').Trim();
    }

    /// <summary>Resolves the display text of an <c>a:fld</c>: slide-number fields render
    /// the actual 1-based slide index; other fields keep their cached text.</summary>
    private static string ResolveFieldText(Drawing.Field field, int slideIndex)
    {
        var typeMatch = Regex.Match(field.OuterXml, @"\btype\s*=\s*""([^""]*)""");
        var fieldType = typeMatch.Success ? typeMatch.Groups[1].Value : null;
        return fieldType == "slidenum"
            ? slideIndex.ToString(CultureInfo.InvariantCulture)
            : field.Text?.Text ?? string.Empty;
    }

    private static int GetLevel(Drawing.ParagraphProperties? pPr)
    {
        if (pPr == null) return 0;
        var match = Regex.Match(pPr.OuterXml, @"\blvl\s*=\s*""(\d+)""");
        return match.Success
            && int.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var level)
            ? Math.Clamp(level, 0, 6)
            : 0;
    }

    private static bool HasBulletMarker(Drawing.ParagraphProperties? pPr)
    {
        if (pPr == null) return false;
        var xml = pPr.OuterXml;
        return Regex.IsMatch(xml, @"<a:buChar\b") || Regex.IsMatch(xml, @"<a:buAutoNum\b");
    }

    private static bool HasNoBullet(Drawing.ParagraphProperties? pPr)
        => pPr != null && Regex.IsMatch(pPr.OuterXml, @"<a:buNone\b");

    /// <summary>
    /// Reads the placeholder <c>type</c> attribute of a shape's <c>p:ph</c> element via
    /// regex on OuterXml (raw XML attribute reads); returns null for non-placeholder shapes.
    /// </summary>
    private static string? GetPlaceholderType(P.Shape shape)
    {
        var placeholder = shape.NonVisualShapeProperties?.Descendants<PlaceholderShape>().FirstOrDefault();
        if (placeholder == null) return null;
        var match = Regex.Match(placeholder.OuterXml, @"\btype\s*=\s*""([^""]*)""");
        return match.Success ? match.Groups[1].Value : null;
    }

    private static string ExtractCellText(Drawing.TableCell cell, int slideIndex)
    {
        var paragraphs = cell.TextBody?.Elements<Drawing.Paragraph>().ToList();
        if (paragraphs == null || paragraphs.Count == 0) return string.Empty;
        return string.Join(" ", paragraphs
            .Select(p => ExtractParagraphText(p, slideIndex))
            .Where(text => text.Length > 0));
    }

    private static string? ReadChartTypeName(SlidePart slidePart, Drawing.GraphicData graphicData)
    {
        var relIdMatch = Regex.Match(graphicData.OuterXml, @"\br:id\s*=\s*""([^""]*)""");
        if (!relIdMatch.Success) return null;

        try
        {
            var part = slidePart.GetPartById(relIdMatch.Groups[1].Value);
            using var stream = part.GetStream();
            using var reader = new StreamReader(stream);
            var xml = reader.ReadToEnd();

            var match = Regex.Match(xml, @"<(?:\w+:)?(\w+Chart)\b");
            return match.Success ? match.Groups[1].Value : null;
        }
        catch
        {
            return null;
        }
    }

    private static string ClassifyOtherGraphicFrame(string uri)
    {
        if (uri.Contains("oleObject", StringComparison.OrdinalIgnoreCase))
            return "embedded object";
        if (uri.Contains("video", StringComparison.OrdinalIgnoreCase)
            || uri.Contains("audio", StringComparison.OrdinalIgnoreCase)
            || uri.Contains("media", StringComparison.OrdinalIgnoreCase))
            return "media";
        return "graphic frame";
    }

    /// <summary>Resolves an <c>r:embed</c> relationship id to its image part. The owner
    /// part's relationships win (rIds collide across slide/layout parts); the layout is
    /// a fallback for pictures inherited from the layout.</summary>
    private static ImagePart? TryGetImagePart(OpenXmlPartContainer? container, string relationshipId)
    {
        if (container == null || string.IsNullOrEmpty(relationshipId)) return null;
        foreach (var part in container.Parts)
        {
            if (part.RelationshipId == relationshipId && part.OpenXmlPart is ImagePart imagePart)
                return imagePart;
        }
        return null;
    }

    /// <summary>Writes the image to the assets directory (when one is configured) and
    /// returns the file name the outline references.</summary>
    private string ExtractImage(ImagePart imagePart)
    {
        var extension = imagePart.ContentType switch
        {
            "image/png" => "png",
            "image/jpeg" => "jpg",
            "image/gif" => "gif",
            "image/bmp" => "bmp",
            "image/tiff" => "tiff",
            "image/x-icon" => "ico",
            "image/svg+xml" => "svg",
            _ => "bin"
        };

        _imageCounter++;
        var fileName = $"image_{_imageCounter}.{extension}";
        if (_assetsDirectory != null)
        {
            var fullPath = Path.Combine(_assetsDirectory, fileName);
            using var stream = imagePart.GetStream();
            using var file = File.Create(fullPath);
            stream.CopyTo(file);
        }
        return fileName;
    }

    private static string? GetPictureAltText(P.Picture picture)
    {
        var cNvPr = picture.NonVisualPictureProperties?.NonVisualDrawingProperties;
        var descr = GetCnvPrAttribute(cNvPr, "descr");
        if (!string.IsNullOrWhiteSpace(descr)) return descr;
        var name = GetCnvPrAttribute(cNvPr, "name");
        return !string.IsNullOrWhiteSpace(name) ? name : null;
    }

    private static string? GetShapeName(OpenXmlElement element)
    {
        var cNvPr = element.Descendants<NonVisualDrawingProperties>().FirstOrDefault()
            ?? element.Descendants().FirstOrDefault(e => e.LocalName == "cNvPr");
        return GetCnvPrAttribute(cNvPr, "name");
    }

    private static string? GetCnvPrAttribute(OpenXmlElement? cNvPr, string attribute)
    {
        if (cNvPr == null) return null;
        var match = Regex.Match(cNvPr.OuterXml, $@"\b{Regex.Escape(attribute)}\s*=\s*""([^""]*)""");
        return match.Success ? match.Groups[1].Value : null;
    }

    /// <summary>
    /// Skips shapes hidden in the selection pane. PowerPoint writes <c>show="0"</c> on
    /// <c>p:cNvPr</c>; the schema also allows <c>hidden="1"</c> — both are honored.
    /// </summary>
    private static bool IsHidden(OpenXmlElement element)
    {
        var cNvPr = element.Descendants<NonVisualDrawingProperties>().FirstOrDefault()
            ?? element.Descendants().FirstOrDefault(e => e.LocalName == "cNvPr");
        if (cNvPr == null) return false;
        var xml = cNvPr.OuterXml;
        return Regex.IsMatch(xml, @"\bshow\s*=\s*""0""") || Regex.IsMatch(xml, @"\bhidden\s*=\s*""1""");
    }

    /// <summary>Escapes inline Markdown syntax so extracted text renders literally.</summary>
    private static string EscapeText(string text)
    {
        var sb = new StringBuilder(text.Length + 8);
        foreach (var c in text)
        {
            switch (c)
            {
                case '\\': sb.Append("\\\\"); break;
                case '*': sb.Append("\\*"); break;
                case '_': sb.Append("\\_"); break;
                case '~': sb.Append("\\~"); break;
                case '`': sb.Append("\\`"); break;
                case '[': sb.Append("\\["); break;
                case ']': sb.Append("\\]"); break;
                default: sb.Append(c); break;
            }
        }
        return sb.ToString();
    }

    /// <summary>
    /// Escapes a full line: inline characters plus a leading block marker (#, &gt;, -,
    /// digits+.) that Markdown would otherwise re-interpret as a heading/list/quote.
    /// </summary>
    private static string EscapeLine(string text)
        => EscapeParagraphLead(EscapeText(text));

    private static string EscapeParagraphLead(string text)
    {
        if (text.Length > 0 && (text[0] is '#' or '>' or '-' or '+'))
            return "\\" + text;
        if (Regex.IsMatch(text, @"^\d+\.\s"))
            return "\\" + text;
        return text;
    }

    private static string EscapeTableCell(string text)
        => EscapeLine(text).Replace("|", "\\|");

    /// <summary>
    /// Escapes only the characters that would break the link syntax inside an image's
    /// <c>![alt](url)</c>: backslash and the closing bracket. Alt text is otherwise kept
    /// verbatim (e.g. <c>cover_photo.jpg</c> stays readable).
    /// </summary>
    private static string EscapeAltText(string text)
        => text.Replace("\\", "\\\\").Replace("]", "\\]");

    /// <summary>Sanitizes text embedded in an HTML comment (never allow "--" or ">").</summary>
    private static string EscapeComment(string text)
        => text.Replace("-", "\u2212").Replace(">", "&gt;");

    private static string Capitalize(string value)
        => value.Length > 1
            ? char.ToUpperInvariant(value[0]) + value.Substring(1)
            : value.ToUpperInvariant();

    /// <summary>A single outline paragraph extracted from a text body.</summary>
    private sealed record TextParagraph(string Text, int Level, bool HasBullet, bool HasNoBullet);
}
