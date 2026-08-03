using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PptxEditor.Core.Models;

namespace PptxEditor.Core.Services;

/// <summary>
/// Outcome of a single element-replacement operation.
/// </summary>
public readonly record struct PptxReplaceResult
{
    public bool Success { get; }

    public string? Error { get; }

    /// <summary>
    /// Non-fatal caveats (e.g. fit-mode fallbacks). Empty when nothing noteworthy happened.
    /// </summary>
    public IReadOnlyList<string> Warnings { get; }

    private PptxReplaceResult(bool success, string? error, IReadOnlyList<string>? warnings = null)
    {
        Success = success;
        Error = error;
        Warnings = warnings ?? Array.Empty<string>();
    }

    public static PptxReplaceResult Ok(IReadOnlyList<string>? warnings = null) => new(true, null, warnings);

    public static PptxReplaceResult Fail(string error) => new(false, error);

    public override string ToString() => Success ? "Success" : $"Failure: {Error}";
}

public class PptxElementReplacer
{
    // Guarded attribute read via regex on OuterXml: OpenXmlElement.GetAttribute()
    // can crash on unreliable OOXML attributes, so read them via regex on OuterXml
    // (canonical pattern: StyleResolver.GetAttributeValue).
    private static readonly Regex IdAttributePattern = new(
        @"\bid\s*=\s*""([^""]*)""", RegexOptions.Compiled);

    private const long DefaultTableWidth = 7200000;
    private const long DefaultRowHeight = 370840;

    public PptxReplaceResult ReplaceText(SlidePart slidePart, uint elementId, string newText)
    {
        ArgumentNullException.ThrowIfNull(slidePart);
        newText ??= string.Empty;

        if (elementId == 0)
            return PptxReplaceResult.Fail("Element id must be greater than zero.");

        var shapeTree = slidePart.Slide?.CommonSlideData?.ShapeTree;
        if (shapeTree == null)
            return PptxReplaceResult.Fail("Slide has no shape tree.");

        var matches = FindById<P.Shape>(shapeTree, elementId);
        if (matches.Count == 0)
            return PptxReplaceResult.Fail($"No shape with id {elementId} found on the slide.");
        if (matches.Count > 1)
            return PptxReplaceResult.Fail($"Multiple shapes with id {elementId} found on the slide; refusing to pick one arbitrarily.");

        var textBody = matches[0].TextBody;
        if (textBody == null)
            return PptxReplaceResult.Fail($"Shape with id {elementId} has no text body.");

        // Capture per-paragraph templates (paragraph properties + first run's run
        // properties) before clearing — run-level semantics, never naked runs.
        var templates = textBody.Elements<Drawing.Paragraph>()
            .Select(p => new ParagraphTemplate(
                p.ParagraphProperties?.CloneNode(true),
                p.Elements<Drawing.Run>().FirstOrDefault()?.RunProperties?.CloneNode(true)))
            .ToList();
        var fallbackPPr = templates.Select(t => t.ParagraphProperties).FirstOrDefault(p => p != null);
        var fallbackRPr = templates.Select(t => t.RunProperties).FirstOrDefault(r => r != null);

        textBody.RemoveAllChildren<Drawing.Paragraph>();

        // Multi-line text becomes multiple paragraphs; each line reuses the
        // template of the paragraph at the same index (last one when the
        // replacement has more lines than the original).
        var lines = newText.Replace("\r\n", "\n").Split('\n');
        for (var i = 0; i < lines.Length; i++)
        {
            var template = templates.Count == 0
                ? ParagraphTemplate.Empty
                : templates[Math.Min(i, templates.Count - 1)];

            var paragraph = new Drawing.Paragraph();
            var pPr = template.ParagraphProperties ?? fallbackPPr;
            if (pPr != null)
                paragraph.Append(pPr.CloneNode(true));

            var run = new Drawing.Run();
            var rPr = template.RunProperties ?? fallbackRPr;
            if (rPr != null)
                run.Append(rPr.CloneNode(true));
            run.Append(CreateText(lines[i]));
            paragraph.Append(run);

            textBody.Append(paragraph);
        }

        return PptxReplaceResult.Ok();
    }

    public PptxReplaceResult ReplaceTableData(SlidePart slidePart, uint elementId, List<List<string>> newData)
    {
        ArgumentNullException.ThrowIfNull(slidePart);
        ArgumentNullException.ThrowIfNull(newData);

        if (elementId == 0)
            return PptxReplaceResult.Fail("Element id must be greater than zero.");

        var shapeTree = slidePart.Slide?.CommonSlideData?.ShapeTree;
        if (shapeTree == null)
            return PptxReplaceResult.Fail("Slide has no shape tree.");

        var matches = FindById<P.GraphicFrame>(shapeTree, elementId);
        if (matches.Count == 0)
            return PptxReplaceResult.Fail($"No graphic frame with id {elementId} found on the slide.");
        if (matches.Count > 1)
            return PptxReplaceResult.Fail($"Multiple graphic frames with id {elementId} found on the slide; refusing to pick one arbitrarily.");

        var graphicData = matches[0].Graphic?.GraphicData;
        if (graphicData == null)
            return PptxReplaceResult.Fail($"Graphic frame with id {elementId} has no graphic data.");

        var existingTable = graphicData.Elements<Drawing.Table>().FirstOrDefault();
        if (existingTable == null)
            return PptxReplaceResult.Fail($"Graphic frame with id {elementId} does not contain a table.");

        if (newData.Count == 0 || newData[0].Count == 0)
            return PptxReplaceResult.Fail("Table data must contain at least one row and one column.");

        var numCols = newData[0].Count;
        if (newData.Any(row => row.Count != numCols))
            return PptxReplaceResult.Fail("All table rows must have the same number of columns.");

        var existingRows = existingTable.Elements<Drawing.TableRow>().ToList();
        var dimensionsMatch = existingRows.Count == newData.Count
            && existingRows.All(row => row.Elements<Drawing.TableCell>().Count() == numCols);

        if (dimensionsMatch)
        {
            // In-place edit: preserves each cell's tcPr and run properties.
            for (var i = 0; i < newData.Count; i++)
            {
                var cells = existingRows[i].Elements<Drawing.TableCell>().ToList();
                for (var j = 0; j < numCols; j++)
                {
                    SetCellText(cells[j], newData[i][j]);
                }
            }
            return PptxReplaceResult.Ok();
        }

        // Dimensions differ → rebuild, but keep the original table properties
        // (tableStyleId, banding flags), grid widths and row heights. Only the
        // a:tbl element is swapped, so the frame's xfrm is untouched.
        var newTable = RebuildTable(existingTable, newData, numCols);
        graphicData.RemoveAllChildren<Drawing.Table>();
        graphicData.Append(newTable);
        return PptxReplaceResult.Ok();
    }

    public PptxReplaceResult ReplaceImage(SlidePart slidePart, uint elementId, string newImagePath)
    {
        ArgumentNullException.ThrowIfNull(slidePart);

        if (!File.Exists(newImagePath))
        {
            throw new FileNotFoundException($"Image not found: {newImagePath}");
        }

        using var stream = new FileStream(newImagePath, FileMode.Open, FileAccess.Read);
        return ReplaceImage(
            slidePart, elementId, stream, Path.GetExtension(newImagePath), ImageFitMode.Stretch);
    }

    /// <summary>
    /// Replaces the image behind a picture element and fits it per <paramref name="fit"/>.
    /// The frame's a:xfrm is never moved except for <see cref="ImageFitMode.Contain"/>.
    /// </summary>
    /// <param name="extension">File-extension hint (e.g. ".png") used to derive the part type.</param>
    /// <param name="crop">Explicit a:srcRect for <see cref="ImageFitMode.Crop"/>; null → treated as Fill.</param>
    public PptxReplaceResult ReplaceImage(SlidePart slidePart, uint elementId, Stream image, string extension, ImageFitMode fit, SourceRect? crop = null)
    {
        ArgumentNullException.ThrowIfNull(slidePart);
        ArgumentNullException.ThrowIfNull(image);
        extension ??= string.Empty;

        if (elementId == 0)
            return PptxReplaceResult.Fail("Element id must be greater than zero.");

        var shapeTree = slidePart.Slide?.CommonSlideData?.ShapeTree;
        if (shapeTree == null)
            return PptxReplaceResult.Fail("Slide has no shape tree.");

        var matches = FindById<P.Picture>(shapeTree, elementId);
        if (matches.Count == 0)
            return PptxReplaceResult.Fail($"No picture with id {elementId} found on the slide.");
        if (matches.Count > 1)
            return PptxReplaceResult.Fail($"Multiple pictures with id {elementId} found on the slide; refusing to pick one arbitrarily.");

        var blip = matches[0].BlipFill?.Blip;
        if (blip == null)
            return PptxReplaceResult.Fail($"Picture with id {elementId} has no blip fill.");

        var oldEmbedId = blip.Embed?.Value;
        if (string.IsNullOrEmpty(oldEmbedId))
            return PptxReplaceResult.Fail($"Picture with id {elementId} is not linked to an image part.");

        var oldImagePart = TryGetPartById(slidePart, oldEmbedId) as ImagePart;
        if (oldImagePart == null)
            return PptxReplaceResult.Fail($"Relationship '{oldEmbedId}' does not resolve to an image part on the slide.");

        // Buffer once: the bytes are fed to the new part and sniffed for dimensions.
        byte[] imageBytes;
        using (var buffer = new MemoryStream())
        {
            image.CopyTo(buffer);
            imageBytes = buffer.ToArray();
        }

        // Add the new part and retarget the blip BEFORE touching the old part.
        var imagePart = slidePart.AddImagePart(GetImagePartTypeFromExtension(extension));
        using (var partStream = new MemoryStream(imageBytes, writable: false))
        {
            imagePart.FeedData(partStream);
        }
        blip.Embed = slidePart.GetIdOfPart(imagePart);

        // Delete the old part only when nothing else references it — the same
        // ImagePart may be shared with other slides; deleting it then corrupts
        // the deck. The count includes this slide's (now stale) relationship,
        // so 1 means "only we referenced it".
        if (CountPartReferences(slidePart.OpenXmlPackage, oldImagePart) <= 1)
        {
            slidePart.DeletePart(oldImagePart);
        }

        var warnings = new List<string>();
        ApplyFit(matches[0], imageBytes, fit, crop, warnings);
        return PptxReplaceResult.Ok(warnings);
    }

    private static void SetCellText(Drawing.TableCell cell, string value)
    {
        var textBody = cell.TextBody;
        var firstParagraph = textBody?.Elements<Drawing.Paragraph>().FirstOrDefault();
        var firstRun = firstParagraph?.Elements<Drawing.Run>().FirstOrDefault();

        if (textBody == null || firstParagraph == null || firstRun == null)
        {
            // No run-level template in this cell — rebuild a minimal text body,
            // leaving the cell's own properties (tcPr) untouched.
            if (textBody == null)
            {
                // a:txBody must precede a:tcPr in a:tc.
                textBody = new Drawing.TextBody(new Drawing.BodyProperties(), new Drawing.ListStyle());
                cell.InsertAt(textBody, 0);
            }
            else
            {
                textBody.RemoveAllChildren<Drawing.Paragraph>();
            }
            textBody.Append(new Drawing.Paragraph(new Drawing.Run(CreateText(value))));
            return;
        }

        // Keep the first paragraph (with its pPr) and the first run (with its
        // rPr); drop extra paragraphs/runs.
        foreach (var paragraph in textBody.Elements<Drawing.Paragraph>().Skip(1).ToList())
        {
            paragraph.Remove();
        }
        foreach (var child in firstParagraph.ChildElements.ToList())
        {
            if (child is Drawing.ParagraphProperties || ReferenceEquals(child, firstRun))
                continue;
            child.Remove();
        }

        var text = firstRun.Text ?? firstRun.AppendChild(new Drawing.Text());
        SetTextValue(text, value);
    }

    private static Drawing.Table RebuildTable(Drawing.Table existingTable, List<List<string>> newData, int numCols)
    {
        var tableProperties = existingTable.TableProperties?.CloneNode(true)
            ?? new Drawing.TableProperties { FirstRow = true };

        var originalWidths = existingTable.TableGrid?
            .Elements<Drawing.GridColumn>()
            .Select(column => column.Width?.Value ?? 0L)
            .ToList() ?? new List<long>();
        var widths = ComputeColumnWidths(originalWidths, numCols);

        var originalHeights = existingTable.Elements<Drawing.TableRow>()
            .Select(row => row.Height?.Value ?? DefaultRowHeight)
            .ToList();

        var newTable = new Drawing.Table();
        newTable.Append(tableProperties);
        newTable.Append(new Drawing.TableGrid(widths.Select(w => new Drawing.GridColumn { Width = w })));

        for (var i = 0; i < newData.Count; i++)
        {
            var height = originalHeights.Count > 0
                ? originalHeights[i % originalHeights.Count]
                : DefaultRowHeight;
            var tableRow = new Drawing.TableRow { Height = height };
            foreach (var cellText in newData[i])
            {
                tableRow.Append(new Drawing.TableCell(
                    new Drawing.TextBody(
                        new Drawing.BodyProperties(),
                        new Drawing.ListStyle(),
                        new Drawing.Paragraph(new Drawing.Run(CreateText(cellText)))),
                    new Drawing.TableCellProperties()));
            }
            newTable.Append(tableRow);
        }

        return newTable;
    }

    private static List<long> ComputeColumnWidths(List<long> originalWidths, int numCols)
    {
        if (originalWidths.Count == 0 || originalWidths.Any(w => w <= 0))
        {
            return Enumerable.Repeat(DefaultTableWidth / numCols, numCols).ToList();
        }

        // Reuse the original column widths cyclically, then normalize so the
        // total table width is preserved.
        var total = originalWidths.Sum();
        var widths = Enumerable.Range(0, numCols)
            .Select(c => originalWidths[c % originalWidths.Count])
            .ToList();
        var current = widths.Sum();
        if (current != total)
        {
            long applied = 0;
            for (var i = 0; i < numCols; i++)
            {
                widths[i] = widths[i] * total / current;
                applied += widths[i];
            }
            widths[^1] += total - applied; // absorb rounding drift
        }
        return widths;
    }

    private static Drawing.Text CreateText(string value)
    {
        var text = new Drawing.Text(value);
        SetSpacePreserve(text, value);
        return text;
    }

    private static void SetTextValue(Drawing.Text text, string value)
    {
        text.Text = value;
        SetSpacePreserve(text, value);
    }

    // Preserve leading/trailing whitespace in a:t (xml:space="preserve").
    private static void SetSpacePreserve(Drawing.Text text, string value)
    {
        if (value.Length > 0 && (value.StartsWith(' ') || value.EndsWith(' ')))
        {
            text.SetAttribute(new OpenXmlAttribute(
                "xml:space", "http://www.w3.org/XML/1998/namespace", "preserve"));
        }
    }

    // Extension → part-type switch mirrored from SlideBuilder.AddImage
    // (SlideBuilder.cs is owned by another workstream and must not be edited).
    private static PartTypeInfo GetImagePartTypeFromExtension(string extension)
    {
        var ext = extension.ToLowerInvariant();
        return ext switch
        {
            ".png" => ImagePartType.Png,
            ".gif" => ImagePartType.Gif,
            ".bmp" => ImagePartType.Bmp,
            ".tiff" or ".tif" => ImagePartType.Tiff,
            ".svg" => ImagePartType.Svg,
            _ => ImagePartType.Jpeg
        };
    }

    #region Fit modes (F7)

    private const int SrcRectScale = 100000; // 1/1000ths of a percent, spcPct family (rule 2)

    private static void ApplyFit(P.Picture picture, byte[] imageBytes, ImageFitMode fit, SourceRect? crop, List<string> warnings)
    {
        var blipFill = picture.BlipFill!;
        NormalizeTileToStretch(blipFill);

        // Crop without an explicit rect is Fill.
        var effectiveFit = fit == ImageFitMode.Crop && crop == null ? ImageFitMode.Fill : fit;

        if (effectiveFit == ImageFitMode.Stretch)
        {
            // Replacing a previously cropped picture must come out clean:
            // the old a:srcRect is replaced, never stacked.
            RemoveSourceRect(blipFill);
            return;
        }

        if (effectiveFit == ImageFitMode.Crop)
        {
            SetSourceRect(blipFill, crop!.Value);
            return;
        }

        // Fill and Contain need the new image's pixel dimensions and the frame extents.
        var dimensions = ImageHeaderSniffer.TryGetPixelDimensions(imageBytes);
        if (dimensions is not { } dims)
        {
            RemoveSourceRect(blipFill);
            warnings.Add(
                $"Image dimensions could not be determined (unsupported or unrecognized format); " +
                $"fit '{ToFitString(effectiveFit)}' fell back to 'stretch'.");
            return;
        }

        var xfrm = picture.ShapeProperties?.Transform2D;
        var cx = xfrm?.Extents?.Cx?.Value;
        var cy = xfrm?.Extents?.Cy?.Value;
        if (cx is not > 0 || cy is not > 0)
        {
            RemoveSourceRect(blipFill);
            warnings.Add(
                "Picture has no usable frame transform; " +
                $"fit '{ToFitString(effectiveFit)}' fell back to 'stretch'.");
            return;
        }

        if (effectiveFit == ImageFitMode.Fill)
        {
            ApplyFill(blipFill, dims.Width, dims.Height, cx.Value, cy.Value);
        }
        else
        {
            RemoveSourceRect(blipFill); // Contain never carries a srcRect
            ApplyContain(xfrm!, dims.Width, dims.Height, cx.Value, cy.Value);
        }
    }

    private static string ToFitString(ImageFitMode fit) => fit.ToString().ToLowerInvariant();

    /// <summary>
    /// Cover: center-crop the image so the visible region matches the frame aspect.
    /// Image wider than the frame → crop left/right; taller → crop top/bottom.
    /// </summary>
    private static void ApplyFill(P.BlipFill blipFill, int imageWidth, int imageHeight, long frameCx, long frameCy)
    {
        // Compare aspects in exact integer cross-products (imageW/imageH vs frameCx/frameCy).
        var imageCross = (long)imageWidth * frameCy;
        var frameCross = (long)imageHeight * frameCx;
        if (imageCross == frameCross)
        {
            RemoveSourceRect(blipFill); // aspects already match: no crop needed
            return;
        }

        int left = 0, top = 0, right = 0, bottom = 0;
        if (imageCross > frameCross)
        {
            var visibleFraction = (double)frameCross / imageCross;
            left = right = CropFractionToInt((1.0 - visibleFraction) / 2.0);
        }
        else
        {
            var visibleFraction = (double)imageCross / frameCross;
            top = bottom = CropFractionToInt((1.0 - visibleFraction) / 2.0);
        }

        if (left == 0 && top == 0 && right == 0 && bottom == 0)
        {
            RemoveSourceRect(blipFill);
            return;
        }
        SetSourceRect(blipFill, new SourceRect(left, top, right, bottom));
    }

    /// <summary>
    /// Fit inside: shrink the frame's a:ext around the frame's center so the frame
    /// aspect matches the image aspect; the image then fills the smaller frame exactly.
    /// </summary>
    private static void ApplyContain(Drawing.Transform2D xfrm, int imageWidth, int imageHeight, long cx, long cy)
    {
        var imageCross = (long)imageWidth * cy;
        var frameCross = (long)imageHeight * cx;
        if (imageCross == frameCross)
        {
            return; // frame already matches the image aspect
        }

        long newCx, newCy;
        if (imageCross > frameCross)
        {
            // Image wider than the frame → width-limited: keep cx, shrink cy.
            newCx = cx;
            newCy = (long)Math.Round(cx * (double)imageHeight / imageWidth, MidpointRounding.AwayFromZero);
        }
        else
        {
            // Image taller than the frame → height-limited: keep cy, shrink cx.
            newCx = (long)Math.Round(cy * (double)imageWidth / imageHeight, MidpointRounding.AwayFromZero);
            newCy = cy;
        }

        // Shrink around the center: shift the offset by half of each delta.
        var offset = xfrm.Offset;
        if (offset == null)
        {
            offset = new Drawing.Offset { X = 0, Y = 0 };
            xfrm.InsertAt(offset, 0); // a:off precedes a:ext in CT_Transform2D
        }
        offset.X = (offset.X?.Value ?? 0) + (cx - newCx) / 2;
        offset.Y = (offset.Y?.Value ?? 0) + (cy - newCy) / 2;

        var extents = xfrm.Extents!;
        extents.Cx = newCx;
        extents.Cy = newCy;
    }

    private static int CropFractionToInt(double fraction) =>
        (int)Math.Round(fraction * SrcRectScale, MidpointRounding.AwayFromZero);

    private static void SetSourceRect(P.BlipFill blipFill, SourceRect rect)
    {
        RemoveSourceRect(blipFill);
        var srcRect = new Drawing.SourceRectangle
        {
            Left = rect.Left,
            Top = rect.Top,
            Right = rect.Right,
            Bottom = rect.Bottom
        };
        // CT_BlipFillProperties sequence: blip?, srcRect?, (tile|stretch)? — srcRect follows the blip.
        if (blipFill.Blip != null)
        {
            blipFill.InsertAfter(srcRect, blipFill.Blip);
        }
        else
        {
            blipFill.InsertAt(srcRect, 0);
        }
    }

    private static void RemoveSourceRect(P.BlipFill blipFill) =>
        blipFill.RemoveAllChildren<Drawing.SourceRectangle>();

    /// <summary>Tiled blip fills don't compose with srcRect fit math; normalize to stretch first.</summary>
    private static void NormalizeTileToStretch(P.BlipFill blipFill)
    {
        var hasTile = blipFill.Elements<Drawing.Tile>().Any();
        if (hasTile)
        {
            blipFill.RemoveAllChildren<Drawing.Tile>();
        }
        if (hasTile || !blipFill.Elements<Drawing.Stretch>().Any())
        {
            // Same shape SlideBuilder.AddImage emits; appended last per the CT_BlipFillProperties sequence.
            blipFill.Append(new Drawing.Stretch(new Drawing.FillRectangle()));
        }
    }

    #endregion

    private static OpenXmlPart? TryGetPartById(OpenXmlPartContainer container, string relationshipId)
    {
        foreach (var pair in container.Parts)
        {
            if (pair.RelationshipId == relationshipId)
                return pair.OpenXmlPart;
        }
        return null;
    }

    private static int CountPartReferences(OpenXmlPackage package, OpenXmlPart target)
    {
        var count = 0;
        var visited = new HashSet<OpenXmlPartContainer>();
        var pending = new Stack<OpenXmlPartContainer>();
        pending.Push(package);
        while (pending.Count > 0)
        {
            var container = pending.Pop();
            if (!visited.Add(container))
                continue;
            foreach (var pair in container.Parts)
            {
                if (ReferenceEquals(pair.OpenXmlPart, target))
                    count++;
                pending.Push(pair.OpenXmlPart);
            }
        }
        return count;
    }

    private static List<T> FindById<T>(ShapeTree shapeTree, uint elementId) where T : OpenXmlElement
    {
        var matches = new List<T>();
        foreach (var element in shapeTree.ChildElements)
        {
            if (element is not T candidate)
                continue;

            var nvProperties = candidate switch
            {
                P.Shape shape => (OpenXmlElement?)shape.NonVisualShapeProperties,
                P.GraphicFrame frame => frame.NonVisualGraphicFrameProperties,
                P.Picture picture => picture.NonVisualPictureProperties,
                _ => null
            };
            if (nvProperties != null && GetElementId(nvProperties) == elementId)
                matches.Add(candidate);
        }
        return matches;
    }

    private static uint GetElementId(OpenXmlElement nvProperties)
    {
        // The cNvPr element might be parsed as OpenXmlUnknownElement due to namespace issues
        // in OpenXML SDK v3.5.1 when Drawing elements are nested inside Presentation elements
        var cnvPr = nvProperties.ChildElements.FirstOrDefault(e => e.LocalName == "cNvPr");
        if (cnvPr == null) return 0;

        var match = IdAttributePattern.Match(cnvPr.OuterXml);
        return match.Success && uint.TryParse(match.Groups[1].Value, out var parsedId)
            ? parsedId
            : 0;
    }

    private sealed record ParagraphTemplate(OpenXmlElement? ParagraphProperties, OpenXmlElement? RunProperties)
    {
        public static readonly ParagraphTemplate Empty = new(null, null);
    }
}
