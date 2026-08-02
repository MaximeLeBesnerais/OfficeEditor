using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Generation.Model;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Flow;

/// <summary>
/// Emits inline images through the <see cref="IImageContentResolver"/> seam. When no resolver
/// is wired — or a source cannot be resolved — an explicit placeholder paragraph is emitted
/// (italic gray, boxed) plus a warning, so inline images are never silently dropped. The
/// drawing construction below is the integration surface for the image workstream.
/// </summary>
internal static class ImageEmitter
{
    private const string PictureNamespace = "http://schemas.openxmlformats.org/drawingml/2006/picture";

    public static void EmitInlineImage(OoxmlEmitContext context, OpenXmlCompositeElement container, ImageElement image, string path)
    {
        if (context.ImageResolver is not { } resolver)
        {
            EmitPlaceholder(context, container, image, path,
                "no image resolver is wired on this branch; the inline image was emitted as a placeholder instead of being dropped.");
            return;
        }

        if (!resolver.TryResolveImage(image.Source, out var content, out var extension, out var naturalWidthPt, out var naturalHeightPt))
        {
            EmitPlaceholder(context, container, image, path,
                $"the image source '{image.Source}' could not be resolved; a placeholder was emitted instead.");
            return;
        }

        if (naturalWidthPt <= 0 || naturalHeightPt <= 0)
        {
            EmitPlaceholder(context, container, image, path,
                $"the image source '{image.Source}' resolved with non-positive natural dimensions; a placeholder was emitted instead.");
            return;
        }

        var (boxWidth, boxHeight) = ResolveDisplayBox(context, image, naturalWidthPt, naturalHeightPt, path);
        var fit = ImageFitGeometry.ComputeFit(image.Fit, naturalWidthPt, naturalHeightPt, boxWidth, boxHeight, image.Crop);

        var imagePart = context.MainPart.AddImagePart(ImageContentTypeFromExtension(extension));
        using (var stream = new MemoryStream(content))
        {
            imagePart.FeedData(stream);
        }

        var relationshipId = context.MainPart.GetIdOfPart(imagePart);
        var drawing = BuildInlineDrawing(context, image, relationshipId, fit);
        container.Append(new W.Paragraph(new W.Run(drawing)));
    }

    private static (double WidthPt, double HeightPt) ResolveDisplayBox(
        OoxmlEmitContext context,
        ImageElement image,
        double naturalWidthPt,
        double naturalHeightPt,
        string path)
    {
        var width = image.WidthPt;
        var height = image.HeightPt;
        switch (width, height)
        {
            case (null, null):
                return (naturalWidthPt, naturalHeightPt);
            case ({ } w, null):
                return (w, w * naturalHeightPt / naturalWidthPt);
            case (null, { } h):
                return (h * naturalWidthPt / naturalHeightPt, h);
            case ({ } w, { } h):
                return (w, h);
        }
    }

    /// <summary>
    /// Builds a <c>w:drawing</c> with an inline picture. Child order follows CT_Inline:
    /// extent, effectExtent, docPr, cNvGraphicFramePr, graphic.
    /// </summary>
    private static W.Drawing BuildInlineDrawing(
        OoxmlEmitContext context,
        ImageElement image,
        string relationshipId,
        ImageFitResult fit)
    {
        var cx = FormattingHelpers.Emus(fit.WidthPt);
        var cy = FormattingHelpers.Emus(fit.HeightPt);
        var id = (uint)context.NextDrawingId();
        var name = $"Image {id}";

        var blipFill = new Pic.BlipFill(new A.Blip { Embed = relationshipId });
        if (fit.SourceRect is { } sourceRect)
        {
            // CT_BlipFillProperties sequence: blip?, srcRect?, (tile|stretch)?.
            blipFill.Append(sourceRect);
        }
        blipFill.Append(new A.Stretch(new A.FillRectangle()));

        var docProperties = new Wp.DocProperties { Id = id, Name = name };
        if (image.Alt is not null)
        {
            docProperties.Description = image.Alt;
        }

        var inline = new Wp.Inline(
            new Wp.Extent { Cx = cx, Cy = cy },
            new Wp.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
            docProperties,
            new Wp.NonVisualGraphicFrameDrawingProperties(),
            new A.Graphic(new A.GraphicData(
                new Pic.Picture(
                    new Pic.NonVisualPictureProperties(
                        new Pic.NonVisualDrawingProperties { Id = id, Name = name },
                        new Pic.NonVisualPictureDrawingProperties()),
                    blipFill,
                    new Pic.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = 0, Y = 0 },
                            new A.Extents { Cx = cx, Cy = cy }),
                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
            {
                Uri = PictureNamespace
            }))
        {
            DistanceFromTop = 0,
            DistanceFromBottom = 0,
            DistanceFromLeft = 0,
            DistanceFromRight = 0
        };

        return new W.Drawing(inline);
    }

    private static void EmitPlaceholder(
        OoxmlEmitContext context,
        OpenXmlCompositeElement container,
        ImageElement image,
        string path,
        string? message)
    {
        context.Warn(path, message ?? $"the image source '{image.Source}' could not be rendered; a placeholder was emitted instead.");

        var paragraph = new W.Paragraph(
            new W.ParagraphProperties(
                new W.ParagraphBorders(
                    new W.TopBorder { Val = BorderValues.Single, Size = 4, Color = "D0D0D0" },
                    new W.LeftBorder { Val = BorderValues.Single, Size = 4, Color = "D0D0D0" },
                    new W.BottomBorder { Val = BorderValues.Single, Size = 4, Color = "D0D0D0" },
                    new W.RightBorder { Val = BorderValues.Single, Size = 4, Color = "D0D0D0" }),
                new W.Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F7F7F7" }),
            new W.Run(
                new W.RunProperties(
                    new Italic(),
                    new Color { Val = "808080" }),
                new W.Text($"[image: {image.Source}]") { Space = SpaceProcessingModeValues.Preserve }));

        container.Append(paragraph);
    }

    private static string ImageContentTypeFromExtension(string extension)
    {
        return extension.ToLowerInvariant() switch
        {
            ".png" => "image/png",
            ".gif" => "image/gif",
            ".bmp" => "image/bmp",
            ".tiff" or ".tif" => "image/tiff",
            ".svg" => "image/svg+xml",
            _ => "image/jpeg"
        };
    }
}
