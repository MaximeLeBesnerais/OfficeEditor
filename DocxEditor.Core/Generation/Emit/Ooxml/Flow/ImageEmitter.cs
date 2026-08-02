using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Generation.Emit.Ooxml.Images;
using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Flow;

/// <summary>Emits flow images through the shared asset loader and image-part manager.</summary>
internal static class ImageEmitter
{
    public static void EmitInlineImage(
        OoxmlEmitContext context,
        OpenXmlCompositeElement container,
        ImageElement image,
        string path)
    {
        var owningPart = context.GetOwningPart(container);
        var (asset, part) = context.Images.Resolve(owningPart, image.Source, path);
        var geometry = ResolvedImageGeometry.Resolve(
            asset,
            image.Fit,
            image.Crop,
            image.WidthPt,
            image.HeightPt);
        var drawing = InlinePictureXmlFactory.BuildInline(
            part,
            geometry,
            context.Images.NextDocPrId(owningPart),
            image.Alt);
        container.Append(new Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(drawing)));
    }
}
