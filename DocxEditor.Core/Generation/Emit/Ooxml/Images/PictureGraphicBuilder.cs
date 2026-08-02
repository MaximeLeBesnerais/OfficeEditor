using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Images;

/// <summary>
/// Shared DrawingML picture construction for the inline and positioned (anchored) image
/// emitters. Both wrap the same <c>pic:pic</c> subtree; the anchored variant adds rotation.
/// </summary>
internal static class PictureGraphicBuilder
{
    private const string PictureNamespaceUri = "http://schemas.openxmlformats.org/drawingml/2006/picture";

    public static A.Graphic BuildGraphic(
        RegisteredImagePart part,
        ResolvedImageGeometry geometry,
        uint docPrId,
        string? altText,
        long rotationSixtiethsOfDegree)
    {
        var name = $"Picture {docPrId}";
        var xfrm = new A.Transform2D(
            new A.Offset { X = geometry.OffsetXEmu, Y = geometry.OffsetYEmu },
            new A.Extents { Cx = geometry.ExtentsCxEmu, Cy = geometry.ExtentsCyEmu });
        if (rotationSixtiethsOfDegree != 0)
        {
            xfrm.Rotation = (int)rotationSixtiethsOfDegree;
        }

        var picture = new Pic.Picture(
            new Pic.NonVisualPictureProperties(
                new Pic.NonVisualDrawingProperties { Id = docPrId, Name = name, Description = altText },
                new Pic.NonVisualPictureDrawingProperties()),
            BuildBlipFill(part, geometry),
            new Pic.ShapeProperties(
                xfrm,
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

        return new A.Graphic(
            new A.GraphicData(picture) { Uri = PictureNamespaceUri });
    }

    /// <summary>
    /// CT_BlipFillProperties sequence: blip, optional srcRect, stretch. The crop (if any)
    /// must sit between a:blip and a:stretch to validate.
    /// </summary>
    private static Pic.BlipFill BuildBlipFill(RegisteredImagePart part, ResolvedImageGeometry geometry)
    {
        var blipFill = new Pic.BlipFill(new A.Blip { Embed = part.RelationshipId });
        if (geometry.SourceRect is { } srcRect)
        {
            blipFill.Append(new A.SourceRectangle
            {
                Left = srcRect.Left,
                Top = srcRect.Top,
                Right = srcRect.Right,
                Bottom = srcRect.Bottom
            });
        }
        blipFill.Append(new A.Stretch(new A.FillRectangle()));
        return blipFill;
    }
}
