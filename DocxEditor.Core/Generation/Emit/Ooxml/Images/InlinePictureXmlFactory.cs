using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Images;

/// <summary>
/// Builds the inline (<c>wp:inline</c>) DrawingML picture for a flow-tier image block. The
/// result is a ready-to-append <c>w:drawing</c> referencing a
/// <see cref="RegisteredImagePart"/> relationship. Names and docPr ids are deterministic.
/// </summary>
public static class InlinePictureXmlFactory
{
    /// <summary>
    /// Builds an inline picture drawing for a registered part and its resolved geometry.
    /// <paramref name="docPrId"/> should come from
    /// <see cref="DocxImagePartManager.NextDocPrId"/> so ids stay unique per document.
    /// </summary>
    public static Drawing BuildInline(
        RegisteredImagePart part,
        ResolvedImageGeometry geometry,
        uint docPrId,
        string? altText)
    {
        ArgumentNullException.ThrowIfNull(part);
        ArgumentNullException.ThrowIfNull(geometry);

        var inline = new Wp.Inline
        {
            DistanceFromTop = 0,
            DistanceFromBottom = 0,
            DistanceFromLeft = 0,
            DistanceFromRight = 0
        };
        inline.Append(new Wp.Extent { Cx = geometry.ExtentsCxEmu, Cy = geometry.ExtentsCyEmu });
        inline.Append(new Wp.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 });
        inline.Append(new Wp.DocProperties
        {
            Id = docPrId,
            Name = $"Picture {docPrId}",
            Description = altText
        });
        inline.Append(new Wp.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }));
        inline.Append(PictureGraphicBuilder.BuildGraphic(part, geometry, docPrId, altText, rotationSixtiethsOfDegree: 0));

        return new Drawing(inline);
    }
}
