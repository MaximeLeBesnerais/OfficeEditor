using DocxEditor.Core.Generation.Assets;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Images;

/// <summary>
/// Positioned-picture content/relationship seam handed to the positioned-tier emitter: the
/// registered image part plus its resolved geometry (crop/extents/offset), the deterministic
/// docPr identity and the accessibility text. The positioned emitter turns this into a
/// <c>wp:anchor</c> drawing via
/// <see cref="PositionedPictureFactory.BuildAnchored(PositionedPictureContent, Model.PositionSpec)"/>
/// and wraps the result in <c>w:drawing</c> inside the run/paragraph of the anchored point.
/// </summary>
public sealed record PositionedPictureContent
{
    /// <summary>The deduplicated image part this picture references.</summary>
    public required RegisteredImagePart ImagePart { get; init; }

    /// <summary>The resolved crop/extents/offset geometry for this placement.</summary>
    public required ResolvedImageGeometry Geometry { get; init; }

    /// <summary>Deterministic docPr id (allocate via <see cref="DocxImagePartManager.NextDocPrId"/>).</summary>
    public required uint DocPrId { get; init; }

    /// <summary>Accessibility alt text for the picture.</summary>
    public string? AltText { get; init; }
}
