using DocxEditor.Core.Generation.Emit.Ooxml.Images;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Positioned;

/// <summary>
/// Resolved placement of a positioned image: the relationship id to embed plus the
/// canonical picture geometry (extents/offset/crop). Produced by an
/// <see cref="IPositionedImageResolver"/> and consumed by the positioned picture emitter.
/// </summary>
public sealed record PositionedImagePlacement
{
    /// <summary>Relationship id (<c>r:embed</c>) of the registered image part.</summary>
    public required string EmbedId { get; init; }

    /// <summary>Canonical display geometry for the picture's <c>a:xfrm</c> and <c>a:srcRect</c>.</summary>
    public required ResolvedImageGeometry Geometry { get; init; }
}
