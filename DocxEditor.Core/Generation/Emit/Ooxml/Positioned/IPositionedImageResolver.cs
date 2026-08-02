using DocumentFormat.OpenXml.Packaging;
using DocxEditor.Core.Generation.Emit.Ooxml.Images;
using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Positioned;

/// <summary>
/// Seam for the positioned image anchor shell. The integrating pipeline owns an
/// image-part manager (part creation, asset path/data-URI loading, deduplication) and
/// supplies both the relationship id to embed and the canonical placement geometry
/// computed from the loaded asset. This emitter never reads image bytes and never creates
/// image parts itself — it only builds the <c>wp:anchor</c>/<c>p:pic</c> shell around
/// whatever the resolver hands back.
/// </summary>
public interface IPositionedImageResolver
{
    /// <summary>
    /// Resolves a positioned image into its embed relationship id (<c>r:embed</c>) and the
    /// canonical placement geometry (<c>a:xfrm</c> extents/offset + <c>a:srcRect</c> crop)
    /// produced by <see cref="ResolvedImageGeometry.Resolve"/> from the loaded asset's
    /// pixels and intrinsic DPI. Adds an <see cref="ImagePart"/> under
    /// <paramref name="owningPart"/> when the manager does not already have one for
    /// <see cref="PositionedImage.Source"/>. Returns null when the image cannot be provided;
    /// the emitter then skips the element with an explicit warning.
    /// </summary>
    /// <param name="owningPart">The part that owns the image relationship (main document
    /// part, header part or footer part).</param>
    /// <param name="image">The positioned image model element.</param>
    /// <param name="path">The JSON path of the image source, used to attribute asset warnings;
    /// null when the caller has no path.</param>
    PositionedImagePlacement? Resolve(OpenXmlPart owningPart, PositionedImage image, string? path);
}
