using DocumentFormat.OpenXml.Packaging;
using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Positioned;

/// <summary>
/// Seam for the positioned image anchor shell. The integrating pipeline owns an
/// image-part manager (part creation, asset path/data-URI loading, deduplication) and
/// supplies both the relationship id to embed and, optionally, the source image's
/// natural pixel size. This emitter never reads image bytes and never creates image
/// parts itself — it only builds the <c>wp:anchor</c>/<c>p:pic</c> shell around whatever
/// the resolver hands back.
/// </summary>
public interface IPositionedImageResolver
{
    /// <summary>
    /// Resolves the image relationship id (<c>r:embed</c>) for a positioned image, adding
    /// an <see cref="ImagePart"/> under <paramref name="owningPart"/> if the manager does
    /// not already have one for <see cref="PositionedImage.Source"/>. Returns null when the
    /// image cannot be provided; the emitter then skips the element with an explicit
    /// warning.
    /// </summary>
    /// <param name="owningPart">The part that owns the image relationship (main document
    /// part, header part or footer part).</param>
    /// <param name="image">The positioned image model element.</param>
    string? ResolveEmbedId(OpenXmlPart owningPart, PositionedImage image);

    /// <summary>
    /// Tries to report the source image's natural pixel dimensions for the given source,
    /// used to compute cover (<see cref="ImageFitMode.Fill"/>) and
    /// <see cref="ImageFitMode.Contain"/> geometry. Return false when unknown — those fit
    /// modes then degrade to stretch with an explicit warning.
    /// </summary>
    bool TryGetNaturalPixelSize(string source, out int width, out int height);
}
