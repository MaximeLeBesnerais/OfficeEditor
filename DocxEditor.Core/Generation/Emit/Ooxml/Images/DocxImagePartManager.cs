using DocumentFormat.OpenXml.Packaging;
using DocxEditor.Core.Generation.Assets;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Images;

/// <summary>
/// Adds and deduplicates image parts on a <see cref="MainDocumentPart"/>. Images are keyed
/// by SHA-256 content hash: registering the same payload twice returns the same
/// <see cref="RegisteredImagePart"/> (and relationship id), so the package embeds one part
/// per unique image. Relationship ids, part names and docPr ids are deterministic given the
/// same input, payload bytes are fed straight from the asset buffer (no copy, no mutation,
/// no temp files) and the caller's buffers are never touched.
/// </summary>
public sealed class DocxImagePartManager
{
    private readonly MainDocumentPart _mainDocumentPart;
    private readonly Dictionary<string, RegisteredImagePart> _byContentHash = new(StringComparer.Ordinal);
    private uint _nextDocPrId = 1;

    /// <summary>Creates a manager over the given main document part.</summary>
    public DocxImagePartManager(MainDocumentPart mainDocumentPart)
    {
        ArgumentNullException.ThrowIfNull(mainDocumentPart);
        _mainDocumentPart = mainDocumentPart;
    }

    /// <summary>
    /// Registers an image asset, adding its part on first use and returning the shared
    /// <see cref="RegisteredImagePart"/> on every subsequent call for the same content hash.
    /// </summary>
    public RegisteredImagePart Register(ImageAsset asset)
    {
        ArgumentNullException.ThrowIfNull(asset);
        if (_byContentHash.TryGetValue(asset.ContentHash, out var existing))
        {
            return existing;
        }

        var relationshipId = AllocateRelationshipId(asset.ContentHash);
        var imagePart = _mainDocumentPart.AddImagePart(
            DocxImageMediaTypes.GetContentType(asset.MediaType), relationshipId);
        imagePart.FeedData(new MemoryStream(asset.Bytes, writable: false));

        var registered = new RegisteredImagePart
        {
            RelationshipId = relationshipId,
            ContentHash = asset.ContentHash,
            MediaType = asset.MediaType,
            PartName = imagePart.Uri?.ToString()
                ?? $"image/{asset.ContentHash[..12]}{DocxImageMediaTypes.GetExtension(asset.MediaType)}",
            Width = asset.Width,
            Height = asset.Height
        };
        _byContentHash[asset.ContentHash] = registered;
        return registered;
    }

    /// <summary>Returns the registered part for a content hash, or null.</summary>
    public RegisteredImagePart? Find(string contentHash)
    {
        ArgumentNullException.ThrowIfNull(contentHash);
        return _byContentHash.GetValueOrDefault(contentHash);
    }

    /// <summary>
    /// Allocates the next deterministic docPr id. Call once per placed picture (inline or
    /// positioned) so ids stay unique across the whole document.
    /// </summary>
    public uint NextDocPrId() => _nextDocPrId++;

    /// <summary>Every registered part, keyed by content hash, in registration order.</summary>
    public IReadOnlyDictionary<string, RegisteredImagePart> RegisteredImagesByHash => _byContentHash;

    /// <summary>
    /// Deterministic relationship id derived from the content hash, bumped with a suffix if
    /// the template already carries a relationship with that id (so the id stays stable per
    /// input and can never collide).
    /// </summary>
    private string AllocateRelationshipId(string contentHash)
    {
        var baseId = $"image{contentHash[..12]}";
        var candidate = baseId;
        var suffix = 2;
        while (_mainDocumentPart.Parts.Any(p =>
                   string.Equals(p.RelationshipId, candidate, StringComparison.Ordinal)))
        {
            candidate = $"{baseId}_{suffix++}";
        }
        return candidate;
    }
}
