using DocumentFormat.OpenXml.Packaging;
using DocxEditor.Core.Generation.Assets;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Images;

/// <summary>
/// Adds and deduplicates image parts on a main document, header, or footer part. Images are keyed
/// by SHA-256 content hash: registering the same payload twice returns the same
/// <see cref="RegisteredImagePart"/> (and relationship id), so the package embeds one part
/// per unique image for this relationship owner. Relationship ids, part names and docPr ids
/// are deterministic given the same input, and payload bytes are fed from the asset-owned
/// read-only buffer without temp files.
/// </summary>
public sealed class DocxImagePartManager
{
    private readonly OpenXmlPart _owningPart;
    private readonly Dictionary<string, RegisteredImagePart> _byContentHash = new(StringComparer.Ordinal);
    private uint _nextDocPrId;

    /// <summary>Creates a manager over the given main document part.</summary>
    public DocxImagePartManager(MainDocumentPart mainDocumentPart)
    {
        ArgumentNullException.ThrowIfNull(mainDocumentPart);
        _owningPart = mainDocumentPart;
        _nextDocPrId = ComputeNextDocPrId();
    }

    /// <summary>Creates a manager over a header part.</summary>
    public DocxImagePartManager(HeaderPart headerPart)
    {
        ArgumentNullException.ThrowIfNull(headerPart);
        _owningPart = headerPart;
        _nextDocPrId = ComputeNextDocPrId();
    }

    /// <summary>Creates a manager over a footer part.</summary>
    public DocxImagePartManager(FooterPart footerPart)
    {
        ArgumentNullException.ThrowIfNull(footerPart);
        _owningPart = footerPart;
        _nextDocPrId = ComputeNextDocPrId();
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
        var contentType = DocxImageMediaTypes.GetContentType(asset.MediaType);
        var imagePart = _owningPart switch
        {
            MainDocumentPart mainPart => mainPart.AddImagePart(contentType, relationshipId),
            HeaderPart headerPart => headerPart.AddImagePart(contentType, relationshipId),
            FooterPart footerPart => footerPart.AddImagePart(contentType, relationshipId),
            _ => throw new NotSupportedException(
                $"Image relationships are not supported for part type '{_owningPart.GetType().Name}'.")
        };
        using var assetStream = asset.OpenRead();
        imagePart.FeedData(assetStream);

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
    /// positioned) so ids stay unique within the relationship-owning part.
    /// </summary>
    public uint NextDocPrId() => _nextDocPrId++;

    /// <summary>Every registered part, keyed by content hash, in registration order.</summary>
    public IReadOnlyDictionary<string, RegisteredImagePart> RegisteredImagesByHash => _byContentHash;

    private uint ComputeNextDocPrId()
    {
        uint maxId = 0;
        if (_owningPart.RootElement is { } root)
        {
            foreach (var element in root.Descendants())
            {
                if (element.LocalName is not ("docPr" or "cNvPr"))
                {
                    continue;
                }
                var id = element.GetAttribute("id", string.Empty).Value;
                if (uint.TryParse(id, System.Globalization.CultureInfo.InvariantCulture, out var parsed))
                {
                    maxId = Math.Max(maxId, parsed);
                }
            }
        }
        return maxId + 1;
    }

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
        while (_owningPart.Parts.Any(p =>
                   string.Equals(p.RelationshipId, candidate, StringComparison.Ordinal)))
        {
            candidate = $"{baseId}_{suffix++}";
        }
        return candidate;
    }
}
