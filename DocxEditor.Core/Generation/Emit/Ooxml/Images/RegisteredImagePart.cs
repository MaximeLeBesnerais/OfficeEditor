using DocxEditor.Core.Generation.Assets;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Images;

/// <summary>
/// A content-deduplicated image part registered on the main document part. Multiple
/// drawings reference the same <see cref="RelationshipId"/> when their payload hashes to the
/// same <see cref="ContentHash"/>, so identical images never bloat the package.
/// </summary>
public sealed record RegisteredImagePart
{
    /// <summary>The relationship id drawings embed in <c>a:blip/@r:embed</c>.</summary>
    public required string RelationshipId { get; init; }

    /// <summary>Lowercase SHA-256 hex of the payload — the deduplication key.</summary>
    public required string ContentHash { get; init; }

    /// <summary>Content-detected media type.</summary>
    public required DocxImageMediaType MediaType { get; init; }

    /// <summary>The package part name (e.g. /word/media/image1.png).</summary>
    public required string PartName { get; init; }

    /// <summary>Intrinsic pixel width of the payload.</summary>
    public required int Width { get; init; }

    /// <summary>Intrinsic pixel height of the payload.</summary>
    public required int Height { get; init; }
}
