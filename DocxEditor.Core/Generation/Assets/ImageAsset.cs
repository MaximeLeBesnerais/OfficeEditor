namespace DocxEditor.Core.Generation.Assets;

/// <summary>Category of an explicit asset warning (metadata the source carries but that
/// cannot be interpreted). Warnings never cause an image to be dropped.</summary>
public enum ImageAssetWarningCode
{
    /// <summary>The declared MIME type / extension conflicts with the sniffed bytes.</summary>
    DeclaredMediaTypeMismatch,

    /// <summary>The format carries no intrinsic DPI; natural size assumes 96 DPI.</summary>
    IntrinsicDpiUnavailable,

    /// <summary>The format carries resolution metadata but it is ambiguous (e.g. a TIFF
    /// resolution unit of "none" or an unparseable unit).</summary>
    ResolutionUnitUnspecified,

    /// <summary>The format declares a DPI with no usable units.</summary>
    DensityUnitsUnspecified,

    /// <summary>X and Y DPI differ significantly; display size is computed per-axis.</summary>
    AspectDpiInconsistent,

    /// <summary>SVG carries no raster fallback; hosts without SVG blip support will not render it.</summary>
    SvgRasterFallbackMissing
}

/// <summary>An explicit, non-fatal warning attached to a loaded <see cref="ImageAsset"/>.</summary>
public sealed record ImageAssetWarning(ImageAssetWarningCode Code, string Message);

/// <summary>
/// A fully loaded, validated image asset: the decoded payload bytes, content-detected
/// media type, intrinsic pixel dimensions, intrinsic DPI when available, the derived
/// natural display size in points, a SHA-256 content hash for deduplication, a
/// deterministic display name and any explicit warnings. Nothing here is mutable after
/// load; the bytes buffer is owned by the asset and never mutated by downstream consumers.
/// </summary>
public sealed record ImageAsset
{
    private byte[] _bytes = [];

    /// <summary>
    /// A copy of the decoded image bytes (from a data URI or local file). Mutating the
    /// returned array cannot change this asset's content or invalidate its hash.
    /// </summary>
    public required byte[] Bytes
    {
        get => (byte[])_bytes.Clone();
        init
        {
            ArgumentNullException.ThrowIfNull(value);
            _bytes = (byte[])value.Clone();
        }
    }

    /// <summary>Content-detected media type (bytes, never the filename).</summary>
    public required DocxImageMediaType MediaType { get; init; }

    /// <summary>Intrinsic width in pixels.</summary>
    public required int Width { get; init; }

    /// <summary>Intrinsic height in pixels.</summary>
    public required int Height { get; init; }

    /// <summary>Horizontal intrinsic DPI, when the format declares it.</summary>
    public double? DpiX { get; init; }

    /// <summary>Vertical intrinsic DPI, when the format declares it.</summary>
    public double? DpiY { get; init; }

    /// <summary>Natural display width in points (pixels ÷ DPI × 72; 96 DPI when unknown).</summary>
    public required double NaturalWidthPt { get; init; }

    /// <summary>Natural display height in points.</summary>
    public required double NaturalHeightPt { get; init; }

    /// <summary>Lowercase SHA-256 hex of the payload — the deduplication key.</summary>
    public required string ContentHash { get; init; }

    /// <summary>Deterministic display name (local file name, or image_{hash}.{ext}).</summary>
    public required string DisplayName { get; init; }

    /// <summary>The original source string.</summary>
    public required string Source { get; init; }

    /// <summary>Explicit warnings for unsupported/ambiguous metadata. Never silent drops.</summary>
    public IReadOnlyList<ImageAssetWarning> Warnings { get; init; } = [];

    /// <summary>Opens the asset-owned buffer for read-only package emission without exposing it.</summary>
    internal Stream OpenRead() => new MemoryStream(_bytes, writable: false);
}
