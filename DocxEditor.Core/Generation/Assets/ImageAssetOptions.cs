namespace DocxEditor.Core.Generation.Assets;

/// <summary>
/// Asset-safety knobs for the image pipeline (an asset option, not a global hostile-document
/// policy — the caller's deployment boundary handles the latter, see SECURITY.md). Every
/// limit is a practical default and overridable per load.
/// </summary>
public sealed record ImageAssetOptions
{
    /// <summary>Default cap on encoded/decoded bytes: 25 MiB.</summary>
    public const long DefaultMaxEncodedBytes = 25 * 1024 * 1024;

    /// <summary>Default cap on either pixel dimension: 16384 px.</summary>
    public const int DefaultMaxPixelDimension = 16384;

    /// <summary>Default cap on total pixel area: 16384 × 16384.</summary>
    public const long DefaultMaxPixelArea = 268_435_456;

    /// <summary>Defaults: 25 MiB encoded, 16384 px per side, 16384² area, all six formats enabled.</summary>
    public static ImageAssetOptions Default { get; } = new();

    /// <summary>
    /// Maximum encoded (and decoded) payload size in bytes. Data-URI payloads are length-
    /// checked before base64/percent decoding and files are length-checked before reading,
    /// so an oversized source is rejected without ever materializing its full contents.
    /// </summary>
    public long MaxEncodedBytes { get; init; } = DefaultMaxEncodedBytes;

    /// <summary>Maximum width <i>and</i> height in pixels, read from the header by sniffing.</summary>
    public int MaxPixelDimension { get; init; } = DefaultMaxPixelDimension;

    /// <summary>Maximum width × height (pixel area).</summary>
    public long MaxPixelArea { get; init; } = DefaultMaxPixelArea;

    /// <summary>
    /// Optional restriction on which <see cref="DocxImageMediaType"/>s are embeddable. Null
    /// (default) enables all supported formats.
    /// </summary>
    public IReadOnlySet<DocxImageMediaType>? SupportedMediaTypes { get; init; }

    /// <summary>Validates the option values, throwing <see cref="ArgumentOutOfRangeException"/> for non-positive limits.</summary>
    public static void Validate(ImageAssetOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);
        if (options.MaxEncodedBytes <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(options.MaxEncodedBytes), options.MaxEncodedBytes, "Maximum encoded bytes must be positive.");
        }
        if (options.MaxPixelDimension <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(options.MaxPixelDimension), options.MaxPixelDimension, "Maximum pixel dimension must be positive.");
        }
        if (options.MaxPixelArea <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(options.MaxPixelArea), options.MaxPixelArea, "Maximum pixel area must be positive.");
        }
    }
}
