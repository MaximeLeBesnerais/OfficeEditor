using OfficeEditor.Core.Exceptions;

namespace DocxEditor.Core.Generation.Assets;

/// <summary>Why an image source could not be resolved or loaded into an asset.</summary>
public enum ImageSourceErrorCode
{
    /// <summary>The source string is empty or whitespace.</summary>
    SourceIsEmpty,

    /// <summary>
    /// The source is an HTTP(S) URL. Remote fetching is intentionally disabled (no network,
    /// no SSRF): only data URIs and local files are acceptable.
    /// </summary>
    RemoteSourceNotAllowed,

    /// <summary>The source parses as an absolute URI with a scheme other than data/http(s)/file.</summary>
    UnsupportedUriScheme,

    /// <summary>The source looks like a data URI but is structurally invalid (no payload
    /// separator, malformed base64, …).</summary>
    InvalidDataUri,

    /// <summary>The source is an absolute filesystem path but absolute paths are not enabled.</summary>
    AbsolutePathNotAllowed,

    /// <summary>
    /// The resolved path escapes the configured allowed root (lexically, or canonically
    /// through symlinks when the path exists).
    /// </summary>
    OutsideAllowedRoot,

    /// <summary>The resolved local file does not exist (or is not a file).</summary>
    FileNotFound,

    /// <summary>The file exists but could not be read under the current permissions.</summary>
    PathNotReadable,

    /// <summary>
    /// The encoded payload or file exceeds the configured maximum encoded bytes. Checked
    /// before decoding so an oversized hostile payload is never fully materialized.
    /// </summary>
    EncodedBytesExceededLimit,

    /// <summary>
    /// The bytes do not match any supported image format (or the format is disabled by the
    /// asset options). The media type is inferred from bytes, never from the filename.
    /// </summary>
    UnsupportedMediaType,

    /// <summary>
    /// The format is recognized but the intrinsic pixel dimensions could not be read from
    /// the header, so it cannot be sized.
    /// </summary>
    DimensionsUnavailable,

    /// <summary>The pixel width, height, or area exceeds the configured limits.</summary>
    DimensionsExceedLimits,

    /// <summary>The payload is empty or structurally invalid (e.g. a truncated header).</summary>
    InvalidImageData
}

/// <summary>
/// Typed domain error raised for every image-source failure, carrying a stable
/// <see cref="Code"/> callers can branch on and a human-readable message. Derives from
/// <see cref="OfficeEditorException"/> per project convention.
/// </summary>
public sealed class ImageSourceException : OfficeEditorException
{
    /// <summary>Stable machine-readable reason.</summary>
    public ImageSourceErrorCode Code { get; }

    public ImageSourceException(ImageSourceErrorCode code, string message)
        : base(message)
    {
        Code = code;
    }

    public ImageSourceException(ImageSourceErrorCode code, string message, Exception innerException)
        : base(message, innerException)
    {
        Code = code;
    }
}
