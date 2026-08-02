namespace DocxEditor.Core.Generation.Assets;

/// <summary>Where an <see cref="ImageSourceResolution"/> gets its payload bytes from.</summary>
public enum ImageSourceKind
{
    /// <summary>An inline <c>data:</c> URI whose encoded payload is decoded by the loader.</summary>
    DataUri,

    /// <summary>A local filesystem path read by the loader under the path policy.</summary>
    LocalFile
}

/// <summary>
/// A resolved image source: the outcome of <see cref="ImageSourceResolver.Resolve"/>. It
/// describes <em>where</em> the bytes come from without reading or decoding them — the
/// loader performs length checks, sniffing, and hashing on the single decoded buffer.
/// </summary>
public sealed record ImageSourceResolution
{
    /// <summary>The kind of source.</summary>
    public required ImageSourceKind Kind { get; init; }

    /// <summary>The original source string as supplied by the caller.</summary>
    public required string Source { get; init; }

    /// <summary>
    /// Media type declared by the source (data-URI media type, or the local file extension).
    /// Advisory only — byte sniffing is authoritative. Null when the source declares nothing.
    /// </summary>
    public string? DeclaredMediaType { get; init; }

    /// <summary>For <see cref="ImageSourceKind.DataUri"/>: the raw (still encoded) payload.</summary>
    public string? Payload { get; init; }

    /// <summary>For <see cref="ImageSourceKind.DataUri"/>: whether the payload is base64.</summary>
    public bool IsBase64 { get; init; }

    /// <summary>For <see cref="ImageSourceKind.LocalFile"/>: the resolved, validated filesystem path.</summary>
    public string? FilePath { get; init; }

    /// <summary>A human-friendly, deterministic display name (file name, or null for data URIs).</summary>
    public string? FileName { get; init; }
}
