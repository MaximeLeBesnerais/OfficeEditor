using DocumentFormat.OpenXml.Packaging;

namespace DocxEditor.Core.Generation.Assets;

/// <summary>
/// Image media types the DOCX generation pipeline can embed. Detection is content-based
/// (magic bytes / markup sniffing, see <see cref="ImageFormatSniffer"/>), never
/// filename- or MIME-declaration based, so a mislabeled file is classified by what it
/// actually is. The set is closed: unsupported payloads fail loudly instead of being
/// guessed.
/// </summary>
public enum DocxImageMediaType
{
    /// <summary>Portable Network Graphics (image/png).</summary>
    Png,

    /// <summary>Joint Photographic Experts Group (image/jpeg).</summary>
    Jpeg,

    /// <summary>Graphics Interchange Format (image/gif).</summary>
    Gif,

    /// <summary>Windows bitmap (image/bmp).</summary>
    Bmp,

    /// <summary>Tagged Image File Format (image/tiff).</summary>
    Tiff,

    /// <summary>
    /// Scalable Vector Graphics (image/svg+xml). Word renders the blip when the host
    /// supports SVG; packaging it requires only the OpenXML SVG image-part content type.
    /// </summary>
    Svg
}

/// <summary>MIME content type, extension and OpenXML part mappings for <see cref="DocxImageMediaType"/>.</summary>
public static class DocxImageMediaTypes
{
    /// <summary>The MIME content type written into the package part for a media type.</summary>
    public static string GetContentType(DocxImageMediaType type) => type switch
    {
        DocxImageMediaType.Png => "image/png",
        DocxImageMediaType.Jpeg => "image/jpeg",
        DocxImageMediaType.Gif => "image/gif",
        DocxImageMediaType.Bmp => "image/bmp",
        DocxImageMediaType.Tiff => "image/tiff",
        DocxImageMediaType.Svg => "image/svg+xml",
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, "Unknown image media type.")
    };

    /// <summary>The conventional file extension (including the leading dot) for a media type.</summary>
    public static string GetExtension(DocxImageMediaType type) => type switch
    {
        DocxImageMediaType.Png => ".png",
        DocxImageMediaType.Jpeg => ".jpeg",
        DocxImageMediaType.Gif => ".gif",
        DocxImageMediaType.Bmp => ".bmp",
        DocxImageMediaType.Tiff => ".tiff",
        DocxImageMediaType.Svg => ".svg",
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, "Unknown image media type.")
    };

    /// <summary>The OpenXML part-type descriptor carrying the part's content type.</summary>
    public static PartTypeInfo ToOpenXmlPartType(DocxImageMediaType type) => type switch
    {
        DocxImageMediaType.Png => ImagePartType.Png,
        DocxImageMediaType.Jpeg => ImagePartType.Jpeg,
        DocxImageMediaType.Gif => ImagePartType.Gif,
        DocxImageMediaType.Bmp => ImagePartType.Bmp,
        DocxImageMediaType.Tiff => ImagePartType.Tiff,
        DocxImageMediaType.Svg => ImagePartType.Svg,
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, "Unknown image media type.")
    };

    /// <summary>
    /// Parses an MIME content type string (e.g. "image/png" or a data-URI media type with
    /// parameters such as "image/png;base64"). Returns null when unrecognized. This is only
    /// a hint: byte sniffing is authoritative.
    /// </summary>
    public static DocxImageMediaType? ParseContentType(string? contentType)
    {
        if (string.IsNullOrWhiteSpace(contentType))
        {
            return null;
        }
        var normalized = contentType.Split(';')[0].Trim().ToLowerInvariant();
        return normalized switch
        {
            "image/png" => DocxImageMediaType.Png,
            "image/jpeg" or "image/jpg" => DocxImageMediaType.Jpeg,
            "image/gif" => DocxImageMediaType.Gif,
            "image/bmp" or "image/x-ms-bmp" => DocxImageMediaType.Bmp,
            "image/tiff" => DocxImageMediaType.Tiff,
            "image/svg+xml" => DocxImageMediaType.Svg,
            _ => null
        };
    }

    /// <summary>
    /// Parses a file extension (or a whole file name/path) into a media type hint. This is
    /// only a hint: byte sniffing is authoritative.
    /// </summary>
    public static DocxImageMediaType? ParseExtension(string? fileNameOrPath)
    {
        if (string.IsNullOrWhiteSpace(fileNameOrPath))
        {
            return null;
        }
        var extension = Path.GetExtension(fileNameOrPath).ToLowerInvariant();
        return extension switch
        {
            ".png" => DocxImageMediaType.Png,
            ".jpg" or ".jpeg" => DocxImageMediaType.Jpeg,
            ".gif" => DocxImageMediaType.Gif,
            ".bmp" => DocxImageMediaType.Bmp,
            ".tif" or ".tiff" => DocxImageMediaType.Tiff,
            ".svg" => DocxImageMediaType.Svg,
            _ => null
        };
    }
}
