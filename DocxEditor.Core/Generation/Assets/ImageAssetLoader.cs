using System.Security.Cryptography;
using System.Text;

namespace DocxEditor.Core.Generation.Assets;

/// <summary>
/// Loads image assets from source strings: resolve (path policy) → read the payload once
/// (respecting the encoded-bytes limit before materializing) → sniff media type/dimensions/DPI
/// from a bounded header → enforce the configured pixel limits → hash. The buffer is decoded
/// exactly once; the resulting <see cref="ImageAsset"/> is immutable and its bytes are never
/// re-parsed.
/// </summary>
public sealed class ImageAssetLoader
{
    /// <summary>
    /// Loads an <see cref="ImageAsset"/> from a source string (data URI or local file).
    /// Throws <see cref="ImageSourceException"/> with a typed code for every rejection —
    /// never returns null and never silently drops an image.
    /// </summary>
    public ImageAsset Load(string source, ImageSourceOptions? sourceOptions = null, ImageAssetOptions? assetOptions = null)
    {
        ArgumentNullException.ThrowIfNull(source);
        var assetOptionsValue = assetOptions ?? ImageAssetOptions.Default;
        ImageAssetOptions.Validate(assetOptionsValue);

        var resolution = ImageSourceResolver.Resolve(source, sourceOptions);
        var (bytes, declaredMediaType, fileName) = ReadPayload(resolution, assetOptionsValue);
        return Build(bytes, declaredMediaType, fileName, resolution, assetOptionsValue);
    }

    /// <summary>
    /// Loads an asset from already-decoded bytes (e.g. an uploaded file buffer). Useful when
    /// the source has no string form; all sniffing, limits, hashing and warnings still apply.
    /// </summary>
    public ImageAsset LoadFromBytes(
        byte[] bytes,
        string? declaredMediaType = null,
        string? displayName = null,
        string source = "<bytes>",
        ImageAssetOptions? assetOptions = null)
    {
        ArgumentNullException.ThrowIfNull(bytes);
        var options = assetOptions ?? ImageAssetOptions.Default;
        ImageAssetOptions.Validate(options);
        return Build(bytes, declaredMediaType, displayName, new ImageSourceResolution
        {
            Kind = ImageSourceKind.DataUri,
            Source = source
        }, options);
    }

    private static (byte[] Bytes, string? DeclaredMediaType, string? FileName) ReadPayload(
        ImageSourceResolution resolution, ImageAssetOptions options)
    {
        if (resolution.Kind == ImageSourceKind.DataUri)
        {
            var payload = resolution.Payload ?? string.Empty;
            if (Encoding.UTF8.GetByteCount(payload) > options.MaxEncodedBytes)
            {
                throw new ImageSourceException(
                    ImageSourceErrorCode.EncodedBytesExceededLimit,
                    $"Image data URI payload is {Encoding.UTF8.GetByteCount(payload)} bytes, exceeding the {options.MaxEncodedBytes}-byte limit.");
            }

            byte[] bytes;
            try
            {
                bytes = resolution.IsBase64
                    ? Convert.FromBase64String(payload)
                    : Encoding.UTF8.GetBytes(Uri.UnescapeDataString(payload));
            }
            catch (FormatException ex)
            {
                throw new ImageSourceException(
                    ImageSourceErrorCode.InvalidDataUri,
                    "Image data URI payload is not valid base64.", ex);
            }

            if (bytes.Length > options.MaxEncodedBytes)
            {
                throw new ImageSourceException(
                    ImageSourceErrorCode.EncodedBytesExceededLimit,
                    $"Image data URI decodes to {bytes.Length} bytes, exceeding the {options.MaxEncodedBytes}-byte limit.");
            }
            return (bytes, resolution.DeclaredMediaType, null);
        }

        var filePath = resolution.FilePath!;
        long length;
        try
        {
            length = new FileInfo(filePath).Length;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or NotSupportedException)
        {
            throw new ImageSourceException(ImageSourceErrorCode.PathNotReadable, $"Image file could not be read: '{filePath}'.", ex);
        }
        if (length > options.MaxEncodedBytes)
        {
            throw new ImageSourceException(
                ImageSourceErrorCode.EncodedBytesExceededLimit,
                $"Image file '{filePath}' is {length} bytes, exceeding the {options.MaxEncodedBytes}-byte limit.");
        }

        byte[] fileBytes;
        try
        {
            fileBytes = File.ReadAllBytes(filePath);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or NotSupportedException)
        {
            throw new ImageSourceException(ImageSourceErrorCode.PathNotReadable, $"Image file could not be read: '{filePath}'.", ex);
        }
        if (fileBytes.Length > options.MaxEncodedBytes)
        {
            throw new ImageSourceException(
                ImageSourceErrorCode.EncodedBytesExceededLimit,
                $"Image file '{filePath}' grew to {fileBytes.Length} bytes while reading, exceeding the {options.MaxEncodedBytes}-byte limit.");
        }
        return (fileBytes, resolution.DeclaredMediaType, resolution.FileName);
    }

    private static ImageAsset Build(byte[] bytes, string? declaredMediaType, string? fileName, ImageSourceResolution resolution, ImageAssetOptions options)
    {
        if (bytes.Length == 0)
        {
            throw new ImageSourceException(ImageSourceErrorCode.InvalidImageData, "The image payload is empty.");
        }

        var metadata = ImageFormatSniffer.TryReadMetadata(bytes)
            ?? throw new ImageSourceException(
                ImageSourceErrorCode.UnsupportedMediaType,
                $"The image bytes do not match a supported format (png, jpeg, gif, bmp, tiff, svg). The media type is inferred from content, never from the file name.");

        var warnings = new List<ImageAssetWarning>(metadata.Warnings);
        if (declaredMediaType is { Length: > 0 } declared)
        {
            var declaredType = DocxImageMediaTypes.ParseContentType(declared);
            if (declaredType is null)
            {
                warnings.Add(new ImageAssetWarning(
                    ImageAssetWarningCode.DeclaredMediaTypeMismatch,
                    $"Declared media type '{declared}' is not a supported image type; the payload was sniffed as {metadata.MediaType}."));
            }
            else if (declaredType != metadata.MediaType)
            {
                warnings.Add(new ImageAssetWarning(
                    ImageAssetWarningCode.DeclaredMediaTypeMismatch,
                    $"Declared media type '{declared}' conflicts with the sniffed payload ({metadata.MediaType}); the content is authoritative."));
            }
        }

        if (metadata.MediaType == DocxImageMediaType.Svg)
        {
            warnings.Add(new ImageAssetWarning(
                ImageAssetWarningCode.SvgRasterFallbackMissing,
                "SVG has no raster fallback; hosts without SVG blip support will not render it."));
        }

        if (metadata.Width > options.MaxPixelDimension
            || metadata.Height > options.MaxPixelDimension
            || (long)metadata.Width * metadata.Height > options.MaxPixelArea)
        {
            throw new ImageSourceException(
                ImageSourceErrorCode.DimensionsExceedLimits,
                $"Image is {metadata.Width}×{metadata.Height}px, exceeding the configured pixel limits (max dimension {options.MaxPixelDimension}px, max area {options.MaxPixelArea}px).");
        }

        if (options.SupportedMediaTypes is { } supported && !supported.Contains(metadata.MediaType))
        {
            throw new ImageSourceException(
                ImageSourceErrorCode.UnsupportedMediaType,
                $"Media type {metadata.MediaType} is not enabled by the current asset options.");
        }

        var (naturalWidthPt, naturalHeightPt) = ComputeNaturalSize(metadata, warnings);
        var hash = Convert.ToHexStringLower(SHA256.HashData(bytes));
        var displayName = fileName ?? $"image_{hash[..12]}{DocxImageMediaTypes.GetExtension(metadata.MediaType)}";

        return new ImageAsset
        {
            Bytes = bytes,
            MediaType = metadata.MediaType,
            Width = metadata.Width,
            Height = metadata.Height,
            DpiX = metadata.DpiX,
            DpiY = metadata.DpiY,
            NaturalWidthPt = naturalWidthPt,
            NaturalHeightPt = naturalHeightPt,
            ContentHash = hash,
            DisplayName = displayName,
            Source = resolution.Source,
            Warnings = warnings
        };
    }

    private static (double WidthPt, double HeightPt) ComputeNaturalSize(ImageHeaderMetadata metadata, List<ImageAssetWarning> warnings)
    {
        const double defaultDpi = 96;
        const double pointsPerInch = 72;

        if (metadata.DpiX is { } dpiX && metadata.DpiY is { } dpiY)
        {
            var ratio = Math.Max(dpiX, dpiY) / Math.Min(dpiX, dpiY);
            if (ratio > 1.5)
            {
                warnings.Add(new ImageAssetWarning(
                    ImageAssetWarningCode.AspectDpiInconsistent,
                    $"X/Y DPI differ significantly ({dpiX:0.##}/{dpiY:0.##}); natural size is computed per-axis."));
            }
            return (metadata.Width / dpiX * pointsPerInch, metadata.Height / dpiY * pointsPerInch);
        }

        var dpi = metadata.DpiX ?? metadata.DpiY ?? defaultDpi;
        if (metadata.DpiX is null && metadata.DpiY is null)
        {
            warnings.Add(new ImageAssetWarning(
                ImageAssetWarningCode.IntrinsicDpiUnavailable,
                $"{metadata.MediaType} carries no intrinsic DPI; natural size assumes {defaultDpi:0} DPI."));
        }
        return (metadata.Width / dpi * pointsPerInch, metadata.Height / dpi * pointsPerInch);
    }
}
