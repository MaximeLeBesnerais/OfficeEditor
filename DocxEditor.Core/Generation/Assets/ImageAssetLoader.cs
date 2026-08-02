using System.Buffers;
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
        return Build((byte[])bytes.Clone(), declaredMediaType, displayName, new ImageSourceResolution
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
            if (resolution.IsBase64)
            {
                try
                {
                    bytes = Convert.FromBase64String(payload);
                }
                catch (FormatException ex)
                {
                    throw new ImageSourceException(
                        ImageSourceErrorCode.InvalidDataUri,
                        "Image data URI payload is not valid base64.", ex);
                }
            }
            else
            {
                bytes = DecodePercentEncodedPayload(payload);
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
        byte[] fileBytes;
        try
        {
            using var stream = new FileStream(
                filePath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.Read,
                bufferSize: 81920,
                FileOptions.SequentialScan);
            var length = stream.Length;
            if (length > options.MaxEncodedBytes || length > int.MaxValue)
            {
                throw new ImageSourceException(
                    ImageSourceErrorCode.EncodedBytesExceededLimit,
                    $"Image file '{filePath}' is {length} bytes, exceeding the {options.MaxEncodedBytes}-byte limit.");
            }

            fileBytes = GC.AllocateUninitializedArray<byte>((int)length);
            stream.ReadExactly(fileBytes);
            if (stream.ReadByte() != -1)
            {
                throw new ImageSourceException(
                    ImageSourceErrorCode.PathNotReadable,
                    $"Image file changed while it was being read: '{filePath}'.");
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or NotSupportedException)
        {
            throw new ImageSourceException(ImageSourceErrorCode.PathNotReadable, $"Image file could not be read: '{filePath}'.", ex);
        }
        return (fileBytes, resolution.DeclaredMediaType, resolution.FileName);
    }

    private static byte[] DecodePercentEncodedPayload(string payload)
    {
        var writer = new ArrayBufferWriter<byte>(Math.Max(1, payload.Length));
        var segmentStart = 0;

        for (var i = 0; i < payload.Length; i++)
        {
            if (payload[i] != '%')
            {
                continue;
            }

            WriteUtf8(payload.AsSpan(segmentStart, i - segmentStart), writer);
            if (i + 2 >= payload.Length
                || !TryReadHexNibble(payload[i + 1], out var high)
                || !TryReadHexNibble(payload[i + 2], out var low))
            {
                throw new ImageSourceException(
                    ImageSourceErrorCode.InvalidDataUri,
                    $"Image data URI payload contains a malformed percent escape at character {i}.");
            }

            writer.GetSpan(1)[0] = (byte)((high << 4) | low);
            writer.Advance(1);
            i += 2;
            segmentStart = i + 1;
        }

        WriteUtf8(payload.AsSpan(segmentStart), writer);
        return writer.WrittenSpan.ToArray();
    }

    private static void WriteUtf8(ReadOnlySpan<char> value, ArrayBufferWriter<byte> writer)
    {
        if (value.IsEmpty)
        {
            return;
        }

        var byteCount = Encoding.UTF8.GetByteCount(value);
        var written = Encoding.UTF8.GetBytes(value, writer.GetSpan(byteCount));
        writer.Advance(written);
    }

    private static bool TryReadHexNibble(char value, out int nibble)
    {
        if (value is >= '0' and <= '9')
        {
            nibble = value - '0';
            return true;
        }
        if (value is >= 'A' and <= 'F')
        {
            nibble = value - 'A' + 10;
            return true;
        }
        if (value is >= 'a' and <= 'f')
        {
            nibble = value - 'a' + 10;
            return true;
        }

        nibble = 0;
        return false;
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

        var dpiX = UsableDpiOrNull(metadata.DpiX);
        var dpiY = UsableDpiOrNull(metadata.DpiY);
        var (naturalWidthPt, naturalHeightPt) = ComputeNaturalSize(metadata, dpiX, dpiY, warnings);
        var hash = Convert.ToHexStringLower(SHA256.HashData(bytes));
        var displayName = fileName ?? $"image_{hash[..12]}{DocxImageMediaTypes.GetExtension(metadata.MediaType)}";

        return new ImageAsset
        {
            Bytes = bytes,
            MediaType = metadata.MediaType,
            Width = metadata.Width,
            Height = metadata.Height,
            DpiX = dpiX,
            DpiY = dpiY,
            NaturalWidthPt = naturalWidthPt,
            NaturalHeightPt = naturalHeightPt,
            ContentHash = hash,
            DisplayName = displayName,
            Source = resolution.Source,
            Warnings = warnings
        };
    }

    private static (double WidthPt, double HeightPt) ComputeNaturalSize(
        ImageHeaderMetadata metadata,
        double? dpiX,
        double? dpiY,
        List<ImageAssetWarning> warnings)
    {
        const double defaultDpi = 96;
        const double pointsPerInch = 72;

        if (dpiX is { } horizontalDpi && dpiY is { } verticalDpi)
        {
            var ratio = Math.Max(horizontalDpi, verticalDpi) / Math.Min(horizontalDpi, verticalDpi);
            if (ratio > 1.5)
            {
                warnings.Add(new ImageAssetWarning(
                    ImageAssetWarningCode.AspectDpiInconsistent,
                    $"X/Y DPI differ significantly ({horizontalDpi:0.##}/{verticalDpi:0.##}); natural size is computed per-axis."));
            }
            var widthPt = metadata.Width / horizontalDpi * pointsPerInch;
            var heightPt = metadata.Height / verticalDpi * pointsPerInch;
            if (double.IsFinite(widthPt) && widthPt > 0 && double.IsFinite(heightPt) && heightPt > 0)
            {
                return (widthPt, heightPt);
            }

            dpiX = null;
            dpiY = null;
        }

        var dpi = dpiX ?? dpiY ?? defaultDpi;
        if (dpiX is null && dpiY is null)
        {
            warnings.Add(new ImageAssetWarning(
                ImageAssetWarningCode.IntrinsicDpiUnavailable,
                $"{metadata.MediaType} carries no intrinsic DPI; natural size assumes {defaultDpi:0} DPI."));
        }
        var naturalWidthPt = metadata.Width / dpi * pointsPerInch;
        var naturalHeightPt = metadata.Height / dpi * pointsPerInch;
        if (double.IsFinite(naturalWidthPt) && naturalWidthPt > 0
            && double.IsFinite(naturalHeightPt) && naturalHeightPt > 0)
        {
            return (naturalWidthPt, naturalHeightPt);
        }

        warnings.Add(new ImageAssetWarning(
            ImageAssetWarningCode.IntrinsicDpiUnavailable,
            $"{metadata.MediaType} intrinsic DPI does not produce finite dimensions; natural size assumes {defaultDpi:0} DPI."));
        return (
            metadata.Width / defaultDpi * pointsPerInch,
            metadata.Height / defaultDpi * pointsPerInch);
    }

    private static double? UsableDpiOrNull(double? dpi) =>
        dpi is { } value && double.IsFinite(value) && value > 0 ? value : null;
}
