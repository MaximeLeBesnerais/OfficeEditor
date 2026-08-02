using System.Buffers.Binary;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace DocxEditor.Core.Generation.Assets;

/// <summary>
/// Header-level image metadata: detected <see cref="DocxImageMediaType"/>, intrinsic pixel
/// dimensions, intrinsic DPI when the format carries it, and explicit warnings for
/// unsupported metadata (e.g. a TIFF whose resolution unit is unspecified, or a BMP core
/// header without resolution). Dimensions and DPI are read by sniffing a bounded prefix —
/// the full payload is never decoded.
/// </summary>
public sealed record ImageHeaderMetadata
{
    /// <summary>Content-detected media type.</summary>
    public required DocxImageMediaType MediaType { get; init; }

    /// <summary>Intrinsic width in pixels.</summary>
    public required int Width { get; init; }

    /// <summary>Intrinsic height in pixels.</summary>
    public required int Height { get; init; }

    /// <summary>Horizontal DPI when the format declares it (PNG pHYs, JPEG JFIF, BMP PPM, TIFF resolution).</summary>
    public double? DpiX { get; init; }

    /// <summary>Vertical DPI when the format declares it.</summary>
    public double? DpiY { get; init; }

    /// <summary>Explicit warnings for metadata the format carries but cannot be interpreted.</summary>
    public IReadOnlyList<ImageAssetWarning> Warnings { get; init; } = [];
}

/// <summary>
/// Dependency-free, content-based format detection and header sniffing for the supported
/// image media types (PNG, JPEG, GIF, BMP, TIFF, SVG). Reads only the bytes needed to find
/// media type, dimensions and DPI; returns null for unrecognized or truncated payloads so
/// callers can fail loudly with a typed error. Detection never trusts the filename.
/// </summary>
public static class ImageFormatSniffer
{
    private const int SvgScanPrefix = 64 * 1024;
    private const int PngChunkScanLimit = 1024 * 1024;

    private static readonly Regex SvgLengthPattern = new(
        @"([a-z][a-z0-9-]*)\s*=\s*""([^""]*)""",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

    /// <summary>Sniffs the media type only.</summary>
    public static DocxImageMediaType? TryDetectFormat(ReadOnlySpan<byte> bytes)
    {
        if (TryReadMetadata(bytes) is { } metadata)
        {
            return metadata.MediaType;
        }
        return null;
    }

    /// <summary>Sniffs media type, dimensions and DPI, or returns null for unrecognized payloads.</summary>
    public static ImageHeaderMetadata? TryReadMetadata(ReadOnlySpan<byte> bytes)
    {
        var warnings = new List<ImageAssetWarning>();
        var metadata = TryRead(bytes, warnings);
        if (metadata is null)
        {
            return null;
        }
        return metadata with { Warnings = warnings };
    }

    private static ImageHeaderMetadata? TryRead(ReadOnlySpan<byte> bytes, List<ImageAssetWarning> warnings)
    {
        if (IsPng(bytes))
        {
            return TryReadPng(bytes, warnings);
        }
        if (IsJpeg(bytes))
        {
            return TryReadJpeg(bytes, warnings);
        }
        if (IsGif(bytes))
        {
            return TryReadGif(bytes, warnings);
        }
        if (IsBmp(bytes))
        {
            return TryReadBmp(bytes, warnings);
        }
        if (IsTiff(bytes))
        {
            return TryReadTiff(bytes, warnings);
        }
        if (TryReadSvg(bytes, warnings) is { } svg)
        {
            return svg;
        }
        return null;
    }

    #region PNG

    private static bool IsPng(ReadOnlySpan<byte> bytes) =>
        bytes.Length >= 8
        && bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47
        && bytes[4] == 0x0D && bytes[5] == 0x0A && bytes[6] == 0x1A && bytes[7] == 0x0A;

    private static ImageHeaderMetadata? TryReadPng(ReadOnlySpan<byte> bytes, List<ImageAssetWarning> warnings)
    {
        if (bytes.Length < 24)
        {
            return null;
        }
        var width = BinaryPrimitives.ReadInt32BigEndian(bytes.Slice(16, 4));
        var height = BinaryPrimitives.ReadInt32BigEndian(bytes.Slice(20, 4));
        if (width <= 0 || height <= 0)
        {
            return null;
        }
        var dpi = ReadPngDpi(bytes);
        return new ImageHeaderMetadata
        {
            MediaType = DocxImageMediaType.Png,
            Width = width,
            Height = height,
            DpiX = dpi.X,
            DpiY = dpi.Y
        };
    }

    /// <summary>
    /// Scans PNG chunks for a pHYs chunk (pixels per meter) and converts to DPI. A unit of 0
    /// (unknown) yields no DPI. Scanning is bounded to a prefix so hostile chunk tables
    /// cannot drive unbounded iteration.
    /// </summary>
    private static (double? X, double? Y) ReadPngDpi(ReadOnlySpan<byte> bytes)
    {
        var offset = 8;
        var limit = Math.Min(bytes.Length, PngChunkScanLimit);
        while (offset + 8 <= limit)
        {
            var length = BinaryPrimitives.ReadInt32BigEndian(bytes.Slice(offset, 4));
            var type = BinaryPrimitives.ReadUInt32BigEndian(bytes.Slice(offset + 4, 4));
            if (length < 0 || offset + 12 + length > limit)
            {
                return (null, null);
            }
            if (type == 0x70485973u && length >= 9) // "pHYs"
            {
                var xPpm = BinaryPrimitives.ReadUInt32BigEndian(bytes.Slice(offset + 8, 4));
                var yPpm = BinaryPrimitives.ReadUInt32BigEndian(bytes.Slice(offset + 12, 4));
                var unit = bytes[offset + 16];
                return unit == 1 && xPpm > 0 && yPpm > 0
                    ? (xPpm * 0.0254, yPpm * 0.0254)
                    : (null, null);
            }
            if (length == 0)
            {
                return (null, null);
            }
            offset += 12 + length;
        }
        return (null, null);
    }

    #endregion

    #region JPEG

    private static bool IsJpeg(ReadOnlySpan<byte> bytes) =>
        bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xD8;

    private static ImageHeaderMetadata? TryReadJpeg(ReadOnlySpan<byte> bytes, List<ImageAssetWarning> warnings)
    {
        var dimensions = ReadJpegDimensions(bytes);
        if (dimensions is not { } dims)
        {
            return null;
        }
        var dpi = ReadJpegDpi(bytes, warnings);
        return new ImageHeaderMetadata
        {
            MediaType = DocxImageMediaType.Jpeg,
            Width = dims.Width,
            Height = dims.Height,
            DpiX = dpi?.X,
            DpiY = dpi?.Y
        };
    }

    private static (int Width, int Height)? ReadJpegDimensions(ReadOnlySpan<byte> bytes)
    {
        var offset = 2;
        while (offset + 3 < bytes.Length)
        {
            if (bytes[offset] != 0xFF)
            {
                return null;
            }
            var marker = bytes[offset + 1];
            if (marker == 0xFF)
            {
                offset++;
                continue;
            }
            if (marker is 0x01 or >= 0xD0 and <= 0xD9)
            {
                offset += 2;
                continue;
            }
            var length = BinaryPrimitives.ReadUInt16BigEndian(bytes.Slice(offset + 2, 2));
            if (length < 2)
            {
                return null;
            }
            if (marker is >= 0xC0 and <= 0xCF and not 0xC4 and not 0xC8 and not 0xCC)
            {
                if (offset + 9 >= bytes.Length)
                {
                    return null;
                }
                var height = BinaryPrimitives.ReadUInt16BigEndian(bytes.Slice(offset + 5, 2));
                var width = BinaryPrimitives.ReadUInt16BigEndian(bytes.Slice(offset + 7, 2));
                return width > 0 && height > 0 ? (width, height) : null;
            }
            offset += 2 + length;
        }
        return null;
    }

    /// <summary>
    /// Reads the JFIF APP0 density (units 1 = dots/inch, 2 = dots/cm). Units 0 (aspect ratio)
    /// yields no DPI with an explicit warning.
    /// </summary>
    private static (double X, double Y)? ReadJpegDpi(ReadOnlySpan<byte> bytes, List<ImageAssetWarning> warnings)
    {
        var offset = 2;
        while (offset + 3 < bytes.Length)
        {
            if (bytes[offset] != 0xFF)
            {
                return null;
            }
            var marker = bytes[offset + 1];
            if (marker == 0xFF)
            {
                offset++;
                continue;
            }
            if (marker is 0x01 or >= 0xD0 and <= 0xD9)
            {
                offset += 2;
                continue;
            }
            var length = BinaryPrimitives.ReadUInt16BigEndian(bytes.Slice(offset + 2, 2));
            if (length < 2)
            {
                return null;
            }
            if (marker == 0xE0 && length >= 14) // APP0
            {
                var data = offset + 4;
                if (data + 12 > bytes.Length)
                {
                    return null;
                }
                if (bytes.Slice(data, 5).SequenceEqual("JFIF\0"u8))
                {
                    var units = bytes[data + 7];
                    var x = BinaryPrimitives.ReadUInt16BigEndian(bytes.Slice(data + 8, 2));
                    var y = BinaryPrimitives.ReadUInt16BigEndian(bytes.Slice(data + 10, 2));
                    if (units == 1)
                    {
                        return (x, y);
                    }
                    if (units == 2)
                    {
                        return (x * 2.54, y * 2.54);
                    }
                    warnings.Add(new ImageAssetWarning(
                        ImageAssetWarningCode.DensityUnitsUnspecified,
                        "JPEG JFIF density units are unspecified; natural size assumes 96 DPI."));
                    return null;
                }
            }
            offset += 2 + length;
        }
        return null;
    }

    #endregion

    #region GIF

    private static bool IsGif(ReadOnlySpan<byte> bytes) =>
        bytes.Length >= 6
        && bytes[0] == 0x47 && bytes[1] == 0x49 && bytes[2] == 0x46 && bytes[3] == 0x38
        && (bytes[4] == 0x37 || bytes[4] == 0x39) && bytes[5] == 0x61;

    private static ImageHeaderMetadata? TryReadGif(ReadOnlySpan<byte> bytes, List<ImageAssetWarning> warnings)
    {
        if (bytes.Length < 10)
        {
            return null;
        }
        var width = BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(6, 2));
        var height = BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(8, 2));
        if (width == 0 || height == 0)
        {
            return null;
        }
        // GIF carries no DPI (its pixel aspect ratio is non-standard); natural size will
        // assume the 96 DPI default and the loader warns once.
        return new ImageHeaderMetadata
        {
            MediaType = DocxImageMediaType.Gif,
            Width = width,
            Height = height
        };
    }

    #endregion

    #region BMP

    private static bool IsBmp(ReadOnlySpan<byte> bytes) =>
        bytes.Length >= 2 && bytes[0] == (byte)'B' && bytes[1] == (byte)'M';

    private static ImageHeaderMetadata? TryReadBmp(ReadOnlySpan<byte> bytes, List<ImageAssetWarning> warnings)
    {
        if (bytes.Length < 26)
        {
            return null;
        }
        var headerSize = BinaryPrimitives.ReadInt32LittleEndian(bytes.Slice(14, 4));
        int width, height;
        double? dpiX = null, dpiY = null;
        if (headerSize >= 40)
        {
            if (bytes.Length < 46)
            {
                return null;
            }
            width = BinaryPrimitives.ReadInt32LittleEndian(bytes.Slice(18, 4));
            height = BinaryPrimitives.ReadInt32LittleEndian(bytes.Slice(22, 4));
            var ppmX = BinaryPrimitives.ReadInt32LittleEndian(bytes.Slice(38, 4));
            var ppmY = BinaryPrimitives.ReadInt32LittleEndian(bytes.Slice(42, 4));
            if (ppmX > 0 && ppmY > 0)
            {
                dpiX = ppmX * 0.0254;
                dpiY = ppmY * 0.0254;
            }
        }
        else if (headerSize == 12) // BITMAPCOREHEADER
        {
            if (bytes.Length < 22)
            {
                return null;
            }
            width = BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(18, 2));
            height = BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(20, 2));
            warnings.Add(new ImageAssetWarning(
                ImageAssetWarningCode.IntrinsicDpiUnavailable,
                "BMP core header carries no resolution metadata; natural size assumes 96 DPI."));
        }
        else
        {
            return null;
        }

        if (width <= 0 || height == 0)
        {
            return null;
        }
        height = Math.Abs(height); // negative = top-down bitmap
        return new ImageHeaderMetadata
        {
            MediaType = DocxImageMediaType.Bmp,
            Width = width,
            Height = height,
            DpiX = dpiX,
            DpiY = dpiY
        };
    }

    #endregion

    #region TIFF

    private static bool IsTiff(ReadOnlySpan<byte> bytes) =>
        bytes.Length >= 8
        && ((bytes[0] == 0x49 && bytes[1] == 0x49) || (bytes[0] == 0x4D && bytes[1] == 0x4D))
        && ReadU16(bytes, 2, littleEndian: bytes[0] == 0x49) == 42;

    private static ImageHeaderMetadata? TryReadTiff(ReadOnlySpan<byte> bytes, List<ImageAssetWarning> warnings)
    {
        var littleEndian = bytes[0] == 0x49;
        var ifdOffset = (int)ReadU32(bytes, 4, littleEndian);
        if (ifdOffset + 2 > bytes.Length)
        {
            return null;
        }
        var entryCount = ReadU16(bytes, ifdOffset, littleEndian);
        if (ifdOffset + 2 + entryCount * 12 > bytes.Length)
        {
            return null;
        }

        int? width = null, height = null;
        double? xResolution = null, yResolution = null;
        int? resolutionUnit = null;

        for (var i = 0; i < entryCount; i++)
        {
            var entry = ifdOffset + 2 + i * 12;
            var tag = ReadU16(bytes, entry, littleEndian);
            var type = ReadU16(bytes, entry + 2, littleEndian);
            var count = ReadU32(bytes, entry + 4, littleEndian);
            switch (tag)
            {
                case 256 when type is 3 or 4 && count == 1: // ImageWidth
                    width = type == 3 ? ReadU16(bytes, entry + 8, littleEndian) : (int)ReadU32(bytes, entry + 8, littleEndian);
                    break;
                case 257 when type is 3 or 4 && count == 1: // ImageLength
                    height = type == 3 ? ReadU16(bytes, entry + 8, littleEndian) : (int)ReadU32(bytes, entry + 8, littleEndian);
                    break;
                case 282 when type == 5: // XResolution (RATIONAL)
                    xResolution = ReadRational(bytes, (int)ReadU32(bytes, entry + 8, littleEndian), littleEndian);
                    break;
                case 283 when type == 5: // YResolution (RATIONAL)
                    yResolution = ReadRational(bytes, (int)ReadU32(bytes, entry + 8, littleEndian), littleEndian);
                    break;
                case 296 when type == 3 && count == 1: // ResolutionUnit
                    resolutionUnit = ReadU16(bytes, entry + 8, littleEndian);
                    break;
            }
        }

        if (width is not > 0 || height is not > 0)
        {
            return null;
        }

        double? dpiX = null, dpiY = null;
        if (xResolution is { } xr && yResolution is { } yr)
        {
            if (resolutionUnit == 2) // inch
            {
                dpiX = xr;
                dpiY = yr;
            }
            else if (resolutionUnit == 3) // cm
            {
                dpiX = xr * 2.54;
                dpiY = yr * 2.54;
            }
            else
            {
                warnings.Add(new ImageAssetWarning(
                    ImageAssetWarningCode.ResolutionUnitUnspecified,
                    "TIFF resolution unit is unspecified; natural size assumes 96 DPI."));
            }
        }

        return new ImageHeaderMetadata
        {
            MediaType = DocxImageMediaType.Tiff,
            Width = width.Value,
            Height = height.Value,
            DpiX = dpiX,
            DpiY = dpiY
        };
    }

    private static double? ReadRational(ReadOnlySpan<byte> bytes, int offset, bool littleEndian)
    {
        if (offset < 0 || offset + 8 > bytes.Length)
        {
            return null;
        }
        var numerator = ReadU32(bytes, offset, littleEndian);
        var denominator = ReadU32(bytes, offset + 4, littleEndian);
        if (denominator == 0)
        {
            return null;
        }
        return numerator / (double)denominator;
    }

    private static ushort ReadU16(ReadOnlySpan<byte> bytes, int offset, bool littleEndian) =>
        littleEndian
            ? BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(offset, 2))
            : BinaryPrimitives.ReadUInt16BigEndian(bytes.Slice(offset, 2));

    private static uint ReadU32(ReadOnlySpan<byte> bytes, int offset, bool littleEndian) =>
        littleEndian
            ? BinaryPrimitives.ReadUInt32LittleEndian(bytes.Slice(offset, 4))
            : BinaryPrimitives.ReadUInt32BigEndian(bytes.Slice(offset, 4));

    #endregion

    #region SVG

    private static ImageHeaderMetadata? TryReadSvg(ReadOnlySpan<byte> bytes, List<ImageAssetWarning> warnings)
    {
        if (bytes.Length == 0)
        {
            return null;
        }
        var prefixLength = Math.Min(bytes.Length, SvgScanPrefix);
        var text = Encoding.UTF8.GetString(bytes[..prefixLength]);

        var tagStart = text.IndexOf("<svg", StringComparison.OrdinalIgnoreCase);
        if (tagStart < 0)
        {
            return null;
        }
        var tagEnd = text.IndexOf('>', tagStart);
        if (tagEnd < 0)
        {
            return null;
        }
        var tag = text[tagStart..tagEnd];
        if (!tag.Contains("http://www.w3.org/2000/svg", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        var widthPx = ReadSvgLength(tag, "width");
        var heightPx = ReadSvgLength(tag, "height");
        if (widthPx is null || heightPx is null)
        {
            var viewBox = ReadViewBox(tag);
            if (viewBox is null)
            {
                return null;
            }
            widthPx = viewBox.Value.Width;
            heightPx = viewBox.Value.Height;
        }

        // SVG length units are CSS pixels: 1in = 96px, so the intrinsic density is 96 DPI.
        // Absolute units (cm/mm/in/pt/pc) are converted to CSS px, keeping physical size exact.
        return new ImageHeaderMetadata
        {
            MediaType = DocxImageMediaType.Svg,
            Width = widthPx.Value,
            Height = heightPx.Value,
            DpiX = 96,
            DpiY = 96
        };
    }

    private static int? ReadSvgLength(string tag, string attribute)
    {
        var match = SvgLengthPattern.Match(tag);
        while (match.Success)
        {
            if (string.Equals(match.Groups[1].Value, attribute, StringComparison.OrdinalIgnoreCase))
            {
                var value = match.Groups[2].Value.Trim();
                if (value.Length == 0)
                {
                    return null;
                }
                var numberText = new string(value.TakeWhile(c => char.IsAsciiDigit(c) || c == '.').ToArray());
                if (numberText.Length == 0
                    || !double.TryParse(numberText, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)
                    || number <= 0)
                {
                    return null;
                }
                var unit = value[numberText.Length..].Trim().ToLowerInvariant();
                double px = unit switch
                {
                    "" or "px" => number,
                    "pt" => number * 96.0 / 72.0,
                    "pc" => number * 16.0,
                    "in" => number * 96.0,
                    "cm" => number * 96.0 / 2.54,
                    "mm" => number * 96.0 / 25.4,
                    _ => -1 // %, em, ex cannot be resolved without a parent viewport
                };
                if (px <= 0)
                {
                    return null;
                }
                return (int)Math.Round(px, MidpointRounding.AwayFromZero);
            }
            match = match.NextMatch();
        }
        return null;
    }

    private static (int Width, int Height)? ReadViewBox(string tag)
    {
        var match = Regex.Match(tag, @"viewbox\s*=\s*""([^""]*)""", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        if (!match.Success)
        {
            return null;
        }
        var parts = match.Groups[1].Value.Split([',', ' ', '\t', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length < 4)
        {
            return null;
        }
        if (!double.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out var width)
            || !double.TryParse(parts[3], NumberStyles.Float, CultureInfo.InvariantCulture, out var height)
            || width <= 0 || height <= 0)
        {
            return null;
        }
        return ((int)Math.Round(width, MidpointRounding.AwayFromZero), (int)Math.Round(height, MidpointRounding.AwayFromZero));
    }

    #endregion
}
