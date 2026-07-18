using System.Buffers.Binary;

namespace PptxEditor.Core.Services;

/// <summary>
/// Minimal, dependency-free pixel-dimension sniffing for PNG, JPEG, GIF and BMP
/// headers. Reads only the bytes needed to find width/height; returns null for
/// SVG/EMF/unknown/truncated payloads so callers can fall back gracefully.
/// </summary>
public static class ImageHeaderSniffer
{
    /// <summary>Sniffs (width, height) in pixels, or null when the format is unrecognized.</summary>
    public static (int Width, int Height)? TryGetPixelDimensions(ReadOnlySpan<byte> bytes)
    {
        if (bytes.Length >= 24
            && bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47
            && bytes[4] == 0x0D && bytes[5] == 0x0A && bytes[6] == 0x1A && bytes[7] == 0x0A)
        {
            // PNG: 8-byte signature, then IHDR length(4)+type(4); width/height are
            // big-endian int32 at offsets 16/20.
            var width = BinaryPrimitives.ReadInt32BigEndian(bytes.Slice(16, 4));
            var height = BinaryPrimitives.ReadInt32BigEndian(bytes.Slice(20, 4));
            return width > 0 && height > 0 ? (width, height) : null;
        }

        if (bytes.Length >= 10
            && bytes[0] == 0x47 && bytes[1] == 0x49 && bytes[2] == 0x46 && bytes[3] == 0x38
            && (bytes[4] == 0x37 || bytes[4] == 0x39) && bytes[5] == 0x61)
        {
            // GIF87a/89a: logical screen width/height, little-endian ushort at 6/8.
            var width = BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(6, 2));
            var height = BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(8, 2));
            return width > 0 && height > 0 ? (width, height) : null;
        }

        if (bytes.Length >= 26 && bytes[0] == 0x42 && bytes[1] == 0x4D)
        {
            // BMP: "BM", BITMAPINFOHEADER width/height are little-endian int32 at
            // 18/22. Height may be negative for top-down bitmaps.
            var width = BinaryPrimitives.ReadInt32LittleEndian(bytes.Slice(18, 4));
            var height = BinaryPrimitives.ReadInt32LittleEndian(bytes.Slice(22, 4));
            height = Math.Abs(height);
            return width > 0 && height > 0 ? (width, height) : null;
        }

        if (bytes.Length >= 4 && bytes[0] == 0xFF && bytes[1] == 0xD8)
        {
            return ReadJpegDimensions(bytes);
        }

        return null;
    }

    private static (int Width, int Height)? ReadJpegDimensions(ReadOnlySpan<byte> bytes)
    {
        // JPEG: FFD8, then a sequence of segments (0xFF marker, big-endian length,
        // payload). Start-of-frame markers (C0–CF except DHT C4, JPG C8, DAC CC)
        // carry height/width as big-endian ushorts right after the precision byte.
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
                offset++; // padding byte before the real marker
                continue;
            }
            if (marker is 0x01 or >= 0xD0 and <= 0xD9)
            {
                offset += 2; // standalone marker without a length (RSTn, SOI, EOI, TEM)
                continue;
            }

            var length = BinaryPrimitives.ReadUInt16BigEndian(bytes.Slice(offset + 2, 2));
            if (length < 2)
            {
                return null;
            }
            if (marker is >= 0xC0 and <= 0xCF and not 0xC4 and not 0xC8 and not 0xCC)
            {
                if (offset + 8 >= bytes.Length)
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
}
