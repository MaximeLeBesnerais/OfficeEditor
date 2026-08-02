using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;
using DocxEditor.Core.Generation.Assets;

namespace DocxEditor.Tests.Generation.Docx;

/// <summary>
/// Dependency-free image fixture factory for the DOCX generation test suite. Each builder
/// produces a byte payload whose header the production <see cref="ImageFormatSniffer"/>
/// recognizes (media type, dimensions, optional DPI). Payloads are built programmatically so
/// the tests never depend on checked-in binary assets.
/// </summary>
public static class TestImages
{
    /// <summary>Builds a valid RGBA PNG of the given dimensions, optionally with a pHYs (pixels-per-meter) chunk.</summary>
    public static byte[] Png(int width, int height, int? dpiX = null, int? dpiY = null)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(width);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(height);

        var signature = new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
        var stream = new MemoryStream();
        stream.Write(signature);

        var header = new byte[13];
        BinaryPrimitives.WriteInt32BigEndian(header.AsSpan(0, 4), width);
        BinaryPrimitives.WriteInt32BigEndian(header.AsSpan(4, 4), height);
        header[8] = 8;   // bit depth
        header[9] = 6;   // color type RGBA
        header[10] = 0;  // compression
        header[11] = 0;  // filter
        header[12] = 0;  // interlace
        WriteChunk(stream, "IHDR", header);

        if (dpiX is not null || dpiY is not null)
        {
            var horizontalDpi = dpiX ?? dpiY ?? 96;
            var verticalDpi = dpiY ?? dpiX ?? 96;
            var phy = new byte[9];
            BinaryPrimitives.WriteUInt32BigEndian(phy.AsSpan(0, 4), DpiToPixelsPerMeter(horizontalDpi));
            BinaryPrimitives.WriteUInt32BigEndian(phy.AsSpan(4, 4), DpiToPixelsPerMeter(verticalDpi));
            phy[8] = 1; // unit: meters
            WriteChunk(stream, "pHYs", phy);
        }

        var scanlines = new byte[height * (1 + width * 4)];
        for (var y = 0; y < height; y++)
        {
            var row = y * (1 + width * 4);
            scanlines[row] = 0; // filter: none
            for (var x = 0; x < width; x++)
            {
                var offset = row + 1 + x * 4;
                scanlines[offset] = 0x2E;
                scanlines[offset + 1] = 0x74;
                scanlines[offset + 2] = 0xB6;
                scanlines[offset + 3] = 0xFF;
            }
        }

        using (var compressed = new MemoryStream())
        using (var zlib = new ZLibStream(compressed, CompressionLevel.SmallestSize, leaveOpen: true))
        {
            zlib.Write(scanlines);
            zlib.Flush();
            WriteChunk(stream, "IDAT", compressed.ToArray());
        }

        WriteChunk(stream, "IEND", Array.Empty<byte>());
        return stream.ToArray();
    }

    /// <summary>
    /// Builds a minimal JFIF JPEG whose header carries <paramref name="density"/>
    /// in <paramref name="units"/> (1 = dots/inch, 2 = dots/cm, 0 = aspect only). The payload
    /// is header-only (no image data); the sniffer never decodes scan data.
    /// </summary>
    public static byte[] Jpeg(int width, int height, int density = 72, int units = 1)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(width);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(height);

        var stream = new MemoryStream();
        stream.WriteByte(0xFF);
        stream.WriteByte(0xD8); // SOI

        // APP0 JFIF: marker, length, "JFIF\0", version 1.1, units, x/y density, thumbnail 0x0.
        stream.WriteByte(0xFF);
        stream.WriteByte(0xE0);
        WriteU16Be(stream, 16);
        stream.Write(Encoding.ASCII.GetBytes("JFIF\0"));
        stream.WriteByte(0x01);
        stream.WriteByte(0x01);
        stream.WriteByte((byte)units);
        WriteU16Be(stream, density);
        WriteU16Be(stream, density);
        stream.WriteByte(0x00);
        stream.WriteByte(0x00);

        // SOF0: marker, length (8 + 3 × components), precision, height, width, components.
        stream.WriteByte(0xFF);
        stream.WriteByte(0xC0);
        WriteU16Be(stream, 11);
        stream.WriteByte(8);
        WriteU16Be(stream, height);
        WriteU16Be(stream, width);
        stream.WriteByte(1);
        stream.WriteByte(0x01);
        stream.WriteByte(0x11);
        stream.WriteByte(0x00);

        stream.WriteByte(0xFF);
        stream.WriteByte(0xD9); // EOI
        return stream.ToArray();
    }

    /// <summary>Builds a GIF89a header with the given logical screen dimensions.</summary>
    public static byte[] Gif(int width, int height)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(width);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(height);

        var bytes = new byte[13];
        Encoding.ASCII.GetBytes("GIF89a").CopyTo(bytes, 0);
        BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(6, 2), (ushort)width);
        BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(8, 2), (ushort)height);
        bytes[10] = 0x00; // no global color table
        bytes[11] = 0x00;
        bytes[12] = 0x3B; // trailer
        return bytes;
    }

    /// <summary>Builds a BMP V3 (BITMAPINFOHEADER) payload; negative <paramref name="height"/> = top-down.</summary>
    public static byte[] Bmp(int width, int height, int pixelsPerMeterX = 0, int pixelsPerMeterY = 0)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(width);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(Math.Abs(height));

        var bytes = new byte[54];
        bytes[0] = (byte)'B';
        bytes[1] = (byte)'M';
        BinaryPrimitives.WriteInt32LittleEndian(bytes.AsSpan(2, 4), bytes.Length);
        BinaryPrimitives.WriteInt32LittleEndian(bytes.AsSpan(10, 4), 54);
        BinaryPrimitives.WriteInt32LittleEndian(bytes.AsSpan(14, 4), 40);  // header size
        BinaryPrimitives.WriteInt32LittleEndian(bytes.AsSpan(18, 4), width);
        BinaryPrimitives.WriteInt32LittleEndian(bytes.AsSpan(22, 4), height);
        BinaryPrimitives.WriteInt16LittleEndian(bytes.AsSpan(26, 2), 1);   // planes
        BinaryPrimitives.WriteInt16LittleEndian(bytes.AsSpan(28, 2), 24);  // bpp
        BinaryPrimitives.WriteInt32LittleEndian(bytes.AsSpan(34, 4), width * Math.Abs(height) * 3);
        BinaryPrimitives.WriteInt32LittleEndian(bytes.AsSpan(38, 4), pixelsPerMeterX);
        BinaryPrimitives.WriteInt32LittleEndian(bytes.AsSpan(42, 4), pixelsPerMeterY);
        return bytes;
    }

    /// <summary>Builds a BMP core-header (BITMAPCOREHEADER) payload, which carries no resolution.</summary>
    public static byte[] BmpCore(int width, int height)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(width);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(height);

        var bytes = new byte[26];
        bytes[0] = (byte)'B';
        bytes[1] = (byte)'M';
        BinaryPrimitives.WriteInt32LittleEndian(bytes.AsSpan(10, 4), 26);
        BinaryPrimitives.WriteInt32LittleEndian(bytes.AsSpan(14, 4), 12);  // core header size
        BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(18, 2), (ushort)width);
        BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(20, 2), (ushort)height);
        BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(22, 2), 1);
        BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(24, 2), 24);
        return bytes;
    }

    /// <summary>
    /// Builds a TIFF payload. Supports both byte orders; optional resolution entries
    /// (RATIONAL 282/283) and resolution unit (SHORT 296). <paramref name="resolutionUnit"/>
    /// uses TIFF values: 1 = none, 2 = inch, 3 = cm.
    /// </summary>
    public static byte[] Tiff(
        int width,
        int height,
        bool littleEndian = true,
        double? dpiX = null,
        double? dpiY = null,
        int? resolutionUnit = null)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(width);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(height);

        var entries = new List<(int Tag, int Type, uint Count, byte[] Value)>();
        entries.Add((256, 4, 1, WriteInt(width, littleEndian, 4)));            // ImageWidth LONG
        entries.Add((257, 4, 1, WriteInt(height, littleEndian, 4)));           // ImageLength LONG
        if (dpiX is not null || dpiY is not null)
        {
            entries.Add((282, 5, 1, WriteUInt(0xFFFFFFFF, littleEndian, 4)));  // XResolution → offset patched below
            entries.Add((283, 5, 1, WriteUInt(0xFFFFFFFF, littleEndian, 4)));  // YResolution → offset patched below
        }
        if (resolutionUnit is { } unit)
        {
            entries.Add((296, 3, 1, WriteInt(unit, littleEndian, 2)));         // ResolutionUnit SHORT
        }

        var headerSize = 8;
        var entryCountFieldSize = 2;
        var ifdSize = entryCountFieldSize + entries.Count * 12 + 4;
        var rationalOffset = headerSize + ifdSize;

        var stream = new MemoryStream();
        stream.Write(littleEndian ? [0x49, 0x49] : [0x4D, 0x4D]);
        WriteInt(42, stream, littleEndian, 2);
        WriteUInt(8, stream, littleEndian, 4);
        WriteInt(entries.Count, stream, littleEndian, 2);

        var currentRationalOffset = rationalOffset;
        foreach (var entry in entries)
        {
            var value = entry.Value;
            if (entry.Type == 5)
            {
                // RATIONAL value slot stores the offset to the 8-byte rational.
                value = WriteUInt((uint)currentRationalOffset, littleEndian, 4);
                currentRationalOffset += 8;
            }
            WriteInt(entry.Tag, stream, littleEndian, 2);
            WriteInt(entry.Type, stream, littleEndian, 2);
            WriteUInt(entry.Count, stream, littleEndian, 4);
            stream.Write(value);
            for (var i = value.Length; i < 4; i++)
            {
                stream.WriteByte(0);
            }
        }

        WriteUInt(0, stream, littleEndian, 4); // next IFD

        if (dpiX is { } x && dpiY is { } y)
        {
            WriteRational(x, stream, littleEndian);
            WriteRational(y, stream, littleEndian);
        }

        return stream.ToArray();
    }

    /// <summary>Builds an SVG payload with the W3C namespace and a width/height or viewBox.</summary>
    public static byte[] Svg(int width, int height)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(width);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(height);
        return Encoding.UTF8.GetBytes(
            $"<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"{width}\" height=\"{height}\"><rect width=\"100%\" height=\"100%\"/></svg>");
    }

    /// <summary>Builds an SVG payload with only a viewBox (dimensions inferred).</summary>
    public static byte[] SvgViewBox(int width, int height) => Encoding.UTF8.GetBytes(
        $"<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 {width} {height}\"/>");

    /// <summary>Encodes a payload as a base64 data URI with an optional declared media type.</summary>
    public static string DataUriBase64(byte[] bytes, string mediaType = "image/png") =>
        $"data:{mediaType};base64,{Convert.ToBase64String(bytes)}";

    /// <summary>Encodes a payload as a percent-encoded data URI (non-base64 form).</summary>
    public static string DataUriPercentEncoded(byte[] bytes, string mediaType = "image/png")
    {
        var sb = new StringBuilder($"data:{mediaType},");
        foreach (var b in bytes)
        {
            sb.Append('%');
            sb.Append(b.ToString("X2", System.Globalization.CultureInfo.InvariantCulture));
        }
        return sb.ToString();
    }

    /// <summary>Returns the real, well-known 1×1 PNG that ships in <c>examples/Docx/generation/comprehensive.json</c>.</summary>
    public static string NorthwindPngDataUri =>
        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAGUlEQVR4nGOQ96v8TwlmGDVg1IBRA4aLAQAAOOUQamUc8AAAAABJRU5ErkJggg==";

    private static void WriteChunk(Stream stream, string type, byte[] data)
    {
        WriteU32Be(stream, (uint)data.Length);
        var typeBytes = Encoding.ASCII.GetBytes(type);
        stream.Write(typeBytes);
        stream.Write(data);
        var crcInput = new byte[typeBytes.Length + data.Length];
        typeBytes.CopyTo(crcInput, 0);
        data.CopyTo(crcInput, typeBytes.Length);
        WriteU32Be(stream, Crc32(crcInput));
    }

    private static void WriteU16Be(Stream stream, int value)
    {
        Span<byte> buffer = stackalloc byte[2];
        BinaryPrimitives.WriteUInt16BigEndian(buffer, (ushort)value);
        stream.Write(buffer);
    }

    private static void WriteU32Be(Stream stream, uint value)
    {
        Span<byte> buffer = stackalloc byte[4];
        BinaryPrimitives.WriteUInt32BigEndian(buffer, value);
        stream.Write(buffer);
    }

    private static void WriteInt(int value, Stream stream, bool littleEndian, int byteCount)
    {
        var buffer = new byte[byteCount];
        if (littleEndian)
        {
            for (var i = 0; i < byteCount; i++)
            {
                buffer[i] = (byte)(value >> (8 * i));
            }
        }
        else
        {
            for (var i = 0; i < byteCount; i++)
            {
                buffer[byteCount - 1 - i] = (byte)(value >> (8 * i));
            }
        }
        stream.Write(buffer);
    }

    private static void WriteUInt(uint value, Stream stream, bool littleEndian, int byteCount)
    {
        var buffer = new byte[byteCount];
        if (littleEndian)
        {
            for (var i = 0; i < byteCount; i++)
            {
                buffer[i] = (byte)(value >> (8 * i));
            }
        }
        else
        {
            for (var i = 0; i < byteCount; i++)
            {
                buffer[byteCount - 1 - i] = (byte)(value >> (8 * i));
            }
        }
        stream.Write(buffer);
    }

    private static byte[] WriteInt(int value, bool littleEndian, int byteCount)
    {
        var buffer = new byte[byteCount];
        if (littleEndian)
        {
            for (var i = 0; i < byteCount; i++)
            {
                buffer[i] = (byte)(value >> (8 * i));
            }
        }
        else
        {
            for (var i = 0; i < byteCount; i++)
            {
                buffer[byteCount - 1 - i] = (byte)(value >> (8 * i));
            }
        }
        return buffer;
    }

    private static byte[] WriteUInt(uint value, bool littleEndian, int byteCount)
    {
        var buffer = new byte[byteCount];
        if (littleEndian)
        {
            for (var i = 0; i < byteCount; i++)
            {
                buffer[i] = (byte)(value >> (8 * i));
            }
        }
        else
        {
            for (var i = 0; i < byteCount; i++)
            {
                buffer[byteCount - 1 - i] = (byte)(value >> (8 * i));
            }
        }
        return buffer;
    }

    private static void WriteRational(double value, Stream stream, bool littleEndian)
    {
        var numerator = (uint)Math.Round(value * 100, MidpointRounding.AwayFromZero);
        WriteUInt(numerator, stream, littleEndian, 4);
        WriteUInt(100, stream, littleEndian, 4);
    }

    private static uint DpiToPixelsPerMeter(int dpi) =>
        (uint)Math.Round(dpi / 0.0254, MidpointRounding.AwayFromZero);

    private static uint Crc32(ReadOnlySpan<byte> data)
    {
        var crc = 0xFFFFFFFFu;
        foreach (var b in data)
        {
            crc ^= b;
            for (var i = 0; i < 8; i++)
            {
                crc = (crc >> 1) ^ (0xEDB88320u & (0u - (crc & 1)));
            }
        }
        return ~crc;
    }
}
