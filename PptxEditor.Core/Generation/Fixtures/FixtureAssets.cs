namespace PptxEditor.Core.Generation.Fixtures;

/// <summary>
/// Deterministic, dependency-free image assets for parity fixtures. The checkerboard PNG
/// is real, renderable content (both PowerPoint and Typst can rasterize it) produced with
/// an uncompressed zlib stream so the bytes are identical on every run and machine —
/// the same source pixels feed both renderers, which is what makes the diff meaningful.
/// </summary>
public static class FixtureAssets
{
    /// <summary>File name the fixture generator materializes under each fixture's assets/ directory.</summary>
    public const string CheckerboardFileName = "checkerboard.png";

    /// <summary>Source image size in pixels. 3:2 aspect so fill/contain/crop visibly diverge on non-3:2 frames.</summary>
    public const int CheckerboardWidth = 240;

    /// <summary>See <see cref="CheckerboardWidth"/>.</summary>
    public const int CheckerboardHeight = 160;

    private static readonly (byte R, byte G, byte B)[] QuadrantColors =
    [
        (11, 61, 145),   // primary
        (255, 107, 0),   // accent
        (138, 148, 166), // muted
        (26, 26, 26)     // ink
    ];

    /// <summary>
    /// Builds the checkerboard PNG: four unequal quadrants (so crop direction is
    /// detectable) separated by a white cross. Deterministic: fixed pixels, fixed
    /// uncompressed zlib encoding, hand-rolled CRC32/Adler32.
    /// </summary>
    public static byte[] CreateCheckerboardPng(int width = CheckerboardWidth, int height = CheckerboardHeight)
    {
        if (width <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(width), "Image width must be positive.");
        }
        if (height <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(height), "Image height must be positive.");
        }

        // 60/40 split makes the quadrants unequal, so a wrong crop origin shows up in a diff.
        int splitX = width * 3 / 5;
        int splitY = height * 3 / 5;
        const int crossHalf = 2;

        var raw = new byte[height * (1 + width * 3)];
        for (int y = 0; y < height; y++)
        {
            int row = y * (1 + width * 3);
            raw[row] = 0; // PNG filter: none
            for (int x = 0; x < width; x++)
            {
                bool onCross = Math.Abs(x - splitX) <= crossHalf || Math.Abs(y - splitY) <= crossHalf;
                int quadrant = (y < splitY ? 0 : 2) + (x < splitX ? 0 : 1);
                (byte r, byte g, byte b) = onCross ? ((byte)255, (byte)255, (byte)255) : QuadrantColors[quadrant];
                int px = row + 1 + x * 3;
                raw[px] = r;
                raw[px + 1] = g;
                raw[px + 2] = b;
            }
        }

        using var output = new MemoryStream();
        output.Write([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);

        var ihdr = new byte[13];
        WriteBigEndian(ihdr, 0, (uint)width);
        WriteBigEndian(ihdr, 4, (uint)height);
        ihdr[8] = 8;  // bit depth
        ihdr[9] = 2;  // color type: truecolor RGB
        // ihdr[10..12] compression/filter/interlace = 0
        WriteChunk(output, "IHDR", ihdr);
        WriteChunk(output, "IDAT", ZlibStore(raw));
        WriteChunk(output, "IEND", []);
        return output.ToArray();
    }

    /// <summary>zlib stream (RFC 1950) wrapping uncompressed deflate stored blocks — fully deterministic.</summary>
    private static byte[] ZlibStore(byte[] data)
    {
        using var stream = new MemoryStream();
        stream.WriteByte(0x78); // CMF: deflate, 32K window
        stream.WriteByte(0x01); // FLG: no preset dictionary, check bits for 0x7801

        int offset = 0;
        do
        {
            int length = Math.Min(data.Length - offset, 65535);
            bool final = offset + length == data.Length;
            stream.WriteByte(final ? (byte)0x01 : (byte)0x00);
            stream.WriteByte((byte)(length & 0xFF));
            stream.WriteByte((byte)(length >> 8));
            stream.WriteByte((byte)(~length & 0xFF));
            stream.WriteByte((byte)(~length >> 8 & 0xFF));
            stream.Write(data, offset, length);
            offset += length;
        }
        while (offset < data.Length);

        WriteBigEndian(stream, Adler32(data));
        return stream.ToArray();
    }

    private static uint Adler32(ReadOnlySpan<byte> data)
    {
        const uint mod = 65521;
        uint a = 1, b = 0;
        foreach (byte value in data)
        {
            a = (a + value) % mod;
            b = (b + a) % mod;
        }
        return (b << 16) | a;
    }

    private static void WriteChunk(Stream output, string type, byte[] data)
    {
        WriteBigEndian(output, (uint)data.Length);
        var typeBytes = System.Text.Encoding.ASCII.GetBytes(type);
        output.Write(typeBytes);
        output.Write(data);
        var crcInput = new byte[typeBytes.Length + data.Length];
        typeBytes.CopyTo(crcInput, 0);
        data.CopyTo(crcInput, typeBytes.Length);
        WriteBigEndian(output, Crc32(crcInput));
    }

    private static uint Crc32(ReadOnlySpan<byte> data)
    {
        uint crc = 0xFFFFFFFF;
        foreach (byte value in data)
        {
            crc ^= value;
            for (int bit = 0; bit < 8; bit++)
            {
                crc = (crc >> 1) ^ (0xEDB88320 & (uint)-(int)(crc & 1));
            }
        }
        return crc ^ 0xFFFFFFFF;
    }

    private static void WriteBigEndian(Stream stream, uint value)
    {
        stream.WriteByte((byte)(value >> 24));
        stream.WriteByte((byte)(value >> 16));
        stream.WriteByte((byte)(value >> 8));
        stream.WriteByte((byte)value);
    }

    private static void WriteBigEndian(byte[] buffer, int offset, uint value)
    {
        buffer[offset] = (byte)(value >> 24);
        buffer[offset + 1] = (byte)(value >> 16);
        buffer[offset + 2] = (byte)(value >> 8);
        buffer[offset + 3] = (byte)value;
    }
}
