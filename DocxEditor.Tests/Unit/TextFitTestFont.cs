using System.Text;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Builds minimal synthetic TTF files with fully controlled metrics for
/// deterministic wrap-simulation tests. Modeled on the sfnt helpers in
/// OpenTypeFontMetricsReaderTests, extended with a name table and a
/// configurable uniform advance width.
///
/// Default face: unitsPerEm=1000, every printable ASCII char (0x20–0x7E) has
/// the same advance (default 500 → char width = fontSize/2), and
/// TypoAsc−TypoDesc+TypoLineGap = 800+200+200 → default line factor 1.2.
/// </summary>
internal static class TextFitTestFont
{
    public const int UnitsPerEm = 1000;

    public static string WriteFont(string directory, string fileName, string family, ushort advance = 500)
    {
        var path = Path.Combine(directory, fileName);
        File.WriteAllBytes(path, BuildFont(family, advance));
        return path;
    }

    public static byte[] BuildFont(string family, ushort advance = 500)
    {
        const ushort glyphCount = 96; // glyphs 0..95; 0x20..0x7E map to 1..95
        return BuildSfnt(
            ("head", HeadTable(UnitsPerEm)),
            ("hhea", HheaTable(800, -200, 0, glyphCount)),
            ("OS/2", Os2Table(800, -200, 200, 900, 300)),
            ("maxp", MaxpTable(glyphCount)),
            ("hmtx", HmtxTable(glyphCount, advance)),
            ("cmap", CmapTable(Format4AsciiSubtable())),
            ("name", NameTable(family)));
    }

    /// <summary>Minimal TrueType Collection bytes — enough for magic detection, never parseable.</summary>
    public static string WriteTrueTypeCollection(string directory, string fileName = "Fake.ttc")
    {
        var path = Path.Combine(directory, fileName);
        var data = new byte[24];
        Encoding.ASCII.GetBytes("ttcf", data.AsSpan(0, 4));
        WriteUInt32(data, 4, 0x00010000); // version
        WriteUInt32(data, 8, 1);          // numFonts
        WriteUInt32(data, 12, 12);        // first offset (bogus on purpose)
        File.WriteAllBytes(path, data);
        return path;
    }

    private static byte[] BuildSfnt(params (string Tag, byte[] Data)[] tables)
    {
        var directoryLength = 12 + tables.Length * 16;
        var totalLength = directoryLength + tables.Sum(t => t.Data.Length);
        var data = new byte[totalLength];
        WriteUInt32(data, 0, 0x00010000);
        WriteUInt16(data, 4, (ushort)tables.Length);

        var tableOffset = directoryLength;
        for (var i = 0; i < tables.Length; i++)
        {
            var recordOffset = 12 + i * 16;
            Encoding.ASCII.GetBytes(tables[i].Tag, data.AsSpan(recordOffset, 4));
            WriteUInt32(data, recordOffset + 8, (uint)tableOffset);
            WriteUInt32(data, recordOffset + 12, (uint)tables[i].Data.Length);
            tables[i].Data.CopyTo(data.AsSpan(tableOffset));
            tableOffset += tables[i].Data.Length;
        }
        return data;
    }

    private static byte[] HeadTable(ushort unitsPerEm)
    {
        var data = new byte[20];
        WriteUInt16(data, 18, unitsPerEm);
        return data;
    }

    private static byte[] HheaTable(short ascender, short descender, short lineGap, ushort numberOfHMetrics)
    {
        var data = new byte[36];
        WriteInt16(data, 4, ascender);
        WriteInt16(data, 6, descender);
        WriteInt16(data, 8, lineGap);
        WriteUInt16(data, 34, numberOfHMetrics);
        return data;
    }

    private static byte[] Os2Table(short typoAscender, short typoDescender, short typoLineGap, ushort winAscent, ushort winDescent)
    {
        var data = new byte[78];
        WriteInt16(data, 68, typoAscender);
        WriteInt16(data, 70, typoDescender);
        WriteInt16(data, 72, typoLineGap);
        WriteUInt16(data, 74, winAscent);
        WriteUInt16(data, 76, winDescent);
        return data;
    }

    private static byte[] MaxpTable(ushort glyphCount)
    {
        var data = new byte[6];
        WriteUInt16(data, 4, glyphCount);
        return data;
    }

    private static byte[] HmtxTable(ushort glyphCount, ushort advance)
    {
        var data = new byte[glyphCount * 4];
        for (var i = 0; i < glyphCount; i++)
        {
            WriteUInt16(data, i * 4, advance);
        }
        return data;
    }

    private static byte[] CmapTable(byte[] subtable)
    {
        var data = new byte[12 + subtable.Length];
        WriteUInt16(data, 2, 1);                    // one encoding record
        WriteUInt16(data, 4, 3);                    // platform: Windows
        WriteUInt16(data, 6, 1);                    // encoding: Unicode BMP
        WriteUInt32(data, 8, 12);
        subtable.CopyTo(data.AsSpan(12));
        return data;
    }

    // Format 4 with a single segment covering 0x20–0x7E (glyph = cp − 0x1F) + terminator.
    private static byte[] Format4AsciiSubtable()
    {
        const ushort segCount = 2;
        var length = (ushort)(16 + segCount * 8);
        var data = new byte[length];
        WriteUInt16(data, 0, 4);
        WriteUInt16(data, 2, length);
        WriteUInt16(data, 6, segCount * 2);

        var endCodeOffset = 14;
        var startCodeOffset = endCodeOffset + segCount * 2 + 2;
        var idDeltaOffset = startCodeOffset + segCount * 2;
        var idRangeOffsetOffset = idDeltaOffset + segCount * 2;

        WriteUInt16(data, endCodeOffset, 0x7E);
        WriteUInt16(data, endCodeOffset + 2, 0xFFFF);
        WriteUInt16(data, startCodeOffset, 0x20);
        WriteUInt16(data, startCodeOffset + 2, 0xFFFF);
        WriteInt16(data, idDeltaOffset, -0x1F);
        WriteInt16(data, idDeltaOffset + 2, 1);
        WriteUInt16(data, idRangeOffsetOffset, 0);
        WriteUInt16(data, idRangeOffsetOffset + 2, 0);
        return data;
    }

    private static byte[] NameTable(string family)
    {
        var value = Encoding.BigEndianUnicode.GetBytes(family + "\0");
        var data = new byte[6 + 12 + value.Length];
        WriteUInt16(data, 2, 1);                    // one record
        WriteUInt16(data, 4, 18);                   // string storage offset
        WriteUInt16(data, 6, 3);                    // platform: Windows
        WriteUInt16(data, 8, 1);                    // encoding: Unicode BMP
        WriteUInt16(data, 10, 0x0409);              // language: en-US
        WriteUInt16(data, 12, 1);                   // nameID 1 = family
        WriteUInt16(data, 14, (ushort)value.Length);
        WriteUInt16(data, 16, 0);
        value.CopyTo(data.AsSpan(18));
        return data;
    }

    private static void WriteUInt16(byte[] data, int offset, ushort value)
    {
        data[offset] = (byte)(value >> 8);
        data[offset + 1] = (byte)value;
    }

    private static void WriteInt16(byte[] data, int offset, short value)
        => WriteUInt16(data, offset, unchecked((ushort)value));

    private static void WriteUInt32(byte[] data, int offset, uint value)
    {
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }
}
