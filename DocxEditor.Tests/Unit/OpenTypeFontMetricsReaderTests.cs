using System.Reflection;
using System.Text;
using PptxEditor.Core.Models;

namespace DocxEditor.Tests.Unit;

public sealed class OpenTypeFontMetricsReaderTests : IDisposable
{
    private static readonly Type ReaderType = Type.GetType(
        "PptxEditor.Core.Converters.OpenTypeFontMetricsReader, PptxEditor.Core",
        throwOnError: true)!;

    private readonly string _tempDir = Path.Combine(Path.GetTempPath(), "OpenTypeFontMetricsReaderTests", Guid.NewGuid().ToString("N"));

    public OpenTypeFontMetricsReaderTests()
    {
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDir))
                Directory.Delete(_tempDir, recursive: true);
        }
        catch
        {
            // Best effort cleanup.
        }
    }

    [Fact]
    public void ReadMetrics_ReturnsNull_ForMissingOrTooShortFonts()
    {
        Assert.Null(ReadMetrics(Path.Combine(_tempDir, "missing.ttf")));

        var path = WriteFont("short.ttf", [0, 1, 2, 3, 4]);

        Assert.Null(ReadMetrics(path));
    }

    [Fact]
    public void ReadMetrics_ReturnsNull_ForTruncatedDirectoryOrMissingHead()
    {
        var truncatedDirectory = new byte[12];
        WriteUInt16(truncatedDirectory, 4, 1);

        Assert.Null(ReadMetrics(WriteFont("truncated-directory.ttf", truncatedDirectory)));

        var missingHead = BuildSfnt(("hhea", HheaTable(700, -200, 30, 1)));

        Assert.Null(ReadMetrics(WriteFont("missing-head.ttf", missingHead)));
    }

    [Fact]
    public void ReadMetrics_ReturnsNull_WhenHeadIsTooShortOrUnitsPerEmIsZero()
    {
        var shortHead = BuildSfnt(("head", new byte[19]));
        Assert.Null(ReadMetrics(WriteFont("short-head.ttf", shortHead)));

        var zeroUnitsPerEmHead = HeadTable(0);
        Assert.Null(ReadMetrics(WriteFont("zero-upem.ttf", BuildSfnt(("head", zeroUnitsPerEmHead)))));
    }

    [Fact]
    public void ReadMetrics_UsesHheaValues_WhenOs2TableIsMissingOrTooShort()
    {
        var data = BuildSfnt(
            ("head", HeadTable(2048)),
            ("hhea", HheaTable(900, -250, 75, 1)),
            ("OS/2", new byte[77]));

        TypstFontMetrics? metrics = ReadMetrics(WriteFont("hhea-fallback.ttf", data));
        Assert.NotNull(metrics);

        Assert.Equal(2048, metrics.UnitsPerEm);
        Assert.Equal(900, metrics.HheaAscender);
        Assert.Equal(-250, metrics.HheaDescender);
        Assert.Equal(75, metrics.HheaLineGap);
        Assert.Equal(900, metrics.TypoAscender);
        Assert.Equal(-250, metrics.TypoDescender);
        Assert.Equal(75, metrics.TypoLineGap);
        Assert.Equal((ushort)900, metrics.WinAscent);
        Assert.Equal((ushort)250, metrics.WinDescent);
        Assert.Empty(metrics.AdvanceWidths);
    }

    [Fact]
    public void ReadMetrics_UsesOs2Metrics_WhenTableIsPresent()
    {
        var data = BuildSfnt(
            ("head", HeadTable(1000)),
            ("hhea", HheaTable(800, -200, 20, 1)),
            ("OS/2", Os2Table(710, -190, 40, 930, 270)));

        TypstFontMetrics? metrics = ReadMetrics(WriteFont("os2.ttf", data));
        Assert.NotNull(metrics);

        Assert.Equal(710, metrics.TypoAscender);
        Assert.Equal(-190, metrics.TypoDescender);
        Assert.Equal(40, metrics.TypoLineGap);
        Assert.Equal((ushort)930, metrics.WinAscent);
        Assert.Equal((ushort)270, metrics.WinDescent);
    }

    [Fact]
    public void ReadMetrics_ReadsFormat4CharacterMapAndAdvanceWidths()
    {
        var data = BuildSfnt(
            ("head", HeadTable(1000)),
            ("hhea", HheaTable(800, -200, 0, numberOfHMetrics: 2)),
            ("maxp", MaxpTable(glyphCount: 4)),
            ("hmtx", HmtxTable(500, 700)),
            ("cmap", CmapTable(Format4Subtable())));

        TypstFontMetrics? metrics = ReadMetrics(WriteFont("format4.ttf", data));
        Assert.NotNull(metrics);

        Assert.Equal((ushort)700, metrics.AdvanceWidths[65]);
        Assert.Equal((ushort)700, metrics.AdvanceWidths[66]);
        Assert.False(metrics.AdvanceWidths.ContainsKey(0xFFFF));
    }

    [Fact]
    public void ReadMetrics_ReadsFormat12CharacterMapBeforeFormat4AndSkipsInvalidGroups()
    {
        var format12 = Format12Subtable(
            (0x1F600u, 0x1F601u, 1u),
            (20u, 10u, 1u),
            (0x110000u, 0x110000u, 1u));
        var data = BuildSfnt(
            ("head", HeadTable(1000)),
            ("hhea", HheaTable(800, -200, 0, numberOfHMetrics: 3)),
            ("maxp", MaxpTable(glyphCount: 3)),
            ("hmtx", HmtxTable(400, 600, 800)),
            ("cmap", CmapTable(format12, Format4Subtable())));

        TypstFontMetrics? metrics = ReadMetrics(WriteFont("format12.ttf", data));
        Assert.NotNull(metrics);

        Assert.Equal((ushort)600, metrics.AdvanceWidths[0x1F600]);
        Assert.Equal((ushort)800, metrics.AdvanceWidths[0x1F601]);
        Assert.False(metrics.AdvanceWidths.ContainsKey(65));
    }

    [Fact]
    public void ReadMetrics_ReturnsEmptyWidths_ForMalformedOrIncompleteWidthTables()
    {
        var malformedCmap = BuildSfnt(
            ("head", HeadTable(1000)),
            ("hhea", HheaTable(800, -200, 0, 1)),
            ("maxp", MaxpTable(2)),
            ("hmtx", HmtxTable(500)),
            ("cmap", CmapTable(Format4Subtable(lengthOverride: 10))));
        TypstFontMetrics? badCmapMetrics = ReadMetrics(WriteFont("bad-cmap.ttf", malformedCmap));
        Assert.NotNull(badCmapMetrics);
        Assert.Empty(badCmapMetrics.AdvanceWidths);

        var missingHmtx = BuildSfnt(
            ("head", HeadTable(1000)),
            ("hhea", HheaTable(800, -200, 0, 1)),
            ("maxp", MaxpTable(2)),
            ("cmap", CmapTable(Format4Subtable())));
        TypstFontMetrics? missingHmtxMetrics = ReadMetrics(WriteFont("missing-hmtx.ttf", missingHmtx));
        Assert.NotNull(missingHmtxMetrics);
        Assert.Empty(missingHmtxMetrics.AdvanceWidths);
    }

    [Fact]
    public void ReadFontFamilyName_ReturnsNull_ForMissingShortOrMalformedNameTables()
    {
        Assert.Null(ReadFontFamilyName(Path.Combine(_tempDir, "missing.ttf")));
        Assert.Null(ReadFontFamilyName(WriteFont("short-name.ttf", [1, 2, 3])));
        Assert.Null(ReadFontFamilyName(WriteFont("no-name.ttf", BuildSfnt(("head", HeadTable(1000))))));
        Assert.Null(ReadFontFamilyName(WriteFont("short-name-table.ttf", BuildSfnt(("name", new byte[5])))));

        var truncatedDirectory = new byte[12];
        WriteUInt16(truncatedDirectory, 4, 1);
        Assert.Null(ReadFontFamilyName(WriteFont("truncated-name-directory.ttf", truncatedDirectory)));
    }

    [Fact]
    public void ReadFontFamilyName_ReadsUnicodeFamilyAndPreferredFamilyTakesPriority()
    {
        var family = NameRecord(3, 1, 0x0409, 1, Encoding.BigEndianUnicode.GetBytes("Fallback\0"));
        var skipped = NameRecord(3, 1, 0x0409, 2, Encoding.BigEndianUnicode.GetBytes("Subfamily"));
        var preferred = NameRecord(3, 10, 0x0409, 16, Encoding.BigEndianUnicode.GetBytes("Preferred\0"));

        var data = BuildSfnt(("name", NameTable(family, skipped, preferred)));

        Assert.Equal("Preferred", ReadFontFamilyName(WriteFont("preferred-name.ttf", data)));
    }

    [Fact]
    public void ReadFontFamilyName_ReadsUtf8FamilyAndSkipsOutOfRangeStrings()
    {
        var outOfRange = new NameRecordSpec(1, 0, 0, 1, Encoding.UTF8.GetBytes("Ignored"), ForcedOffset: 200);
        var family = NameRecord(1, 0, 0, 1, Encoding.UTF8.GetBytes("Utf8 Family\0"));

        var data = BuildSfnt(("name", NameTable(outOfRange, family)));

        Assert.Equal("Utf8 Family", ReadFontFamilyName(WriteFont("utf8-name.ttf", data)));
    }

    private TypstFontMetrics? ReadMetrics(string path)
    {
        var method = ReaderType.GetMethod("ReadMetrics", BindingFlags.Public | BindingFlags.Static)!;
        return (TypstFontMetrics?)method.Invoke(null, [path]);
    }

    private static string? ReadFontFamilyName(string path)
    {
        var method = ReaderType.GetMethod("ReadFontFamilyName", BindingFlags.Public | BindingFlags.Static)!;
        return (string?)method.Invoke(null, [path]);
    }

    private string WriteFont(string name, byte[] data)
    {
        var path = Path.Combine(_tempDir, name);
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

    private static byte[] HmtxTable(params ushort[] advances)
    {
        var data = new byte[advances.Length * 4];
        for (var i = 0; i < advances.Length; i++)
            WriteUInt16(data, i * 4, advances[i]);
        return data;
    }

    private static byte[] CmapTable(params byte[][] subtables)
    {
        var headerLength = 4 + subtables.Length * 8;
        var data = new byte[headerLength + subtables.Sum(s => s.Length)];
        WriteUInt16(data, 2, (ushort)subtables.Length);

        var subtableOffset = headerLength;
        for (var i = 0; i < subtables.Length; i++)
        {
            var recordOffset = 4 + i * 8;
            WriteUInt16(data, recordOffset, i == 0 ? (ushort)3 : (ushort)0);
            WriteUInt16(data, recordOffset + 2, i == 0 ? (ushort)10 : (ushort)3);
            WriteUInt32(data, recordOffset + 4, (uint)subtableOffset);
            subtables[i].CopyTo(data.AsSpan(subtableOffset));
            subtableOffset += subtables[i].Length;
        }

        return data;
    }

    private static byte[] Format4Subtable(ushort? lengthOverride = null)
    {
        const ushort segCount = 2;
        var length = (ushort)(16 + segCount * 8);
        var data = new byte[length];
        WriteUInt16(data, 0, 4);
        WriteUInt16(data, 2, lengthOverride ?? length);
        WriteUInt16(data, 6, segCount * 2);

        var endCodeOffset = 14;
        var startCodeOffset = endCodeOffset + segCount * 2 + 2;
        var idDeltaOffset = startCodeOffset + segCount * 2;
        var idRangeOffsetOffset = idDeltaOffset + segCount * 2;

        WriteUInt16(data, endCodeOffset, 66);
        WriteUInt16(data, endCodeOffset + 2, 0xFFFF);
        WriteUInt16(data, startCodeOffset, 65);
        WriteUInt16(data, startCodeOffset + 2, 0xFFFF);
        WriteInt16(data, idDeltaOffset, -64);
        WriteInt16(data, idDeltaOffset + 2, 1);
        WriteUInt16(data, idRangeOffsetOffset, 0);
        WriteUInt16(data, idRangeOffsetOffset + 2, 0);
        return data;
    }

    private static byte[] Format12Subtable(params (uint StartCode, uint EndCode, uint StartGlyphId)[] groups)
    {
        var length = (uint)(16 + groups.Length * 12);
        var data = new byte[length];
        WriteUInt16(data, 0, 12);
        WriteUInt32(data, 4, length);
        WriteUInt32(data, 12, (uint)groups.Length);
        for (var i = 0; i < groups.Length; i++)
        {
            var offset = 16 + i * 12;
            WriteUInt32(data, offset, groups[i].StartCode);
            WriteUInt32(data, offset + 4, groups[i].EndCode);
            WriteUInt32(data, offset + 8, groups[i].StartGlyphId);
        }

        return data;
    }

    private static NameRecordSpec NameRecord(ushort platformId, ushort encodingId, ushort languageId, ushort nameId, byte[] value)
    {
        return new NameRecordSpec(platformId, encodingId, languageId, nameId, value);
    }

    private static byte[] NameTable(params NameRecordSpec[] records)
    {
        var stringOffset = 6 + records.Length * 12;
        var stringLength = records.Where(r => r.ForcedOffset is null).Sum(r => r.Value.Length);
        var data = new byte[stringOffset + stringLength];
        WriteUInt16(data, 2, (ushort)records.Length);
        WriteUInt16(data, 4, (ushort)stringOffset);

        var nextStringOffset = 0;
        for (var i = 0; i < records.Length; i++)
        {
            var recordOffset = 6 + i * 12;
            var record = records[i];
            var offset = record.ForcedOffset ?? nextStringOffset;
            WriteUInt16(data, recordOffset, record.PlatformId);
            WriteUInt16(data, recordOffset + 2, record.EncodingId);
            WriteUInt16(data, recordOffset + 4, record.LanguageId);
            WriteUInt16(data, recordOffset + 6, record.NameId);
            WriteUInt16(data, recordOffset + 8, (ushort)record.Value.Length);
            WriteUInt16(data, recordOffset + 10, (ushort)offset);

            if (record.ForcedOffset is null)
            {
                record.Value.CopyTo(data.AsSpan(stringOffset + nextStringOffset));
                nextStringOffset += record.Value.Length;
            }
        }

        return data;
    }

    private static void WriteUInt16(byte[] data, int offset, ushort value)
    {
        data[offset] = (byte)(value >> 8);
        data[offset + 1] = (byte)value;
    }

    private static void WriteInt16(byte[] data, int offset, short value)
    {
        WriteUInt16(data, offset, unchecked((ushort)value));
    }

    private static void WriteUInt32(byte[] data, int offset, uint value)
    {
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }

    private sealed record NameRecordSpec(
        ushort PlatformId,
        ushort EncodingId,
        ushort LanguageId,
        ushort NameId,
        byte[] Value,
        int? ForcedOffset = null);
}
