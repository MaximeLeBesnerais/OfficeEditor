using PptxEditor.Core.Models;

namespace PptxEditor.Core.Converters;

internal static class OpenTypeFontMetricsReader
{
    public static TypstFontMetrics? TryRead(ReadOnlySpan<byte> data)
    {
        if (data.Length < 12) return null;

        var tableCount = ReadUInt16(data, 4);
        var tables = new Dictionary<string, (int Offset, int Length)>(StringComparer.Ordinal);

        for (var i = 0; i < tableCount; i++)
        {
            var recordOffset = 12 + i * 16;
            if (recordOffset + 16 > data.Length) return null;

            var tag = new string([
                (char)data[recordOffset],
                (char)data[recordOffset + 1],
                (char)data[recordOffset + 2],
                (char)data[recordOffset + 3]
            ]);
            var offset = (int)ReadUInt32(data, recordOffset + 8);
            var length = (int)ReadUInt32(data, recordOffset + 12);
            if (offset < 0 || length < 0 || offset + length > data.Length) continue;

            tables[tag] = (offset, length);
        }

        if (!tables.TryGetValue("head", out var head) || head.Length < 20)
            return null;

        var unitsPerEm = ReadUInt16(data, head.Offset + 18);
        if (unitsPerEm == 0) return null;

        short hheaAscender = 0;
        short hheaDescender = 0;
        short hheaLineGap = 0;
        if (tables.TryGetValue("hhea", out var hhea) && hhea.Length >= 10)
        {
            hheaAscender = ReadInt16(data, hhea.Offset + 4);
            hheaDescender = ReadInt16(data, hhea.Offset + 6);
            hheaLineGap = ReadInt16(data, hhea.Offset + 8);
        }

        short typoAscender = hheaAscender;
        short typoDescender = hheaDescender;
        short typoLineGap = hheaLineGap;
        ushort winAscent = (ushort)Math.Max(0, (int)hheaAscender);
        ushort winDescent = (ushort)Math.Max(0, -(int)hheaDescender);

        if (tables.TryGetValue("OS/2", out var os2) && os2.Length >= 78)
        {
            typoAscender = ReadInt16(data, os2.Offset + 68);
            typoDescender = ReadInt16(data, os2.Offset + 70);
            typoLineGap = ReadInt16(data, os2.Offset + 72);
            winAscent = ReadUInt16(data, os2.Offset + 74);
            winDescent = ReadUInt16(data, os2.Offset + 76);
        }

        return new TypstFontMetrics
        {
            UnitsPerEm = unitsPerEm,
            TypoAscender = typoAscender,
            TypoDescender = typoDescender,
            TypoLineGap = typoLineGap,
            HheaAscender = hheaAscender,
            HheaDescender = hheaDescender,
            HheaLineGap = hheaLineGap,
            WinAscent = winAscent,
            WinDescent = winDescent,
            AdvanceWidths = ReadAdvanceWidths(data, tables)
        };
    }

    private static Dictionary<int, ushort> ReadAdvanceWidths(ReadOnlySpan<byte> data, Dictionary<string, (int Offset, int Length)> tables)
    {
        var codePointToGlyph = ReadCharacterMap(data, tables);
        if (codePointToGlyph.Count == 0) return new Dictionary<int, ushort>();

        if (!tables.TryGetValue("hhea", out var hhea) || hhea.Length < 36)
            return new Dictionary<int, ushort>();

        if (!tables.TryGetValue("hmtx", out var hmtx))
            return new Dictionary<int, ushort>();

        if (!tables.TryGetValue("maxp", out var maxp) || maxp.Length < 6)
            return new Dictionary<int, ushort>();

        var numberOfHMetrics = ReadUInt16(data, hhea.Offset + 34);
        var glyphCount = ReadUInt16(data, maxp.Offset + 4);
        if (numberOfHMetrics == 0 || glyphCount == 0 || hmtx.Length < 2)
            return new Dictionary<int, ushort>();

        var advances = new ushort[glyphCount];
        var metricCount = Math.Min(numberOfHMetrics, glyphCount);
        for (var glyphId = 0; glyphId < metricCount; glyphId++)
        {
            var offset = hmtx.Offset + glyphId * 4;
            if (offset + 2 > data.Length) break;
            advances[glyphId] = ReadUInt16(data, offset);
        }

        var lastAdvance = advances[metricCount - 1];
        for (var glyphId = metricCount; glyphId < advances.Length; glyphId++)
        {
            advances[glyphId] = lastAdvance;
        }

        var widths = new Dictionary<int, ushort>(codePointToGlyph.Count);
        foreach (var (codePoint, glyphId) in codePointToGlyph)
        {
            if (glyphId > 0 && glyphId < advances.Length)
            {
                widths[codePoint] = advances[glyphId];
            }
        }

        return widths;
    }

    private static Dictionary<int, ushort> ReadCharacterMap(ReadOnlySpan<byte> data, Dictionary<string, (int Offset, int Length)> tables)
    {
        if (!tables.TryGetValue("cmap", out var cmap) || cmap.Length < 4)
            return new Dictionary<int, ushort>();

        var subtableOffsets = new List<int>();
        var encodingRecordCount = ReadUInt16(data, cmap.Offset + 2);
        for (var i = 0; i < encodingRecordCount; i++)
        {
            var recordOffset = cmap.Offset + 4 + i * 8;
            if (recordOffset + 8 > data.Length) break;

            var subtableOffset = cmap.Offset + (int)ReadUInt32(data, recordOffset + 4);
            if (subtableOffset >= cmap.Offset && subtableOffset + 2 <= cmap.Offset + cmap.Length && !subtableOffsets.Contains(subtableOffset))
            {
                subtableOffsets.Add(subtableOffset);
            }
        }

        var format12 = 0;
        foreach (var offset in subtableOffsets)
        {
            if (ReadUInt16(data, offset) == 12)
            {
                format12 = offset;
                break;
            }
        }
        if (format12 != 0)
        {
            var map = ReadFormat12CharacterMap(data, cmap, format12);
            if (map.Count > 0) return map;
        }

        foreach (var offset in subtableOffsets)
        {
            if (ReadUInt16(data, offset) != 4) continue;

            var map = ReadFormat4CharacterMap(data, cmap, offset);
            if (map.Count > 0) return map;
        }

        return new Dictionary<int, ushort>();
    }

    private static Dictionary<int, ushort> ReadFormat4CharacterMap(ReadOnlySpan<byte> data, (int Offset, int Length) cmap, int offset)
    {
        if (offset + 16 > data.Length) return new Dictionary<int, ushort>();

        var length = ReadUInt16(data, offset + 2);
        if (length < 16 || offset + length > cmap.Offset + cmap.Length || offset + length > data.Length)
            return new Dictionary<int, ushort>();

        var segCount = ReadUInt16(data, offset + 6) / 2;
        if (segCount == 0) return new Dictionary<int, ushort>();

        var endCodeOffset = offset + 14;
        var startCodeOffset = endCodeOffset + segCount * 2 + 2;
        var idDeltaOffset = startCodeOffset + segCount * 2;
        var idRangeOffsetOffset = idDeltaOffset + segCount * 2;
        if (idRangeOffsetOffset + segCount * 2 > offset + length)
            return new Dictionary<int, ushort>();

        var map = new Dictionary<int, ushort>();
        for (var i = 0; i < segCount; i++)
        {
            var endCode = ReadUInt16(data, endCodeOffset + i * 2);
            var startCode = ReadUInt16(data, startCodeOffset + i * 2);
            var idDelta = ReadInt16(data, idDeltaOffset + i * 2);
            var idRangeOffset = ReadUInt16(data, idRangeOffsetOffset + i * 2);

            if (startCode == 0xFFFF && endCode == 0xFFFF) continue;
            if (startCode > endCode) continue;

            for (var codePoint = startCode; codePoint <= endCode; codePoint++)
            {
                ushort glyphId;
                if (idRangeOffset == 0)
                {
                    glyphId = (ushort)((codePoint + idDelta) & 0xFFFF);
                }
                else
                {
                    var glyphOffset = idRangeOffsetOffset + i * 2 + idRangeOffset + (codePoint - startCode) * 2;
                    if (glyphOffset + 2 > offset + length || glyphOffset + 2 > data.Length) continue;

                    var glyphIndex = ReadUInt16(data, glyphOffset);
                    glyphId = glyphIndex == 0
                        ? (ushort)0
                        : (ushort)((glyphIndex + idDelta) & 0xFFFF);
                }

                if (glyphId > 0)
                {
                    map[codePoint] = glyphId;
                }
            }
        }

        return map;
    }

    private static Dictionary<int, ushort> ReadFormat12CharacterMap(ReadOnlySpan<byte> data, (int Offset, int Length) cmap, int offset)
    {
        if (offset + 16 > data.Length) return new Dictionary<int, ushort>();

        var length = ReadUInt32(data, offset + 4);
        if (length < 16 || offset + length > cmap.Offset + cmap.Length || offset + length > data.Length)
            return new Dictionary<int, ushort>();

        var groupCount = ReadUInt32(data, offset + 12);
        var map = new Dictionary<int, ushort>();
        for (var i = 0; i < groupCount; i++)
        {
            var groupOffset = offset + 16 + i * 12;
            if (groupOffset + 12 > offset + length || groupOffset + 12 > data.Length) break;

            var startCode = ReadUInt32(data, groupOffset);
            var endCode = ReadUInt32(data, groupOffset + 4);
            var startGlyphId = ReadUInt32(data, groupOffset + 8);
            if (startCode > endCode || endCode > 0x10FFFF) continue;

            for (var codePoint = startCode; codePoint <= endCode; codePoint++)
            {
                var glyphId = startGlyphId + codePoint - startCode;
                if (glyphId > 0 && glyphId <= ushort.MaxValue)
                {
                    map[(int)codePoint] = (ushort)glyphId;
                }
            }
        }

        return map;
    }

    private static ushort ReadUInt16(ReadOnlySpan<byte> data, int offset)
    {
        return (ushort)((data[offset] << 8) | data[offset + 1]);
    }

    private static short ReadInt16(ReadOnlySpan<byte> data, int offset)
    {
        return unchecked((short)ReadUInt16(data, offset));
    }

    private static uint ReadUInt32(ReadOnlySpan<byte> data, int offset)
    {
        return ((uint)data[offset] << 24) |
               ((uint)data[offset + 1] << 16) |
               ((uint)data[offset + 2] << 8) |
               data[offset + 3];
    }
}
