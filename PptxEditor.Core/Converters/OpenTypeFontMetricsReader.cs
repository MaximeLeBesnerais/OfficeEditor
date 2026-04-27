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
            WinDescent = winDescent
        };
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
