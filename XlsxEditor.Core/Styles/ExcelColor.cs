using XlsxEditor.Core.Exceptions;

namespace XlsxEditor.Core.Styles;

/// <summary>
/// Shared color contract for the XLSX style vocabulary. A color is accepted in one of
/// three forms: a 6-digit RGB hex string ("FF0000"), an 8-digit ARGB hex string
/// ("FFFF0000", the alpha channel is usually FF for opaque), or a common CSS/Excel color
/// name ("red", "white", "darkgray", …). Named colors normalize to a canonical 8-digit
/// ARGB hex value so validation and generation agree on exactly one form, and the
/// stylesheet never sees anything but hex.
///
/// Both the validator (<see cref="IsValid"/>, read-only, for path-qualified diagnostics)
/// and the style builders (<see cref="ToArgb"/>, throwing, for the mutation path) use
/// this single source of truth, so a color that passes semantic validation is always
/// representable at generation time.
/// </summary>
internal static class ExcelColor
{
    /// <summary>
    /// Common CSS/Excel color names → 6-digit RGB hex. The base CSS level-1 palette plus
    /// the widely used extended names; both "gray" and "grey" spellings are accepted.
    /// Anything not listed here must be spelled as explicit hex.
    /// </summary>
    private static readonly Dictionary<string, string> NamedColors = new(StringComparer.Ordinal)
    {
        ["black"] = "000000",
        ["silver"] = "C0C0C0",
        ["gray"] = "808080",
        ["grey"] = "808080",
        ["white"] = "FFFFFF",
        ["maroon"] = "800000",
        ["red"] = "FF0000",
        ["purple"] = "800080",
        ["fuchsia"] = "FF00FF",
        ["magenta"] = "FF00FF",
        ["green"] = "008000",
        ["lime"] = "00FF00",
        ["olive"] = "808000",
        ["yellow"] = "FFFF00",
        ["navy"] = "000080",
        ["blue"] = "0000FF",
        ["teal"] = "008080",
        ["aqua"] = "00FFFF",
        ["cyan"] = "00FFFF",
        ["orange"] = "FFA500",
        ["gold"] = "FFD700",
        ["brown"] = "A52A2A",
        ["pink"] = "FFC0CB",
        ["crimson"] = "DC143C",
        ["violet"] = "EE82EE",
        ["indigo"] = "4B0082",
        ["orchid"] = "DA70D6",
        ["plum"] = "DDA0DD",
        ["turquoise"] = "40E0D0",
        ["khaki"] = "F0E68C",
        ["lavender"] = "E6E6FA",
        ["beige"] = "F5F5DC",
        ["ivory"] = "FFFFF0",
        ["coral"] = "FF7F50",
        ["salmon"] = "FA8072",
        ["tomato"] = "FF6347",
        ["chocolate"] = "D2691E",
        ["tan"] = "D2B48C",
        ["peru"] = "CD853F",
        ["sienna"] = "A0522D",
        ["darkred"] = "8B0000",
        ["darkgreen"] = "006400",
        ["darkblue"] = "00008B",
        ["darkcyan"] = "008B8B",
        ["darkmagenta"] = "8B008B",
        ["darkorange"] = "FF8C00",
        ["darkgray"] = "A9A9A9",
        ["darkgrey"] = "A9A9A9",
        ["dimgray"] = "696969",
        ["dimgrey"] = "696969",
        ["lightgray"] = "D3D3D3",
        ["lightgrey"] = "D3D3D3",
        ["lightblue"] = "ADD8E6",
        ["lightgreen"] = "90EE90",
        ["lightcyan"] = "E0FFFF",
        ["lightpink"] = "FFB6C1",
        ["lightyellow"] = "FFFFE0",
        ["mediumblue"] = "0000CD",
        ["mediumseagreen"] = "3CB371",
        ["mediumspringgreen"] = "00FA9A",
        ["mediumturquoise"] = "48D1CC",
        ["mediumvioletred"] = "C71585",
        ["firebrick"] = "B22222",
        ["forestgreen"] = "228B22",
        ["seagreen"] = "2E8B57",
        ["springgreen"] = "00FF7F",
        ["skyblue"] = "87CEEB",
        ["steelblue"] = "4682B4",
        ["royalblue"] = "4169E1",
        ["slateblue"] = "6A5ACD",
        ["slategray"] = "708090",
        ["slategrey"] = "708090",
        ["lightslategray"] = "778899",
        ["lightslategrey"] = "778899",
        ["gainsboro"] = "DCDCDC",
        ["whitesmoke"] = "F5F5F5",
        ["mintcream"] = "F5FFFA",
        ["honeydew"] = "F0FFF0",
        ["azure"] = "F0FFFF",
        ["snow"] = "FFFAFA",
        ["seashell"] = "FFF5EE",
        ["wheat"] = "F5DEB3",
        ["yellowgreen"] = "9ACD32",
        ["olivedrab"] = "6B8E23",
        ["darkolivegreen"] = "556B2F",
        ["darkgoldenrod"] = "B8860B",
        ["goldenrod"] = "DAA520",
        ["darkkhaki"] = "BDB76B",
        ["lightsalmon"] = "FFA07A",
        ["darksalmon"] = "E9967A",
        ["hotpink"] = "FF69B4",
        ["deeppink"] = "FF1493",
        ["palevioletred"] = "DB7093"
    };

    /// <summary>True when <paramref name="color"/> is a 6-digit RGB or 8-digit ARGB hex string.</summary>
    internal static bool IsHex(string color) =>
        (color.Length == 6 || color.Length == 8) && color.All(Uri.IsHexDigit);

    /// <summary>
    /// True when <paramref name="color"/> is a valid hex color or a known common color name.
    /// The validator uses this so a name that passes semantic validation always normalizes
    /// at generation time — validation and execution can never disagree about a color.
    /// </summary>
    internal static bool IsValid(string? color) =>
        color is not null
        && (IsHex(color.Trim()) || NamedColors.ContainsKey(color.Trim().ToLowerInvariant()));

    /// <summary>
    /// Normalizes a hex or named color to its canonical 8-digit ARGB form (a 6-digit RGB
    /// gains an opaque "FF" alpha prefix, names resolve to their RGB hex, 8-digit hex is
    /// upper-cased). Throws <see cref="XlsxException"/> for anything that is neither a
    /// known name nor a 6/8-digit hex string.
    /// </summary>
    internal static string ToArgb(string color, string context)
    {
        var trimmed = color.Trim();
        if (NamedColors.TryGetValue(trimmed.ToLowerInvariant(), out var rgb))
        {
            return "FF" + rgb;
        }

        if (!IsHex(trimmed))
        {
            throw new XlsxException(
                $"{context} color '{color}' must be a 6-digit RGB (e.g. 'FF0000'), an 8-digit " +
                "ARGB (e.g. 'FFFF0000') hex string, or a common color name (e.g. 'red', " +
                "'white', 'darkgray').");
        }

        return trimmed.Length == 6 ? "FF" + trimmed : trimmed.ToUpperInvariant();
    }
}
