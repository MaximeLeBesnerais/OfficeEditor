using System.Globalization;
using System.Text.RegularExpressions;

namespace VisualDiff;

/// <summary>
/// Defensive parser for ImageMagick <c>compare -metric RMSE</c> output.
/// Handles the usual <c>2771.24 (0.0422826)</c> shape, <c>0 (0)</c>, scientific
/// notation, a lone number without the parenthesized normalized value, and
/// inf/nan output. Never throws: unparseable input yields null metrics, which
/// callers surface as "metric unavailable" instead of crashing the report.
/// </summary>
internal static partial class RmseParser
{
    private const string Number = @"(?:\d+(?:\.\d+)?|\.\d+)(?:[eE][+-]?\d+)?";

    public static (double? Rmse, double? Normalized) Parse(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return (null, null);
        }

        // Preferred shape: "<rmse> (<normalized>)", e.g. "2771.24 (0.0422826)".
        Match match = RmseWithNormalizedPattern().Match(text);
        if (match.Success)
        {
            return (ParseNumber(match.Groups["rmse"].Value), ParseNumber(match.Groups["norm"].Value));
        }

        // Fallback: a lone number anywhere in the output (some ImageMagick
        // builds print only the raw metric).
        match = LoneNumberPattern().Match(text);
        return match.Success ? (ParseNumber(match.Groups["rmse"].Value), null) : (null, null);
    }

    private static double? ParseNumber(string value)
    {
        if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            || double.IsNaN(parsed)
            || double.IsInfinity(parsed))
        {
            // inf/nan are not JSON-serializable and carry no threshold meaning.
            return null;
        }

        return parsed;
    }

    [GeneratedRegex($@"(?<rmse>{Number})\s*\((?<norm>{Number})\)")]
    private static partial Regex RmseWithNormalizedPattern();

    [GeneratedRegex($@"(?<rmse>{Number})")]
    private static partial Regex LoneNumberPattern();
}
