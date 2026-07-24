using System.Globalization;

namespace PptxEditor.Core.Converters.Charts;

/// <summary>
/// Auto value-axis scaling with PowerPoint-style nice rounding. Pure static math so the
/// bar/column builders (and round-2 line charts) share one implementation.
/// </summary>
public static class ChartAxisScale
{
    /// <summary>Target number of major intervals for an auto axis.</summary>
    private const double TargetIntervals = 5.0;

    /// <summary>
    /// Computes axis min/max/major unit from data bounds. Explicit values from
    /// <c>c:scaling</c>/<c>c:majorUnit</c> win; otherwise the axis is nice-rounded.
    /// When data labels are shown, PowerPoint extends the axis past the data max so the
    /// end labels fit (observed on both REF decks: data max 100 → axis 120, 48 → 60);
    /// approximated with half a major unit of headroom.
    /// </summary>
    public static (double Min, double Max, double Unit) Compute(
        double dataMin, double dataMax, bool extendForLabels,
        double? explicitMin = null, double? explicitMax = null, double? explicitUnit = null)
    {
        if (dataMax < dataMin)
            (dataMin, dataMax) = (dataMax, dataMin);

        var unit = explicitUnit ?? ComputeAutoUnit(dataMin, dataMax);
        if (unit <= 0)
            unit = 1.0;

        var min = explicitMin ?? ComputeAutoMin(dataMin, unit);
        var max = explicitMax ?? ComputeAutoMax(dataMax, min, unit, extendForLabels);

        if (max <= min)
            max = min + unit;

        return (min, max, unit);
    }

    /// <summary>Major tick values from <paramref name="min"/> to <paramref name="max"/> inclusive.</summary>
    public static IReadOnlyList<double> Ticks(double min, double max, double unit)
    {
        var ticks = new List<double>();
        if (unit <= 0)
            return ticks;

        // Guard against pathological ranges (unit tiny relative to the span).
        var count = (int)Math.Ceiling((max - min) / unit);
        count = Math.Min(count, 1000);
        for (var i = 0; i <= count; i++)
        {
            var tick = min + i * unit;
            if (tick > max + unit * 1e-6)
                break;
            ticks.Add(tick);
        }
        return ticks;
    }

    /// <summary>
    /// Smallest "nice" number (1/2/2.5/5/10 × 10^k) greater than or equal to
    /// <paramref name="value"/>. Assumes <paramref name="value"/> &gt; 0.
    /// </summary>
    public static double NiceCeiling(double value)
    {
        if (value <= 0 || double.IsNaN(value) || double.IsInfinity(value))
            return 1.0;

        var step = Math.Pow(10, Math.Floor(Math.Log10(value)));
        foreach (var multiplier in new[] { 1.0, 2.0, 2.5, 5.0, 10.0 })
        {
            var candidate = multiplier * step;
            if (candidate >= value - step * 1e-9)
                return candidate;
        }
        return 10.0 * step;
    }

    private static double ComputeAutoUnit(double dataMin, double dataMax)
    {
        var span = dataMax - Math.Min(dataMin, 0.0);
        if (span <= 0)
            span = Math.Abs(dataMax) > 0 ? Math.Abs(dataMax) : 1.0;
        return NiceCeiling(span / TargetIntervals);
    }

    private static double ComputeAutoMin(double dataMin, double unit)
    {
        if (dataMin >= 0)
            return 0.0;
        // Extend below zero to the next nice multiple of the unit.
        return -unit * Math.Ceiling(-dataMin / unit - 1e-9);
    }

    private static double ComputeAutoMax(double dataMax, double min, double unit, bool extendForLabels)
    {
        var effective = dataMax + (extendForLabels ? unit * 0.5 : 0.0);
        var intervals = Math.Ceiling((effective - min) / unit - 1e-9);
        if (intervals < 1)
            intervals = 1;
        return min + unit * intervals;
    }

    /// <summary>
    /// Formats a cached numeric value for a data label or tick. Honours simple percent
    /// format codes (<c>0%</c>, <c>0.0%</c>); everything else (incl. "General") uses a
    /// compact invariant decimal.
    /// </summary>
    public static string FormatValue(double value, string? formatCode)
    {
        if (!string.IsNullOrEmpty(formatCode) && formatCode.Contains('%'))
        {
            var percent = value * 100.0;
            var decimals = formatCode.Contains("0.0", StringComparison.Ordinal) ? 1 : 0;
            return percent.ToString(decimals == 1 ? "0.0" : "0", CultureInfo.InvariantCulture) + "%";
        }

        // "General"-style: integers without decimals, fractions trimmed.
        return value.ToString("0.########", CultureInfo.InvariantCulture);
    }
}
