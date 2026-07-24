using PptxEditor.Core.Converters.Charts;
using Xunit;

namespace DocxEditor.Tests.Unit.Charts;

/// <summary>
/// Auto value-axis math: nice-rounded units, axis max extended to fit data labels
/// (PowerPoint behaviour observed on both REF decks), negative ranges, explicit
/// overrides, and value formatting.
/// </summary>
public sealed class ChartAxisScaleTests
{
    [Theory]
    // Both REF decks: labels shown → axis extends past the data max.
    [InlineData(0, 100, true, 0, 120, 20)]   // sales_acceleration_deck slide 9
    [InlineData(0, 48, true, 0, 60, 10)]     // AetherLink slide 9
    // Without labels the axis stops at the nice ceiling of the data max.
    [InlineData(0, 100, false, 0, 100, 20)]
    [InlineData(0, 68, false, 0, 80, 20)]
    [InlineData(0, 7, false, 0, 8, 2)]
    // Negative data extends the axis below zero to a unit multiple.
    [InlineData(-35, 50, false, -40, 60, 20)]
    [InlineData(-7.5, 30, false, -10, 30, 10)]
    public void Compute_AutoAxis_NiceRoundsLikePowerPoint(
        double dataMin, double dataMax, bool labels,
        double expectedMin, double expectedMax, double expectedUnit)
    {
        var (min, max, unit) = ChartAxisScale.Compute(dataMin, dataMax, labels);

        Assert.Equal(expectedMin, min);
        Assert.Equal(expectedMax, max);
        Assert.Equal(expectedUnit, unit);
    }

    [Fact]
    public void Compute_ExplicitBoundsAndUnit_Win()
    {
        var (min, max, unit) = ChartAxisScale.Compute(0, 100, true, explicitMin: -50, explicitMax: 200, explicitUnit: 50);

        Assert.Equal(-50, min);
        Assert.Equal(200, max);
        Assert.Equal(50, unit);
    }

    [Fact]
    public void Compute_AllZeroData_StillProducesAUsableAxis()
    {
        var (min, max, unit) = ChartAxisScale.Compute(0, 0, false);

        Assert.Equal(0, min);
        Assert.True(max > min);
        Assert.True(unit > 0);
    }

    [Fact]
    public void Ticks_SpanMinToMaxInclusive()
    {
        Assert.Equal(new double[] { 0, 20, 40, 60, 80, 100, 120 }, ChartAxisScale.Ticks(0, 120, 20));
        Assert.Equal(new double[] { -40, -20, 0, 20, 40, 60 }, ChartAxisScale.Ticks(-40, 60, 20));
    }

    [Theory]
    [InlineData(9.6, 10)]
    [InlineData(10, 10)]
    [InlineData(13.6, 20)]
    [InlineData(20, 20)]
    [InlineData(21, 25)]
    [InlineData(0.13, 0.2)]
    [InlineData(950, 1000)]
    public void NiceCeiling_PicksHumanSteps(double value, double expected)
        => Assert.Equal(expected, ChartAxisScale.NiceCeiling(value));

    [Theory]
    [InlineData(100, "General", "100")]
    [InlineData(68.25, "General", "68.25")]
    [InlineData(0.42, "0%", "42%")]
    [InlineData(0.426, "0.0%", "42.6%")]
    public void FormatValue_HonoursGeneralAndPercentFormats(double value, string format, string expected)
        => Assert.Equal(expected, ChartAxisScale.FormatValue(value, format));
}
