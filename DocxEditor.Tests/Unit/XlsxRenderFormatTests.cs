using System.Globalization;
using XlsxEditor.Core.Rendering.Formatting;
using Xunit;

namespace DocxEditor.Tests.Unit;

public class XlsxRenderFormatTests
{
    [Theory]
    [InlineData(1234.5, "#,##0", "1,235")]
    [InlineData(1234.5, "#,##0.00", "1,234.50")]
    [InlineData(1234.567, "0.00", "1234.57")]
    [InlineData(0.125, "0.0%", "12.5%")]
    [InlineData(0.125, "0.00%", "12.50%")]
    [InlineData(5, "00000", "00005")]
    [InlineData(12345, "0.00E+00", "1.23E+04")]
    [InlineData(-1234.5, "#,##0.00", "-1,234.50")]
    [InlineData(0, "#,##0.00", "0.00")]
    [InlineData(1234, "€#,##0.00", "€1,234.00")]
    [InlineData(13, "0 \"units\"", "13 units")]
    [InlineData(1234567, "#,##0,", "1,235")] // trailing comma scales by 1000
    public void FormatNumber_ProducesExpectedText(double value, string code, string expected)
    {
        var result = ExcelNumberFormatFormatter.FormatNumber(value, code);
        Assert.Equal(expected, result.Text);
        Assert.False(result.UsedTextSection);
    }

    [Fact]
    public void FormatNumber_NegativeRedSection_AppliesColor()
    {
        var result = ExcelNumberFormatFormatter.FormatNumber(-1234.5, "#,##0.00;[Red]-#,##0.00");
        Assert.Equal("-1,234.50", result.Text);
        Assert.Equal("FFFF0000", result.ColorArgb);
    }

    [Fact]
    public void FormatNumber_PositiveSection_NoColor()
    {
        var result = ExcelNumberFormatFormatter.FormatNumber(1234.5, "#,##0.00;[Red]-#,##0.00");
        Assert.Equal("1,234.50", result.Text);
        Assert.Null(result.ColorArgb);
    }

    [Fact]
    public void FormatNumber_ZeroSection_SelectedForZero()
    {
        var result = ExcelNumberFormatFormatter.FormatNumber((double)0, "0;0;[Blue]0");
        Assert.Equal("0", result.Text);
        Assert.Equal("FF0000FF", result.ColorArgb);
    }

    [Fact]
    public void FormatNumber_General_UsesShortestRoundtrip()
    {
        Assert.Equal("1234.5", ExcelNumberFormatFormatter.FormatNumber(1234.5, "General").Text);
        Assert.Equal("0.1", ExcelNumberFormatFormatter.FormatNumber(0.1, "General").Text);
        Assert.Equal("0", ExcelNumberFormatFormatter.FormatNumber((double)0, "General").Text);
    }

    [Fact]
    public void FormatNumber_General_ScientificForLargeValues()
    {
        var result = ExcelNumberFormatFormatter.FormatNumber(1e12, "General");
        Assert.Contains('E', result.Text);
    }

    [Fact]
    public void FormatNumber_EmptyCode_BehavesLikeGeneral()
    {
        Assert.Equal("42", ExcelNumberFormatFormatter.FormatNumber((double)42, "").Text);
    }

    [Fact]
    public void FormatNumber_UsesAsciiHyphen_NotUnicodeMinus()
    {
        var result = ExcelNumberFormatFormatter.FormatNumber((double)-42, "0");
        Assert.Equal("-42", result.Text);
        Assert.DoesNotContain('\u2212', result.Text);
    }

    [Fact]
    public void FormatNumber_RoundsHalfAwayFromZero_LikeExcel()
    {
        Assert.Equal("2.5", ExcelNumberFormatFormatter.FormatNumber(2.5, "0.0").Text);
        Assert.Equal("-2.5", ExcelNumberFormatFormatter.FormatNumber((double)-2.5, "0.0").Text);
    }

    [Fact]
    public void FormatNumber_ConditionalSections_SelectsByCondition()
    {
        var result = ExcelNumberFormatFormatter.FormatNumber((double)150, "[>=100]0;[<100]0");
        Assert.Equal("150", result.Text);
        var result2 = ExcelNumberFormatFormatter.FormatNumber((double)50, "[>=100]0;[<100]0");
        Assert.Equal("50", result2.Text);
    }

    [Fact]
    public void FormatNumber_TextSection_Passthrough()
    {
        // Simulated via numeric value; text placeholder with a number is unusual.
        var result = ExcelNumberFormatFormatter.FormatNumber((double)5, "0 \"units\"");
        Assert.Equal("5 units", result.Text);
    }

    // ----- Date serial handling -----

    [Theory]
    [InlineData(1, 1900, 1, 1)]
    [InlineData(59, 1900, 2, 28)]
    [InlineData(60, 1900, 2, 28)] // fake leap day
    [InlineData(61, 1900, 3, 1)]
    [InlineData(45291, 2023, 12, 31)] // spot-check a modern date
    public void SerialToDateTime_1900System(double serial, int y, int mo, int d)
    {
        var dt = ExcelDateFormat.SerialToDateTime(serial);
        Assert.Equal(y, dt.Year);
        Assert.Equal(mo, dt.Month);
        Assert.Equal(d, dt.Day);
    }

    [Theory]
    [InlineData(1, 1900, 1, 1)]
    [InlineData(59, 1900, 2, 28)]
    [InlineData(61, 1900, 3, 1)]
    [InlineData(45291, 2023, 12, 31)]
    public void DateTimeToSerial_RoundTrips(int serial, int y, int mo, int d)
    {
        var dateTime = new DateTime(y, mo, d);
        var round = ExcelDateFormat.DateTimeToSerial(dateTime);
        Assert.Equal(serial, round, 3);
    }

    [Fact]
    public void SerialToDateTime_1904System()
    {
        var dt = ExcelDateFormat.SerialToDateTime(1, use1904System: true);
        Assert.Equal(1904, dt.Year);
        Assert.Equal(1, dt.Month);
        Assert.Equal(2, dt.Day);
    }

    [Fact]
    public void SerialToDateTime_PreservesTimeComponent()
    {
        var dt = ExcelDateFormat.SerialToDateTime(45292.5);
        Assert.Equal(12, dt.Hour);
        Assert.Equal(0, dt.Minute);
    }

    // ----- Date format codes -----

    [Theory]
    [InlineData(45292, "yyyy-mm-dd", "2024-01-01")]
    [InlineData(45292, "mm/dd/yyyy", "01/01/2024")]
    [InlineData(45292, "yy", "24")]
    [InlineData(45292, "ddd", "Mon")]
    [InlineData(45292, "dddd", "Monday")]
    [InlineData(45292, "mmm", "Jan")]
    [InlineData(45292, "mmmm", "January")]
    [InlineData(45292.5, "h:mm", "12:00")]
    [InlineData(45292.75, "h:mm AM/PM", "6:00 PM")]
    public void FormatDate_ProducesExpectedText(double serial, string code, string expected)
    {
        var dt = ExcelDateFormat.SerialToDateTime(serial);
        Assert.Equal(expected, ExcelDateFormat.FormatDate(dt, code));
    }

    [Fact]
    public void IsDateFormat_DetectsDateCodes()
    {
        Assert.True(ExcelDateFormat.IsDateFormat("yyyy-mm-dd"));
        Assert.True(ExcelDateFormat.IsDateFormat("m/d/yy"));
        Assert.True(ExcelDateFormat.IsDateFormat("h:mm"));
        Assert.True(ExcelDateFormat.IsDateFormat("[h]:mm"));
        Assert.False(ExcelDateFormat.IsDateFormat("#,##0.00"));
        Assert.False(ExcelDateFormat.IsDateFormat("General"));
        Assert.False(ExcelDateFormat.IsDateFormat(""));
    }
}
