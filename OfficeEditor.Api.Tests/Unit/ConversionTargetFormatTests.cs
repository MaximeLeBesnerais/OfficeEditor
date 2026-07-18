using OfficeEditor.Api.Services;

namespace OfficeEditor.Api.Tests.Unit;

/// <summary>
/// Tests the target-format parsing contract used by the /api/convert endpoint
/// (case-insensitive parse + <see cref="Enum.IsDefined"/> guard). The endpoint's
/// helper is private, so these tests pin the semantics on the enum itself.
/// </summary>
public sealed class ConversionTargetFormatTests
{
    private static bool TryParseTargetFormat(string? value, out ConversionTargetFormat targetFormat)
    {
        targetFormat = default;
        return !string.IsNullOrWhiteSpace(value)
            && Enum.TryParse(value, ignoreCase: true, out targetFormat)
            && Enum.IsDefined(typeof(ConversionTargetFormat), targetFormat);
    }

    [Theory]
    [InlineData("pdf", ConversionTargetFormat.Pdf)]
    [InlineData("PDF", ConversionTargetFormat.Pdf)]
    [InlineData("Pdf", ConversionTargetFormat.Pdf)]
    [InlineData("png", ConversionTargetFormat.Png)]
    [InlineData("PNG", ConversionTargetFormat.Png)]
    [InlineData("svg", ConversionTargetFormat.Svg)]
    [InlineData("Svg", ConversionTargetFormat.Svg)]
    [InlineData("docx", ConversionTargetFormat.Docx)]
    [InlineData("DOCX", ConversionTargetFormat.Docx)]
    [InlineData("pptx", ConversionTargetFormat.Pptx)]
    [InlineData("xlsx", ConversionTargetFormat.Xlsx)]
    public void Parse_SupportedFormats_AreCaseInsensitive(string value, ConversionTargetFormat expected)
    {
        var parsed = TryParseTargetFormat(value, out var format);

        Assert.True(parsed);
        Assert.Equal(expected, format);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("gif")]
    [InlineData("jpeg")]
    [InlineData("pdfx")]
    [InlineData("pd")]
    [InlineData("99")]
    [InlineData("-1")]
    public void Parse_UnsupportedOrMalformedValues_AreRejected(string? value)
    {
        // Only the boolean result is contractual: for numeric strings
        // Enum.TryParse leaves the parsed number in the out parameter even
        // though IsDefined rejects it, so no assertion on the out value here.
        var parsed = TryParseTargetFormat(value, out _);

        Assert.False(parsed);
    }

    [Fact]
    public void Parse_NumericStringForDefinedValue_IsAccepted()
    {
        // Documented quirk of Enum.TryParse + IsDefined: a numeric string that
        // maps onto a defined member parses successfully. Pin the behavior so a
        // future tightening of the endpoint validation is a deliberate change.
        var parsed = TryParseTargetFormat("0", out var format);

        Assert.True(parsed);
        Assert.Equal(ConversionTargetFormat.Pdf, format);
    }

    [Fact]
    public void Enum_ContainsExactlyTheDocumentedFormats()
    {
        var names = Enum.GetNames<ConversionTargetFormat>();

        Assert.Equal(new[] { "Pdf", "Png", "Svg", "Docx", "Pptx", "Xlsx" }, names);
    }
}
