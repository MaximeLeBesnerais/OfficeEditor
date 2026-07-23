using OfficeEditor.Api.Models;
using OfficeEditor.Api.Services;

namespace OfficeEditor.Api.Tests.Unit;

/// <summary>
/// Pins the shape and defaults of the API request/response models. These
/// records are the API's wire contract; accidental shape changes break clients.
/// </summary>
public sealed class ConversionModelTests
{
    [Fact]
    public void ConversionRequest_MapsPositionalPropertiesAndDefaultsOptionsToNull()
    {
        byte[] bytes = [1, 2, 3];

        var request = new ConversionRequest(bytes, "deck.pptx", ConversionTargetFormat.Pdf);

        Assert.Same(bytes, request.SourceBytes);
        Assert.Equal("deck.pptx", request.SourceFileName);
        Assert.Equal(ConversionTargetFormat.Pdf, request.TargetFormat);
        Assert.Null(request.Options);
    }

    [Fact]
    public void ConversionRequest_AcceptsExplicitOptions()
    {
        IReadOnlyDictionary<string, string> options = new Dictionary<string, string>
        {
            ["ppi"] = "144"
        };

        var request = new ConversionRequest([1], "deck.pptx", ConversionTargetFormat.Png, options);

        Assert.Same(options, request.Options);
        Assert.Equal("144", request.Options!["ppi"]);
    }

    [Fact]
    public void ConversionResult_MapsPositionalPropertiesAndDefaultsOptionalMembers()
    {
        var result = new ConversionResult(true, [9, 9], "out.pdf", "application/pdf");

        Assert.True(result.Success);
        Assert.Equal([9, 9], result.OutputBytes);
        Assert.Equal("out.pdf", result.OutputFileName);
        Assert.Equal("application/pdf", result.ContentType);
        Assert.Null(result.ErrorMessage);
        Assert.Null(result.Messages);
    }

    [Fact]
    public void ConversionResult_FailureShapeCarriesErrorAndMessages()
    {
        var result = new ConversionResult(
            Success: false,
            OutputBytes: null,
            OutputFileName: "",
            ContentType: "",
            ErrorMessage: "compile failed",
            Messages: ["line 3: unexpected token"]);

        Assert.False(result.Success);
        Assert.Null(result.OutputBytes);
        Assert.Equal("compile failed", result.ErrorMessage);
        Assert.Equal(["line 3: unexpected token"], result.Messages);
    }

    [Fact]
    public void ConvertResponse_SuccessDefaultsAllOptionalMembersToNull()
    {
        var response = new ConvertResponse(Success: true);

        Assert.True(response.Success);
        Assert.Null(response.OutputFileName);
        Assert.Null(response.ContentType);
        Assert.Null(response.DownloadUrl);
        Assert.Null(response.PreviewUrl);
        Assert.Null(response.ErrorMessage);
        Assert.Null(response.Messages);
    }

    [Fact]
    public void ConvertResponse_RecordEqualityHoldsForIdenticalValues()
    {
        var first = new ConvertResponse(true, "out.pdf", "application/pdf", "http://x/dl/1", null, null, null);
        var second = new ConvertResponse(true, "out.pdf", "application/pdf", "http://x/dl/1", null, null, null);

        Assert.Equal(first, second);
    }

    [Fact]
    public void ConvertRequest_RequiresTargetFormatAndDefaultsOptionalMembers()
    {
        var request = new ConvertRequest("pdf");

        Assert.Equal("pdf", request.TargetFormat);
        Assert.Null(request.SampleName);
        Assert.Null(request.Options);
    }

    [Fact]
    public void SampleFileDto_MapsPositionalProperties()
    {
        var dto = new SampleFileDto("northwind-demo.pptx", "pptx", "A sample deck.");

        Assert.Equal("northwind-demo.pptx", dto.Name);
        Assert.Equal("pptx", dto.Format);
        Assert.Equal("A sample deck.", dto.Description);
    }

    [Fact]
    public void StoredResult_MapsPositionalProperties()
    {
        byte[] bytes = [1];

        var stored = new StoredResult(bytes, "image/png", "page-001.png");

        Assert.Same(bytes, stored.Bytes);
        Assert.Equal("image/png", stored.ContentType);
        Assert.Equal("page-001.png", stored.FileName);
    }
}
