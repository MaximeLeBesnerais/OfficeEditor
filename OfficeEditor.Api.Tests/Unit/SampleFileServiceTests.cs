using OfficeEditor.Api.Services;
using OfficeEditor.Core.Services;

namespace OfficeEditor.Api.Tests.Unit;

public sealed class SampleFileServiceTests
{
    private readonly SampleFileService _service = new();

    [Fact]
    public void GetSampleFiles_ReturnsExpectedCatalog()
    {
        var samples = _service.GetSampleFiles();

        Assert.Equal(5, samples.Count);
        Assert.Equal(
            new[]
            {
                "sales_acceleration_deck.pptx",
                "northwind-demo.pptx",
                "annual-report.docx",
                "monitoring-report.docx",
                "sample.md"
            },
            samples.Select(s => s.Name));
    }

    [Fact]
    public void GetSampleFiles_EntriesAreWellFormed()
    {
        var samples = _service.GetSampleFiles();

        Assert.All(samples, sample =>
        {
            Assert.False(string.IsNullOrWhiteSpace(sample.Name));
            Assert.False(string.IsNullOrWhiteSpace(sample.Path));
            Assert.False(string.IsNullOrWhiteSpace(sample.Description));
            Assert.True(Enum.IsDefined(sample.Format));
        });
        Assert.Equal(samples.Count, samples.Select(s => s.Name).Distinct(StringComparer.Ordinal).Count());
    }

    [Theory]
    [InlineData("sales_acceleration_deck.pptx", OfficeDocumentFormat.Pptx)]
    [InlineData("northwind-demo.pptx", OfficeDocumentFormat.Pptx)]
    [InlineData("annual-report.docx", OfficeDocumentFormat.Docx)]
    [InlineData("monitoring-report.docx", OfficeDocumentFormat.Docx)]
    [InlineData("sample.md", OfficeDocumentFormat.Markdown)]
    public void GetSampleFiles_DeclaresExpectedFormats(string name, OfficeDocumentFormat expectedFormat)
    {
        var sample = Assert.Single(_service.GetSampleFiles(), s => s.Name == name);
        Assert.Equal(expectedFormat, sample.Format);
    }

    [Fact]
    public async Task LoadAsync_UnknownName_ThrowsFileNotFoundException()
    {
        var exception = await Assert.ThrowsAsync<FileNotFoundException>(
            () => _service.LoadAsync("does-not-exist.pptx"));

        Assert.Contains("does-not-exist.pptx", exception.Message);
    }

    [Fact]
    public async Task LoadAsync_PptxSample_ReturnsZipPayload()
    {
        var bytes = await _service.LoadAsync("northwind-demo.pptx");

        Assert.NotEmpty(bytes);
        // PPTX is a ZIP container: expect the local-file-header signature.
        Assert.True(bytes.Length > 4);
        Assert.Equal((byte)'P', bytes[0]);
        Assert.Equal((byte)'K', bytes[1]);
    }

    [Fact]
    public async Task LoadAsync_MarkdownSample_ReturnsNonEmptyPayload()
    {
        var bytes = await _service.LoadAsync("sample.md");

        Assert.NotEmpty(bytes);
    }

    [Fact]
    public async Task LoadAsync_NameMatchingIsCaseSensitive()
    {
        // Names are matched with the default (ordinal, case-sensitive) comparer;
        // a differently-cased name must not silently resolve to a sample.
        await Assert.ThrowsAsync<FileNotFoundException>(
            () => _service.LoadAsync("Northwind-Demo.pptx"));
    }
}
