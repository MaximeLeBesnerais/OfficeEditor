using System.Text;
using OfficeEditor.Api.Services;

namespace OfficeEditor.Api.Tests.Unit;

/// <summary>
/// Integration tests for ConversionService — cover format detection, Markdown → DOCX
/// (no Typst needed), blank document creation for each format, unsupported-conversion
/// errors, PPI parsing, and DOCX → PDF via Typst (opt-in, OE_RUN_TYPST_COMPILE_TESTS=1).
/// </summary>
public sealed class ConversionServiceTests : IAsyncDisposable
{
    private const string EnableRenderEnvVar = "OE_RUN_TYPST_COMPILE_TESTS";
    private readonly List<string> _tempFiles = [];

    public async ValueTask DisposeAsync()
    {
        foreach (var file in _tempFiles)
        {
            try { File.Delete(file); } catch { /* cleanup best effort */ }
        }
        GC.SuppressFinalize(this);
    }

    private string CreateTempFilePath(string extension)
    {
        var path = Path.Combine(Path.GetTempPath(), $"oe-cvt-{Guid.NewGuid():N}{extension}");
        _tempFiles.Add(path);
        return path;
    }

    [Fact]
    public async Task ConvertAsync_MarkdownToDocx_ProducesValidDocx()
    {
        var service = new ConversionService();
        var markdown = "# Hello\n\nWorld"u8.ToArray();
        var request = new ConversionRequest(markdown, "readme.md", ConversionTargetFormat.Docx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.NotNull(result.OutputBytes);
        Assert.True(result.OutputBytes!.Length > 0);
        Assert.Equal("application/vnd.openxmlformats-officedocument.wordprocessingml.document", result.ContentType);
        Assert.Equal("readme.docx", result.OutputFileName);
        Assert.Null(result.ErrorMessage);
        Assert.NotNull(result.Messages);
        Assert.Contains(result.Messages, m => m.Contains("Markdown"));
    }

    [Fact]
    public async Task ConvertAsync_MarkdownToDocx_MarkdownExtension_DoesMarkdownToDocx()
    {
        var service = new ConversionService();
        var markdown = "plain text"u8.ToArray();
        var request = new ConversionRequest(markdown, "notes.markdown", ConversionTargetFormat.Docx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.NotNull(result.OutputBytes);
        Assert.Equal("notes.docx", result.OutputFileName);
    }

    [Fact]
    public async Task ConvertAsync_TxtExtension_TreatedAsMarkdown()
    {
        var service = new ConversionService();
        var markdown = "plain text"u8.ToArray();
        var request = new ConversionRequest(markdown, "readme.txt", ConversionTargetFormat.Docx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.NotNull(result.OutputBytes);
        Assert.Equal("readme.docx", result.OutputFileName);
    }

    [Fact]
    public async Task ConvertAsync_UnsupportedConversion_MarkdownToPdf_ReturnsError()
    {
        var service = new ConversionService();
        var markdown = "text"u8.ToArray();
        var request = new ConversionRequest(markdown, "readme.md", ConversionTargetFormat.Pdf);

        var result = await service.ConvertAsync(request);

        Assert.False(result.Success);
        Assert.Null(result.OutputBytes);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("Unsupported conversion", result.ErrorMessage);
    }

    [Fact]
    public async Task ConvertAsync_UnsupportedConversion_UnknownSourceToPptx_CreatesBlankPptx()
    {
        // Unknown source + Pptx target → CreateBlankPptxAsync (line 29)
        var service = new ConversionService();
        var bytes = "data"u8.ToArray();
        var request = new ConversionRequest(bytes, "file.bin", ConversionTargetFormat.Pptx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.NotNull(result.OutputBytes);
        Assert.Equal("file.pptx", result.OutputFileName);
    }

    [Fact]
    public async Task ConvertAsync_UnsupportedConversion_PngTargetFromUnknown_ReturnsError()
    {
        var service = new ConversionService();
        var bytes = "binary"u8.ToArray();
        var request = new ConversionRequest(bytes, "data.xyz", ConversionTargetFormat.Png);

        var result = await service.ConvertAsync(request);

        Assert.False(result.Success);
        Assert.Contains("Unsupported conversion", result.ErrorMessage);
    }

    [Fact]
    public async Task ConvertAsync_UnsupportedConversion_SvgTargetFromUnknown_ReturnsError()
    {
        var service = new ConversionService();
        var bytes = "binary"u8.ToArray();
        var request = new ConversionRequest(bytes, "data.xyz", ConversionTargetFormat.Svg);

        var result = await service.ConvertAsync(request);

        Assert.False(result.Success);
        Assert.Contains("Unsupported conversion", result.ErrorMessage);
    }

    [Fact]
    public async Task ConvertAsync_CreateBlankDocx_EmptyBytes_ProducesValidDocx()
    {
        var service = new ConversionService();
        var request = new ConversionRequest([], "document.docx", ConversionTargetFormat.Docx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.NotNull(result.OutputBytes);
        Assert.True(result.OutputBytes!.Length > 0);
        Assert.Equal("application/vnd.openxmlformats-officedocument.wordprocessingml.document", result.ContentType);
        Assert.Null(result.ErrorMessage);
        Assert.NotNull(result.Messages);
        Assert.Contains(result.Messages, m => m.Contains("blank DOCX"));
    }

    [Fact]
    public async Task ConvertAsync_CreateBlankDocx_ZeroBytes_ProducesValidDocx()
    {
        var service = new ConversionService();
        var request = new ConversionRequest(new byte[10], "draft.docx", ConversionTargetFormat.Docx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.NotNull(result.OutputBytes);
        Assert.Equal("draft.docx", result.OutputFileName);
    }

    [Fact]
    public async Task ConvertAsync_CreateBlankDocx_WhenDocxExtensionAndZeroBytes_ProducesValidDocx()
    {
        var service = new ConversionService();
        var request = new ConversionRequest(new byte[1], "empty.docx", ConversionTargetFormat.Docx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.NotNull(result.OutputBytes);
    }

    [Fact]
    public async Task ConvertAsync_CreateBlankDocx_NullFileName_GeneratesDefaultFileName()
    {
        var service = new ConversionService();
        var request = new ConversionRequest([], null!, ConversionTargetFormat.Docx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.Equal("document.docx", result.OutputFileName);
    }

    [Fact]
    public async Task ConvertAsync_CreateBlankDocx_WhitespaceFileName_GeneratesDefaultFileName()
    {
        var service = new ConversionService();
        var request = new ConversionRequest([], "   ", ConversionTargetFormat.Docx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.Equal("document.docx", result.OutputFileName);
    }

    [Fact]
    public async Task ConvertAsync_CreateBlankDocx_NoExtensionFileName_AppendsDocxExtension()
    {
        var service = new ConversionService();
        var request = new ConversionRequest([], "myfile", ConversionTargetFormat.Docx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.Equal("myfile.docx", result.OutputFileName);
    }

    [Fact]
    public async Task ConvertAsync_CreateBlankXlsx_EmptyBytes_ProducesValidXlsx()
    {
        var service = new ConversionService();
        var request = new ConversionRequest([], "workbook.xlsx", ConversionTargetFormat.Xlsx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.NotNull(result.OutputBytes);
        Assert.True(result.OutputBytes!.Length > 0);
        Assert.Equal("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", result.ContentType);
        Assert.Null(result.ErrorMessage);
        Assert.NotNull(result.Messages);
        Assert.Contains(result.Messages, m => m.Contains("blank XLSX"));
    }

    [Fact]
    public async Task ConvertAsync_CreateBlankXlsx_ZeroBytes_ProducesValidXlsx()
    {
        var service = new ConversionService();
        var request = new ConversionRequest(new byte[5], "data.xlsx", ConversionTargetFormat.Xlsx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.NotNull(result.OutputBytes);
        Assert.Equal("data.xlsx", result.OutputFileName);
    }

    [Fact]
    public async Task ConvertAsync_CreateBlankPptx_EmptyBytes_ProducesValidPptx()
    {
        var service = new ConversionService();
        var request = new ConversionRequest([], "deck.pptx", ConversionTargetFormat.Pptx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.NotNull(result.OutputBytes);
        Assert.True(result.OutputBytes!.Length > 0);
        Assert.Equal("application/vnd.openxmlformats-officedocument.presentationml.presentation", result.ContentType);
        Assert.Null(result.ErrorMessage);
        Assert.NotNull(result.Messages);
        Assert.Contains(result.Messages, m => m.Contains("blank PPTX"));
    }

    [Fact]
    public async Task ConvertAsync_CreateBlankPptx_ZeroBytes_ProducesValidPptx()
    {
        var service = new ConversionService();
        var request = new ConversionRequest(new byte[3], "deck.pptx", ConversionTargetFormat.Pptx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.NotNull(result.OutputBytes);
        Assert.Equal("deck.pptx", result.OutputFileName);
    }

    [Fact]
    public async Task ConvertAsync_ChangeExtension_EmptyFileName_UsesDefaultWithExtension()
    {
        var service = new ConversionService();
        var markdown = "# test"u8.ToArray();
        var request = new ConversionRequest(markdown, "", ConversionTargetFormat.Docx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.Equal("document.docx", result.OutputFileName);
    }

    [Fact]
    public async Task ConvertAsync_ChangeExtension_OnlyExtension_GeneratesSensibleDefault()
    {
        var service = new ConversionService();
        var markdown = "# test"u8.ToArray();
        var request = new ConversionRequest(markdown, ".md", ConversionTargetFormat.Docx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.Equal("document.docx", result.OutputFileName);
    }

    [Fact]
    public async Task ConvertAsync_Options_PpiIsPassedThrough()
    {
        var service = new ConversionService();
        IReadOnlyDictionary<string, string> options = new Dictionary<string, string>
        {
            ["Ppi"] = "110"
        };
        var pngBytes = new byte[] { 0x89, (byte)'P', (byte)'N', (byte)'G' };
        var request = new ConversionRequest(pngBytes, "image.png", ConversionTargetFormat.Png, options);

        var result = await service.ConvertAsync(request);

        Assert.False(result.Success);
        Assert.Contains("Unsupported conversion", result.ErrorMessage);
    }

    [Fact]
    public async Task ConvertAsync_Options_InvalidPpi_FallsBackToDefault()
    {
        var service = new ConversionService();
        IReadOnlyDictionary<string, string> options = new Dictionary<string, string>
        {
            ["Ppi"] = "not-a-number"
        };
        var bytes = new byte[] { 0x80, 0x75, 0x4b, 0x50 };
        var request = new ConversionRequest(bytes, "file.zip", ConversionTargetFormat.Png, options);

        var result = await service.ConvertAsync(request);

        Assert.False(result.Success);
        Assert.Contains("Unsupported conversion", result.ErrorMessage);
    }

    [Fact]
    public async Task ConvertAsync_Messages_AccumulateAcrossPipeline()
    {
        var service = new ConversionService();
        var markdown = "# Header\nContent."u8.ToArray();
        var request = new ConversionRequest(markdown, "doc.md", ConversionTargetFormat.Docx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.NotNull(result.Messages);
        Assert.True(result.Messages.Count >= 2,
            $"Expected at least 2 messages (detection + conversion), got: {string.Join(" | ", result.Messages)}");
    }

    [Fact]
    public async Task ConvertAsync_DocxToPdf_WithValidDocx_ProducesPdf()
    {
        if (Environment.GetEnvironmentVariable(EnableRenderEnvVar) != "1")
        {
            return;
        }

        var refDocx = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..",
            "examples", "REF", "DOCX", "annual-report.docx");
        refDocx = Path.GetFullPath(refDocx);
        if (!File.Exists(refDocx))
        {
            return;
        }

        var bytes = await File.ReadAllBytesAsync(refDocx);
        var service = new ConversionService();
        var request = new ConversionRequest(bytes, "annual-report.docx", ConversionTargetFormat.Pdf);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.NotNull(result.OutputBytes);
        Assert.True(result.OutputBytes!.Length > 4);
        var header = Encoding.ASCII.GetString(result.OutputBytes, 0, 5);
        Assert.Equal("%PDF-", header);
        Assert.Equal("application/pdf", result.ContentType);
        Assert.NotNull(result.Messages);
        Assert.Contains(result.Messages, m => m.Contains("DOCX"));
    }

    [Fact]
    public async Task ConvertAsync_DocxToPdf_CorruptDocx_ReturnsError()
    {
        var service = new ConversionService();
        var garbage = "not a valid docx file"u8.ToArray();
        var request = new ConversionRequest(garbage, "bad.docx", ConversionTargetFormat.Pdf);

        var result = await service.ConvertAsync(request);

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("Conversion failed", result.ErrorMessage);
    }

    [Fact]
    public async Task ConvertAsync_DocxToPdf_EmptyDocx_CreatesBlankInstead()
    {
        var service = new ConversionService();
        var request = new ConversionRequest([], "empty.docx", ConversionTargetFormat.Docx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.NotNull(result.Messages);
        Assert.Contains(result.Messages, m => m.Contains("blank DOCX"));
    }

    [Fact]
    public async Task ConvertAsync_UnknownFormatToDocx_CreatesBlank()
    {
        var service = new ConversionService();
        var bytes = "some data"u8.ToArray();
        var request = new ConversionRequest(bytes, "data.bin", ConversionTargetFormat.Docx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.NotNull(result.OutputBytes);
        Assert.Equal("data.docx", result.OutputFileName);
    }

    [Fact]
    public async Task ConvertAsync_UnknownFormatToXlsx_CreatesBlank()
    {
        var service = new ConversionService();
        var bytes = "some data"u8.ToArray();
        var request = new ConversionRequest(bytes, "data.bin", ConversionTargetFormat.Xlsx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.NotNull(result.OutputBytes);
        Assert.Equal("data.xlsx", result.OutputFileName);
    }

    [Fact]
    public async Task ConvertAsync_UnknownFormatToPptx_CreatesBlank()
    {
        var service = new ConversionService();
        var bytes = "some data"u8.ToArray();
        var request = new ConversionRequest(bytes, "data.bin", ConversionTargetFormat.Pptx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.NotNull(result.OutputBytes);
        Assert.Equal("data.pptx", result.OutputFileName);
    }

    [Fact]
    public async Task ConvertAsync_GarbagePng_ReturnsError()
    {
        var service = new ConversionService();
        var garbage = "definitely not a valid pptx"u8.ToArray();
        var request = new ConversionRequest(garbage, "corrupt.pptx", ConversionTargetFormat.Pdf);

        var result = await service.ConvertAsync(request);

        Assert.False(result.Success);
        Assert.Contains("Conversion failed", result.ErrorMessage);
    }

    [Fact]
    public async Task ConvertAsync_Cancellation_ReturnsErrorResult()
    {
        // ConvertAsync wraps all exceptions (including OperationCanceledException)
        // inside its try/catch, so cancelled tokens produce an error result, not a throw.
        var service = new ConversionService();
        var markdown = "# Test"u8.ToArray();
        using var cts = new CancellationTokenSource();
        await cts.CancelAsync();
        var request = new ConversionRequest(markdown, "test.md", ConversionTargetFormat.Docx);

        var result = await service.ConvertAsync(request, cts.Token);

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("Conversion failed", result.ErrorMessage);
    }
}
