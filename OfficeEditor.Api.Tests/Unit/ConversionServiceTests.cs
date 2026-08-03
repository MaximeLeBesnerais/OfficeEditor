using System.Text;
using OfficeEditor.Api.Services;

namespace OfficeEditor.Api.Tests.Unit;

/// <summary>
/// Integration tests for ConversionService — cover format detection, Markdown → DOCX
/// (no Typst needed), blank document creation for each format, unsupported-conversion
/// errors, PPI parsing, and DOCX → PDF via Typst (opt-in, OE_RUN_TYPST_COMPILE_TESTS=1).
/// </summary>
public sealed class ConversionServiceTests
{
    private const string EnableRenderEnvVar = "OE_RUN_TYPST_COMPILE_TESTS";

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
    public async Task ConvertAsync_UnknownSourceToPptx_CreatesBlankPptx()
    {
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
    public async Task ConvertAsync_Options_UnsupportedSrcDst_Rejected()
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
    public async Task ConvertAsync_Options_InvalidPpi_UnsupportedSrcDst_Rejected()
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

    [Theory]
    [InlineData("1", 36)]
    [InlineData("110", 110)]
    [InlineData("50000", 600)]
    [InlineData("NaN", 150)]
    [InlineData("Infinity", 150)]
    [InlineData("not-a-number", 150)]
    public void GetPpi_ClampsUntrustedRenderCost(string value, float expected)
    {
        IReadOnlyDictionary<string, string> options = new Dictionary<string, string>
        {
            ["Ppi"] = value
        };

        Assert.Equal(expected, ConversionService.GetPpi(options));
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
    public async Task ConvertAsync_GarbagePptx_ToPdf_ReturnsError()
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

    [Fact]
    public async Task ConvertAsync_GarbagePptx_ToPng_ReturnsError()
    {
        var service = new ConversionService();
        var garbage = "this is not a pptx"u8.ToArray();
        var request = new ConversionRequest(garbage, "corrupt.pptx", ConversionTargetFormat.Png);

        var result = await service.ConvertAsync(request);

        Assert.False(result.Success);
        Assert.Null(result.OutputBytes);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("Conversion failed", result.ErrorMessage);
        Assert.NotNull(result.Messages);
        Assert.Contains(result.Messages, m => m.Contains("Detected source format: Pptx"));
    }

    [Fact]
    public async Task ConvertAsync_GarbagePptx_ToSvg_ReturnsError()
    {
        var service = new ConversionService();
        var garbage = "this is not a pptx"u8.ToArray();
        var request = new ConversionRequest(garbage, "corrupt.pptx", ConversionTargetFormat.Svg);

        var result = await service.ConvertAsync(request);

        Assert.False(result.Success);
        Assert.Null(result.OutputBytes);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("Conversion failed", result.ErrorMessage);
        Assert.NotNull(result.Messages);
        Assert.Contains(result.Messages, m => m.Contains("Detected source format: Pptx"));
    }

    [Fact]
    public async Task ConvertAsync_EmptyPptxBytes_ToPng_ReturnsError()
    {
        var service = new ConversionService();
        var request = new ConversionRequest([], "empty.pptx", ConversionTargetFormat.Png);

        var result = await service.ConvertAsync(request);

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("Conversion failed", result.ErrorMessage);
    }

    [Fact]
    public async Task ConvertAsync_EmptyPptxBytes_ToSvg_ReturnsError()
    {
        var service = new ConversionService();
        var request = new ConversionRequest([], "empty.pptx", ConversionTargetFormat.Svg);

        var result = await service.ConvertAsync(request);

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("Conversion failed", result.ErrorMessage);
    }

    [Fact]
    public async Task ConvertAsync_Markdown_ToXlsx_ReturnsUnsupportedError()
    {
        var service = new ConversionService();
        var markdown = "# test"u8.ToArray();
        var request = new ConversionRequest(markdown, "notes.md", ConversionTargetFormat.Xlsx);

        var result = await service.ConvertAsync(request);

        Assert.False(result.Success);
        Assert.Null(result.OutputBytes);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("Unsupported conversion", result.ErrorMessage);
        Assert.Contains("Markdown", result.ErrorMessage);
        Assert.Contains("Xlsx", result.ErrorMessage);
    }

    [Fact]
    public async Task ConvertAsync_Markdown_ToPptx_ReturnsUnsupportedError()
    {
        var service = new ConversionService();
        var markdown = "# test"u8.ToArray();
        var request = new ConversionRequest(markdown, "notes.md", ConversionTargetFormat.Pptx);

        var result = await service.ConvertAsync(request);

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("Unsupported conversion", result.ErrorMessage);
        Assert.Contains("Markdown", result.ErrorMessage);
        Assert.Contains("Pptx", result.ErrorMessage);
    }

    [Fact]
    public async Task ConvertAsync_Markdown_ToSvg_ReturnsUnsupportedError()
    {
        var service = new ConversionService();
        var markdown = "# test"u8.ToArray();
        var request = new ConversionRequest(markdown, "notes.md", ConversionTargetFormat.Svg);

        var result = await service.ConvertAsync(request);

        Assert.False(result.Success);
        Assert.Contains("Unsupported conversion", result.ErrorMessage);
    }

    [Fact]
    public async Task ConvertAsync_CreateBlankXlsx_NullFileName_GeneratesDefaultFileName()
    {
        var service = new ConversionService();
        var request = new ConversionRequest([], null!, ConversionTargetFormat.Xlsx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.Equal("workbook.xlsx", result.OutputFileName);
    }

    [Fact]
    public async Task ConvertAsync_CreateBlankXlsx_WhitespaceFileName_GeneratesDefaultFileName()
    {
        var service = new ConversionService();
        var request = new ConversionRequest([], "   ", ConversionTargetFormat.Xlsx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.Equal("workbook.xlsx", result.OutputFileName);
    }

    [Fact]
    public async Task ConvertAsync_CreateBlankXlsx_NoExtensionFileName_AppendsXlsxExtension()
    {
        var service = new ConversionService();
        var request = new ConversionRequest([], "mydata", ConversionTargetFormat.Xlsx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.Equal("mydata.xlsx", result.OutputFileName);
    }

    [Fact]
    public async Task ConvertAsync_CreateBlankPptx_NullFileName_GeneratesDefaultFileName()
    {
        var service = new ConversionService();
        var request = new ConversionRequest([], null!, ConversionTargetFormat.Pptx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.Equal("presentation.pptx", result.OutputFileName);
    }

    [Fact]
    public async Task ConvertAsync_CreateBlankPptx_WhitespaceFileName_GeneratesDefaultFileName()
    {
        var service = new ConversionService();
        var request = new ConversionRequest([], "   ", ConversionTargetFormat.Pptx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.Equal("presentation.pptx", result.OutputFileName);
    }

    [Fact]
    public async Task ConvertAsync_CreateBlankPptx_NoExtensionFileName_AppendsPptxExtension()
    {
        var service = new ConversionService();
        var request = new ConversionRequest([], "mydeck", ConversionTargetFormat.Pptx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.Equal("mydeck.pptx", result.OutputFileName);
    }

    [Fact]
    public async Task ConvertAsync_DocxExtensionToXlsx_ReturnsUnsupportedError()
    {
        // A .docx source targeting XLSX goes to the unsupported default arm, NOT blank creation.
        var service = new ConversionService();
        var bytes = "some bytes"u8.ToArray();
        var request = new ConversionRequest(bytes, "report.docx", ConversionTargetFormat.Xlsx);

        var result = await service.ConvertAsync(request);

        Assert.False(result.Success);
        Assert.Contains("Unsupported conversion", result.ErrorMessage);
    }

    [Fact]
    public async Task ConvertAsync_XlsxExtensionToDocx_ReturnsUnsupportedError()
    {
        var service = new ConversionService();
        var bytes = "some bytes"u8.ToArray();
        var request = new ConversionRequest(bytes, "data.xlsx", ConversionTargetFormat.Docx);

        var result = await service.ConvertAsync(request);

        Assert.False(result.Success);
        Assert.Contains("Unsupported conversion", result.ErrorMessage);
    }

    [Fact]
    public async Task ConvertAsync_ChangeExtension_PathSeparatorsInFileName_KeepsOnlyBaseName()
    {
        var service = new ConversionService();
        var markdown = "# test"u8.ToArray();
        var request = new ConversionRequest(markdown, "subdir/file.md", ConversionTargetFormat.Docx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.Equal("file.docx", result.OutputFileName);
    }

    [Fact]
    public async Task ConvertAsync_ChangeExtension_OnlyDotFile_UsesDefaultExtension()
    {
        var service = new ConversionService();
        // Path.GetFileNameWithoutExtension(".gitignore") returns ".gitignore"
        // because the leading dot makes it "just an extension" — but in practice
        // the result depends on the platform. Verify behavioral contract.
        var markdown = "# test"u8.ToArray();
        var request = new ConversionRequest(markdown, ".gitignore", ConversionTargetFormat.Docx);

        var result = await service.ConvertAsync(request);

        // .gitignore has no base name → ChangeExtension returns "document.xxx"
        Assert.True(result.Success);
        Assert.Equal("document.docx", result.OutputFileName);
    }

    [Fact]
    public async Task ConvertAsync_NullSourceFileName_MarkdownToDocx_DetectsUnknownAndCreatesBlank()
    {
        // When the source filename is null, DetectSourceFormat → Unknown,
        // and since Unknown → Docx creates a blank document (not markdown conversion).
        var service = new ConversionService();
        var bytes = "some data"u8.ToArray();
        var request = new ConversionRequest(bytes, null!, ConversionTargetFormat.Docx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.Equal("document.docx", result.OutputFileName);
        Assert.NotNull(result.Messages);
        Assert.Contains(result.Messages, m => m.Contains("Detected source format: Unknown"));
    }

    [Fact]
    public async Task ConvertAsync_JsonToDocx_ProducesValidDocx()
    {
        var service = new ConversionService();
        var json = """{ "version": "1.0", "sections": [ { "blocks": [ { "type": "paragraph", "text": "Hello JSON DOCX" } ] } ] }"""u8.ToArray();
        var request = new ConversionRequest(json, "report.json", ConversionTargetFormat.Docx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.NotNull(result.OutputBytes);
        Assert.True(result.OutputBytes!.Length > 0);
        Assert.Equal("application/vnd.openxmlformats-officedocument.wordprocessingml.document", result.ContentType);
        Assert.Equal("report.docx", result.OutputFileName);
        Assert.Null(result.ErrorMessage);
        Assert.NotNull(result.Messages);
        Assert.Contains(result.Messages, m => m.Contains("JSON"));
    }

    [Fact]
    public async Task ConvertAsync_JsonToXlsx_ProducesValidXlsx()
    {
        var service = new ConversionService();
        var json = """{ "version": "1.0", "worksheets": [ { "name": "S", "headers": ["A"], "rows": [["1"]] } ] }"""u8.ToArray();
        var request = new ConversionRequest(json, "data.json", ConversionTargetFormat.Xlsx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.NotNull(result.OutputBytes);
        Assert.True(result.OutputBytes!.Length > 0);
        Assert.Equal("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", result.ContentType);
        Assert.Equal("data.xlsx", result.OutputFileName);
        Assert.Null(result.ErrorMessage);
        Assert.NotNull(result.Messages);
        Assert.Contains(result.Messages, m => m.Contains("JSON"));
    }

    [Fact]
    public async Task ConvertAsync_InvalidJsonToDocx_ReturnsActionableError()
    {
        var service = new ConversionService();
        var json = """{ "version": "1.0", "sections": [] }"""u8.ToArray();
        var request = new ConversionRequest(json, "report.json", ConversionTargetFormat.Docx);

        var result = await service.ConvertAsync(request);

        Assert.False(result.Success);
        Assert.Null(result.OutputBytes);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("Invalid DOCX generation JSON", result.ErrorMessage);
        Assert.Contains("sections", result.ErrorMessage);
    }

    [Fact]
    public async Task ConvertAsync_InvalidJsonToXlsx_ReturnsActionableError()
    {
        var service = new ConversionService();
        var json = """{ "version": "9.9", "worksheets": [{"name": "S"}] }"""u8.ToArray();
        var request = new ConversionRequest(json, "data.json", ConversionTargetFormat.Xlsx);

        var result = await service.ConvertAsync(request);

        Assert.False(result.Success);
        Assert.Null(result.OutputBytes);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("Invalid XLSX instruction JSON", result.ErrorMessage);
        Assert.Contains("version", result.ErrorMessage);
    }

    [Fact]
    public async Task ConvertAsync_MalformedJsonToXlsx_ReturnsActionableError()
    {
        var service = new ConversionService();
        var json = "{ not valid json"u8.ToArray();
        var request = new ConversionRequest(json, "data.json", ConversionTargetFormat.Xlsx);

        var result = await service.ConvertAsync(request);

        Assert.False(result.Success);
        Assert.Null(result.OutputBytes);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("Invalid XLSX instruction JSON", result.ErrorMessage);
    }

    [Fact]
    public async Task ConvertAsync_NonEmptyJsonToPptx_ReturnsUnsupportedError()
    {
        // JSON sources are routed to the declarative generators for DOCX/XLSX only; a JSON
        // source targeting PPTX is unsupported rather than silently turned into a blank deck.
        var service = new ConversionService();
        var json = """{ "version": "1.0", "worksheets": [] }"""u8.ToArray();
        var request = new ConversionRequest(json, "data.json", ConversionTargetFormat.Pptx);

        var result = await service.ConvertAsync(request);

        Assert.False(result.Success);
        Assert.Null(result.OutputBytes);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("Unsupported conversion", result.ErrorMessage);
    }

    [Fact]
    public async Task ConvertAsync_EmptyJsonToDocx_CreatesBlankInstead()
    {
        // Empty JSON sources keep the generic blank-document compatibility path.
        var service = new ConversionService();
        var request = new ConversionRequest([], "empty.json", ConversionTargetFormat.Docx);

        var result = await service.ConvertAsync(request);

        Assert.True(result.Success);
        Assert.NotNull(result.OutputBytes);
        Assert.Equal("empty.docx", result.OutputFileName);
        Assert.NotNull(result.Messages);
        Assert.Contains(result.Messages, m => m.Contains("blank DOCX"));
    }

    [Fact]
    public async Task ConvertAsync_UnsupportedConversion_ErrorResultHasNonNullMessagesList()
    {
        // When the switch _ arm fires, ConversionResultWithError is called without
        // the messages parameter (it defaults to null, which becomes an empty list).
        // The detection message IS collected but not passed through this code path —
        // the result always carries a non-null list for safe enumeration.
        var service = new ConversionService();
        var bytes = new byte[] { 0x89, (byte)'P', (byte)'N', (byte)'G' };
        var request = new ConversionRequest(bytes, "image.png", ConversionTargetFormat.Pdf);

        var result = await service.ConvertAsync(request);

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("Unsupported conversion", result.ErrorMessage);
        Assert.NotNull(result.Messages);
    }
}
