using System.Text;
using DocxEditor.Core.Builders;
using DocxEditor.Core.Converters;
using DocumentFormat.OpenXml.Packaging;
using OfficeEditor.Core.Services;
using PptxEditor.Core.Builders;
using XlsxEditor.Core.Builders;

namespace OfficeEditor.Api.Services;

public sealed class ConversionService : IConversionService
{
    public async Task<ConversionResult> ConvertAsync(ConversionRequest request, CancellationToken ct = default)
    {
        var messages = new List<string>();
        try
        {
            var sourceFormat = DetectSourceFormat(request.SourceFileName);
            messages.Add($"Detected source format: {sourceFormat}");

            return request.TargetFormat switch
            {
                ConversionTargetFormat.Pdf when sourceFormat == SourceFormat.Pptx => await ConvertPptxToPdfAsync(request, messages, ct),
                ConversionTargetFormat.Png when sourceFormat == SourceFormat.Pptx => await ConvertPptxToPngAsync(request, messages, ct),
                ConversionTargetFormat.Svg when sourceFormat == SourceFormat.Pptx => await ConvertPptxToSvgAsync(request, messages, ct),
                ConversionTargetFormat.Docx when sourceFormat == SourceFormat.Markdown => await ConvertMarkdownToDocxAsync(request, messages, ct),
                ConversionTargetFormat.Docx when sourceFormat == SourceFormat.Unknown || (sourceFormat == SourceFormat.Docx && IsEmptyDocument(request)) => await CreateBlankDocxAsync(request, messages, ct),
                ConversionTargetFormat.Xlsx when sourceFormat == SourceFormat.Unknown || (sourceFormat == SourceFormat.Xlsx && IsEmptyDocument(request)) => await CreateBlankXlsxAsync(request, messages, ct),
                ConversionTargetFormat.Pptx when sourceFormat == SourceFormat.Unknown || (sourceFormat == SourceFormat.Pptx && IsEmptyDocument(request)) => await CreateBlankPptxAsync(request, messages, ct),
                ConversionTargetFormat.Pdf when sourceFormat == SourceFormat.Docx => await ConvertDocxToPdfAsync(request, messages, ct),
                _ => ConversionResultWithError($"Unsupported conversion: {sourceFormat} to {request.TargetFormat}.")
            };
        }
        catch (Exception ex)
        {
            return ConversionResultWithError($"Conversion failed: {ex.Message}", messages);
        }
    }

    private static async Task<ConversionResult> ConvertPptxToPdfAsync(ConversionRequest request, List<string> messages, CancellationToken ct)
    {
        await Task.Yield();
        ct.ThrowIfCancellationRequested();

        using var builder = PresentationBuilder.Open(request.SourceBytes);
        messages.Add("Opened PPTX presentation.");

        byte[] pdfBytes;
        try
        {
            pdfBytes = builder.ExportToPdf();
            messages.Add("Exported PPTX to PDF.");
        }
        catch
        {
            ct.ThrowIfCancellationRequested();
            var typstSource = builder.ExportToTypst();
            messages.Add("Falling back: exported PPTX to Typst source.");

            using var compiler = new TypstCompilerService();
            var options = new CompileOptions { Format = OutputFormat.Pdf };
            var result = compiler.Compile(typstSource, options);

            if (!result.Success || result.Pages.Length == 0)
            {
                return ConversionResultWithError($"PDF compilation failed: {result.ErrorMessage}", messages);
            }

            pdfBytes = result.Pages[0];
            messages.Add("Compiled Typst source to PDF.");
        }

        return new ConversionResult(
            true,
            pdfBytes,
            ChangeExtension(request.SourceFileName, ".pdf"),
            "application/pdf",
            null,
            messages);
    }

    private static async Task<ConversionResult> ConvertPptxToPngAsync(ConversionRequest request, List<string> messages, CancellationToken ct)
    {
        await Task.Yield();
        ct.ThrowIfCancellationRequested();

        using var builder = PresentationBuilder.Open(request.SourceBytes);
        messages.Add("Opened PPTX presentation.");

        var ppi = GetPpi(request.Options);
        var thumbnails = builder.ExportThumbnails(new ThumbnailOptions { Ppi = ppi, Format = "png" });
        messages.Add($"Exported {thumbnails.Length} slide thumbnail(s).");

        if (thumbnails.Length == 0)
        {
            return ConversionResultWithError("PNG export produced no images.", messages);
        }

        messages.Add($"Slide 1 of {thumbnails.Length} rendered.");
        return new ConversionResult(
            true,
            thumbnails[0],
            ChangeExtension(request.SourceFileName, ".png"),
            "image/png",
            null,
            messages);
    }

    private static async Task<ConversionResult> ConvertPptxToSvgAsync(ConversionRequest request, List<string> messages, CancellationToken ct)
    {
        await Task.Yield();
        ct.ThrowIfCancellationRequested();

        using var builder = PresentationBuilder.Open(request.SourceBytes);
        messages.Add("Opened PPTX presentation.");

        var typstSource = builder.ExportToTypst();
        messages.Add("Exported PPTX to Typst source.");

        using var compiler = new TypstCompilerService();
        var result = compiler.Compile(typstSource, new CompileOptions { Format = OutputFormat.Svg });

        if (!result.Success || result.Pages.Length == 0)
        {
            return ConversionResultWithError($"SVG compilation failed: {result.ErrorMessage}", messages);
        }

        messages.Add($"Compiled Typst source to SVG ({result.Pages.Length} page(s)).");
        return new ConversionResult(
            true,
            result.Pages[0],
            ChangeExtension(request.SourceFileName, ".svg"),
            "image/svg+xml",
            null,
            messages);
    }

    private static async Task<ConversionResult> ConvertMarkdownToDocxAsync(ConversionRequest request, List<string> messages, CancellationToken ct)
    {
        await Task.Yield();
        ct.ThrowIfCancellationRequested();

        var markdown = Encoding.UTF8.GetString(request.SourceBytes);
        messages.Add($"Decoded {markdown.Length} characters of Markdown.");

        using var builder = DocumentBuilder.Create();
        builder.AddMarkdown(markdown);
        messages.Add("Converted Markdown to DOCX content.");

        var docxBytes = builder.SaveToBytes();
        messages.Add("Saved DOCX to memory.");

        return new ConversionResult(
            true,
            docxBytes,
            ChangeExtension(request.SourceFileName, ".docx"),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            null,
            messages);
    }

    private static async Task<ConversionResult> CreateBlankDocxAsync(ConversionRequest request, List<string> messages, CancellationToken ct)
    {
        await Task.Yield();
        ct.ThrowIfCancellationRequested();

        using var builder = DocumentBuilder.Create();
        builder.AddParagraph("Blank document created by OfficeEditor.", "Normal");
        messages.Add("Created blank DOCX with a sample paragraph.");

        var docxBytes = builder.SaveToBytes();
        messages.Add("Saved DOCX to memory.");

        return new ConversionResult(
            true,
            docxBytes,
            string.IsNullOrWhiteSpace(request.SourceFileName) ? "document.docx" : ChangeExtension(request.SourceFileName, ".docx"),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            null,
            messages);
    }

    private static async Task<ConversionResult> CreateBlankXlsxAsync(ConversionRequest request, List<string> messages, CancellationToken ct)
    {
        await Task.Yield();
        ct.ThrowIfCancellationRequested();

        using var builder = WorkbookBuilder.Create();
        builder.AddWorksheet("Sheet1");
        messages.Add("Created blank XLSX workbook with Sheet1.");

        var xlsxBytes = builder.SaveToBytes();
        messages.Add("Saved XLSX to memory.");

        return new ConversionResult(
            true,
            xlsxBytes,
            string.IsNullOrWhiteSpace(request.SourceFileName) ? "workbook.xlsx" : ChangeExtension(request.SourceFileName, ".xlsx"),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            null,
            messages);
    }

    private static async Task<ConversionResult> CreateBlankPptxAsync(ConversionRequest request, List<string> messages, CancellationToken ct)
    {
        await Task.Yield();
        ct.ThrowIfCancellationRequested();

        using var builder = PresentationBuilder.Create();
        builder.AddSlide();
        messages.Add("Created blank PPTX presentation with one slide.");

        var pptxBytes = builder.SaveToBytes();
        messages.Add("Saved PPTX to memory.");

        return new ConversionResult(
            true,
            pptxBytes,
            string.IsNullOrWhiteSpace(request.SourceFileName) ? "presentation.pptx" : ChangeExtension(request.SourceFileName, ".pptx"),
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            null,
            messages);
    }

    private static async Task<ConversionResult> ConvertDocxToPdfAsync(ConversionRequest request, List<string> messages, CancellationToken ct)
    {
        await Task.Yield();
        ct.ThrowIfCancellationRequested();

        using var stream = new MemoryStream(request.SourceBytes, writable: false);
        using var document = WordprocessingDocument.Open(stream, false);
        messages.Add("Opened DOCX document.");

        using var converter = new DocxToTypstConverter(document);
        var typstDocument = converter.Convert();
        var typstSource = converter.GenerateTypstSource(typstDocument);
        messages.Add("Converted DOCX to Typst source.");

        using var compiler = new TypstCompilerService();
        var result = compiler.Compile(typstSource, new CompileOptions
        {
            Format = OutputFormat.Pdf,
            WorkingDirectory = typstDocument.TempDirectory
        });

        if (!result.Success || result.Pages.Length == 0)
        {
            return ConversionResultWithError($"PDF compilation failed: {result.ErrorMessage}", messages);
        }

        messages.Add("Compiled Typst source to PDF.");
        return new ConversionResult(
            true,
            result.Pages[0],
            ChangeExtension(request.SourceFileName, ".pdf"),
            "application/pdf",
            null,
            messages);
    }

    private static SourceFormat DetectSourceFormat(string? fileName)
    {
        var extension = Path.GetExtension(fileName ?? string.Empty).TrimStart('.').ToLowerInvariant();
        return extension switch
        {
            "docx" => SourceFormat.Docx,
            "pptx" => SourceFormat.Pptx,
            "xlsx" => SourceFormat.Xlsx,
            "md" or "markdown" or "txt" => SourceFormat.Markdown,
            _ => SourceFormat.Unknown
        };
    }

    private static bool IsEmptyDocument(ConversionRequest request)
    {
        return request.SourceBytes.Length == 0
            || request.SourceBytes.All(b => b == 0);
    }

    private static float GetPpi(IReadOnlyDictionary<string, string>? options)
    {
        if (options is not null
            && options.TryGetValue("Ppi", out var ppiText)
            && float.TryParse(ppiText, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var ppi))
        {
            return ppi;
        }

        return 150;
    }

    private static string ChangeExtension(string? fileName, string extension)
    {
        if (string.IsNullOrWhiteSpace(fileName))
        {
            return $"document{extension}";
        }

        var baseName = Path.GetFileNameWithoutExtension(fileName);
        if (string.IsNullOrWhiteSpace(baseName))
        {
            return $"document{extension}";
        }

        return baseName + extension;
    }

    private static ConversionResult ConversionResultWithError(string errorMessage, List<string>? messages = null)
    {
        return new ConversionResult(
            false,
            null,
            string.Empty,
            string.Empty,
            errorMessage,
            messages ?? new List<string>());
    }

    private enum SourceFormat
    {
        Unknown,
        Docx,
        Pptx,
        Xlsx,
        Markdown
    }
}
