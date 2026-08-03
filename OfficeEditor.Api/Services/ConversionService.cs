using System.Text;
using System.Text.Json;
using DocxEditor.Core.Builders;
using DocxEditor.Core.Generation;
using DocxEditor.Core.Generation.Schema;
using DocxEditor.Core.Markdown.Rendering;
using OfficeEditor.Core.Rendering;
using PptxEditor.Core.Builders;
using XlsxEditor.Core.Builders;
using XlsxEditor.Core.Instructions;

namespace OfficeEditor.Api.Services;

public sealed class ConversionService : IConversionService
{
    // All PDF/PNG/SVG rendering delegates to the shared render facade (CLI/API/MCP parity).
    // Constructed lazily because renderer discovery scans loaded assemblies, and the vast
    // majority of conversions are generation-only.
    private readonly Lazy<IDocumentRenderer> _renderer = new(() => new DocumentRenderer());

    public async Task<ConversionResult> ConvertAsync(ConversionRequest request, CancellationToken ct = default)
    {
        var messages = new List<string>();
        try
        {
            var sourceFormat = DetectSourceFormat(request.SourceFileName);
            messages.Add($"Detected source format: {sourceFormat}");

            return request.TargetFormat switch
            {
                ConversionTargetFormat.Pdf when sourceFormat == SourceFormat.Pptx => await RenderViaFacadeAsync(request, sourceFormat, DocumentOutputFormat.Pdf, ".pdf", "application/pdf", messages, ct),
                ConversionTargetFormat.Png when sourceFormat == SourceFormat.Pptx => await RenderViaFacadeAsync(request, sourceFormat, DocumentOutputFormat.Png, ".png", "image/png", messages, ct),
                ConversionTargetFormat.Svg when sourceFormat == SourceFormat.Pptx => await RenderViaFacadeAsync(request, sourceFormat, DocumentOutputFormat.Svg, ".svg", "image/svg+xml", messages, ct),
                ConversionTargetFormat.Docx when sourceFormat == SourceFormat.Markdown => await ConvertMarkdownToDocxAsync(request, messages, ct),
                // Non-empty JSON sources route through the declarative generators, never blank
                // creation; empty JSON falls back to blank-document compatibility.
                ConversionTargetFormat.Docx when sourceFormat == SourceFormat.Json && !IsEmptyDocument(request) => await ConvertJsonToDocxAsync(request, messages, ct),
                ConversionTargetFormat.Xlsx when sourceFormat == SourceFormat.Json && !IsEmptyDocument(request) => await ConvertJsonToXlsxAsync(request, messages, ct),
                ConversionTargetFormat.Docx when sourceFormat == SourceFormat.Json && IsEmptyDocument(request) => await CreateBlankDocxAsync(request, messages, ct),
                ConversionTargetFormat.Xlsx when sourceFormat == SourceFormat.Json && IsEmptyDocument(request) => await CreateBlankXlsxAsync(request, messages, ct),
                ConversionTargetFormat.Docx when sourceFormat == SourceFormat.Unknown || (sourceFormat == SourceFormat.Docx && IsEmptyDocument(request)) => await CreateBlankDocxAsync(request, messages, ct),
                ConversionTargetFormat.Xlsx when sourceFormat == SourceFormat.Unknown || (sourceFormat == SourceFormat.Xlsx && IsEmptyDocument(request)) => await CreateBlankXlsxAsync(request, messages, ct),
                ConversionTargetFormat.Pptx when sourceFormat == SourceFormat.Unknown || (sourceFormat == SourceFormat.Pptx && IsEmptyDocument(request)) => await CreateBlankPptxAsync(request, messages, ct),
                ConversionTargetFormat.Pdf when sourceFormat == SourceFormat.Docx => await RenderViaFacadeAsync(request, sourceFormat, DocumentOutputFormat.Pdf, ".pdf", "application/pdf", messages, ct),
                ConversionTargetFormat.Png when sourceFormat == SourceFormat.Docx => await RenderViaFacadeAsync(request, sourceFormat, DocumentOutputFormat.Png, "-1.png", "image/png", messages, ct),
                ConversionTargetFormat.Svg when sourceFormat == SourceFormat.Docx => await RenderViaFacadeAsync(request, sourceFormat, DocumentOutputFormat.Svg, "-1.svg", "image/svg+xml", messages, ct),
                ConversionTargetFormat.Pdf when sourceFormat == SourceFormat.Xlsx => await RenderViaFacadeAsync(request, sourceFormat, DocumentOutputFormat.Pdf, ".pdf", "application/pdf", messages, ct),
                ConversionTargetFormat.Png when sourceFormat == SourceFormat.Xlsx => await RenderViaFacadeAsync(request, sourceFormat, DocumentOutputFormat.Png, "-1.png", "image/png", messages, ct),
                ConversionTargetFormat.Svg when sourceFormat == SourceFormat.Xlsx => await RenderViaFacadeAsync(request, sourceFormat, DocumentOutputFormat.Svg, "-1.svg", "image/svg+xml", messages, ct),
                _ => ConversionResultWithError($"Unsupported conversion: {sourceFormat} to {request.TargetFormat}.")
            };
        }
        catch (Exception ex)
        {
            return ConversionResultWithError($"Conversion failed: {ex.Message}", messages);
        }
    }

    /// <summary>
    /// Renders a binary source (PPTX/DOCX/XLSX) to PDF/PNG/SVG through the shared render facade
    /// (<see cref="IDocumentRenderer"/>). The facade takes a source file path, so the uploaded
    /// bytes are staged in a uniquely named temp file under the process temp root (following the
    /// repo's <c>officeeditor-*</c> convention) and deleted in a finally block. PNG/SVG yield one
    /// buffer per page but the API contract is a single buffer, so the first page is returned;
    /// <paramref name="outputExtension"/> carries the exact suffix (".png" for PPTX, "-1.png" for
    /// DOCX/XLSX, …). Failures are returned as failed results, never thrown.
    /// </summary>
    private async Task<ConversionResult> RenderViaFacadeAsync(
        ConversionRequest request,
        SourceFormat sourceFormat,
        DocumentOutputFormat format,
        string outputExtension,
        string contentType,
        List<string> messages,
        CancellationToken ct)
    {
        await Task.Yield();
        ct.ThrowIfCancellationRequested();

        string? tempPath = null;
        try
        {
            tempPath = WriteSourceToTempFile(request);
            var result = _renderer.Value.Render(new DocumentRenderRequest
            {
                SourcePath = tempPath,
                Format = format,
                Ppi = GetPpi(request.Options)
            });

            if (!result.Success)
            {
                return ConversionResultWithError($"Conversion failed: {result.ErrorMessage}", messages);
            }

            if (result.Pages.Length == 0)
            {
                return ConversionResultWithError("Conversion failed: the render produced no output.", messages);
            }

            messages.Add($"Rendered {sourceFormat.ToString().ToUpperInvariant()} to {format} via the shared render facade (Typst pipeline).");
            return new ConversionResult(
                true,
                result.Pages[0],
                ChangeExtension(request.SourceFileName, outputExtension),
                contentType,
                null,
                messages);
        }
        finally
        {
            if (tempPath is not null)
            {
                try
                {
                    File.Delete(tempPath);
                }
                catch
                {
                    // Best-effort cleanup: a lingering temp file is preferable to failing an
                    // otherwise successful conversion over cleanup.
                }
            }
        }
    }

    private static string WriteSourceToTempFile(ConversionRequest request)
    {
        var extension = Path.GetExtension(request.SourceFileName);
        if (string.IsNullOrWhiteSpace(extension))
        {
            extension = "." + request.TargetFormat.ToString().ToLowerInvariant();
        }

        var tempPath = Path.Combine(Path.GetTempPath(), "officeeditor-convert-" + Guid.NewGuid().ToString("N") + extension);
        File.WriteAllBytes(tempPath, request.SourceBytes);
        return tempPath;
    }

    private static async Task<ConversionResult> ConvertMarkdownToDocxAsync(ConversionRequest request, List<string> messages, CancellationToken ct)
    {
        await Task.Yield();
        ct.ThrowIfCancellationRequested();

        var markdown = Encoding.UTF8.GetString(request.SourceBytes);
        messages.Add($"Decoded {markdown.Length} characters of Markdown.");

        using var builder = DocumentBuilder.Create();
        // Untrusted uploads must never resolve local-file image sources: the API image policy
        // (data URIs only) disables every local path while keeping data-URI images embeddable.
        builder.AddRichMarkdown(markdown, new MarkdownRenderOptions
        {
            ImageSourceOptions = ConversionSourcePolicy.DataUriOnlyImageSources
        });
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

    /// <summary>
    /// Routes a non-empty JSON source to the declarative DOCX generator. Validation failures
    /// return every collected path-qualified issue as an actionable error and never produce
    /// partial output; network image fetching is never enabled. Untrusted input is constrained
    /// to the API policy: a top-level <c>template</c> path is rejected before generation, and
    /// image sources are data URIs only.
    /// </summary>
    private static async Task<ConversionResult> ConvertJsonToDocxAsync(ConversionRequest request, List<string> messages, CancellationToken ct)
    {
        await Task.Yield();
        ct.ThrowIfCancellationRequested();

        var json = Encoding.UTF8.GetString(request.SourceBytes);
        messages.Add($"Decoded {json.Length} characters of JSON.");

        if (RejectJsonTemplatePath(json) is { } templateError)
        {
            messages.Add("Rejected 'template' in DOCX generation JSON.");
            return ConversionResultWithError($"Invalid DOCX generation JSON:{Environment.NewLine} - {templateError}", messages);
        }

        var generator = new DocxGenerator();
        GeneratedDocx generated;
        try
        {
            generated = generator.GenerateToBytes(json, new DocxGeneratorOptions
            {
                ImageSourceOptions = ConversionSourcePolicy.DataUriOnlyImageSources
            });
        }
        catch (DocxGenerationValidationException ex)
        {
            var details = string.Join(Environment.NewLine + " - ", ex.Issues);
            messages.Add("Rejected invalid DOCX generation JSON.");
            return ConversionResultWithError($"Invalid DOCX generation JSON:{Environment.NewLine} - {details}", messages);
        }

        foreach (var warning in generated.Result.Warnings)
        {
            messages.Add($"Warning: {warning}");
        }

        messages.Add($"Generated DOCX package ({generated.Content.Length} bytes).");
        return new ConversionResult(
            true,
            generated.Content,
            ChangeExtension(request.SourceFileName, ".docx"),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            null,
            messages);
    }

    /// <summary>
    /// Routes a non-empty JSON source to the JSON→XLSX generator. The generator's canonical
    /// diagnostics (JSON path + message) surface as an actionable error on rejection; nothing
    /// is written on failure.
    /// </summary>
    private static async Task<ConversionResult> ConvertJsonToXlsxAsync(ConversionRequest request, List<string> messages, CancellationToken ct)
    {
        await Task.Yield();
        ct.ThrowIfCancellationRequested();

        var json = Encoding.UTF8.GetString(request.SourceBytes);
        messages.Add($"Decoded {json.Length} characters of JSON.");

        var result = XlsxGenerator.Generate(json);
        if (!result.IsValid || result.Bytes is null)
        {
            var details = string.Join(Environment.NewLine + " - ", result.Validation.Errors.Select(FormatXlsxDiagnostic));
            messages.Add("Rejected invalid XLSX instruction JSON.");
            return ConversionResultWithError($"Invalid XLSX instruction JSON:{Environment.NewLine} - {details}", messages);
        }

        foreach (var warning in result.Validation.Warnings)
        {
            messages.Add($"Warning: {FormatXlsxDiagnostic(warning)}");
        }

        messages.Add($"Generated XLSX workbook ({result.Bytes.Length} bytes).");
        return new ConversionResult(
            true,
            result.Bytes,
            ChangeExtension(request.SourceFileName, ".xlsx"),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            null,
            messages);
    }

    private static string FormatXlsxDiagnostic(XlsxDiagnostic diagnostic)
    {
        var location = string.IsNullOrEmpty(diagnostic.Path) ? diagnostic.CodeName : $"{diagnostic.Path} ({diagnostic.CodeName})";
        return $"{diagnostic.Message} at {location}";
    }

    /// <summary>
    /// Rejects a top-level <c>template</c> path in untrusted generation JSON. The HTTP
    /// conversion service must never read a server-local template file that an upload names
    /// (<see cref="DocxGeneratorOptions.TemplatePath"/> reaches <c>File.ReadAllBytes</c> in the
    /// emitter). Only a non-empty string <c>template</c> is rejected here; malformed JSON and
    /// empty/non-string values fall through so the generator reports its own actionable
    /// validation error. Returns the path-qualified issue message, or null when allowed.
    /// </summary>
    private static string? RejectJsonTemplatePath(string json)
    {
        try
        {
            using var document = JsonDocument.Parse(json);
            var root = document.RootElement;
            if (root.ValueKind != JsonValueKind.Object)
            {
                return null;
            }

            foreach (var property in root.EnumerateObject())
            {
                if (!string.Equals(property.Name, "template", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (property.Value.ValueKind == JsonValueKind.String
                    && !string.IsNullOrWhiteSpace(property.Value.GetString()))
                {
                    return "\"$.template\": template files are not allowed for API conversions; remove 'template' to generate from a blank document.";
                }

                return null;
            }

            return null;
        }
        catch (JsonException)
        {
            return null;
        }
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

    private static SourceFormat DetectSourceFormat(string? fileName)
    {
        var extension = Path.GetExtension(fileName ?? string.Empty).TrimStart('.').ToLowerInvariant();
        return extension switch
        {
            "docx" => SourceFormat.Docx,
            "pptx" => SourceFormat.Pptx,
            "xlsx" => SourceFormat.Xlsx,
            "md" or "markdown" or "txt" => SourceFormat.Markdown,
            "json" => SourceFormat.Json,
            _ => SourceFormat.Unknown
        };
    }

    private static bool IsEmptyDocument(ConversionRequest request)
    {
        return request.SourceBytes.Length == 0
            || request.SourceBytes.All(b => b == 0);
    }

    internal static float GetPpi(IReadOnlyDictionary<string, string>? options)
    {
        if (options is not null
            && options.TryGetValue("Ppi", out var ppiText)
            && float.TryParse(ppiText, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var ppi))
        {
            return float.IsFinite(ppi)
                ? Math.Clamp(ppi, DeckPreviewValidators.MinPpi, DeckPreviewValidators.MaxPpi)
                : DeckPreviewValidators.DefaultPpi;
        }

        return DeckPreviewValidators.DefaultPpi;
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
        Markdown,
        Json
    }
}
