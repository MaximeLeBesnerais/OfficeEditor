using System.Text;
using DocxEditor.Core.Builders;
using DocxEditor.Core.Generation;
using DocxEditor.Core.Generation.Schema;
using OfficeEditor.Core.Rendering;
using OfficeEditor.Core.Services;

namespace DocxEditor.Core.Rendering;

/// <summary>
/// DOCX format renderer for the shared facade (<see cref="IDocumentRenderer"/>). Renders an
/// existing .docx (or a .docx generated from DOCX generation JSON) to PDF/PNG/SVG/Typst by
/// reusing <see cref="DocumentBuilder.ExportToPdf"/>, <see cref="DocumentBuilder.ExportToPng"/>,
/// <see cref="DocumentBuilder.ExportToSvg"/> and <see cref="DocumentBuilder.ExportToTypst"/>.
/// All output compiles through the repository's own Typst pipeline — no office suite involved.
/// </summary>
[DocumentSourceFormat(".docx", 1)]
internal sealed class DocxFormatRenderer : IFormatRenderer
{
    public DocxFormatRenderer()
    {
    }

    public bool CanRenderSource(string sourcePath)
    {
        var extension = Path.GetExtension(sourcePath);
        if (string.Equals(extension, ".docx", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (!string.Equals(extension, ".json", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        try
        {
            return new DocxGenerationDocumentParser().Validate(File.ReadAllText(sourcePath)).IsValid;
        }
        catch
        {
            return false;
        }
    }

    public DocumentRenderResult Render(DocumentRenderRequest request)
    {
        try
        {
            using var builder = OpenBuilder(request, out var warnings);

            if (request.Format == DocumentOutputFormat.Typ)
            {
                return Success(warnings, Encoding.UTF8.GetBytes(builder.ExportToTypst()));
            }

            var options = new CompileOptions { Ppi = request.Ppi > 0 ? request.Ppi : 150f, FontDirectory = request.FontPath };
            return request.Format switch
            {
                DocumentOutputFormat.Pdf => Success(warnings, builder.ExportToPdf(options)),
                DocumentOutputFormat.Png => Success(warnings, builder.ExportToPng(options)),
                DocumentOutputFormat.Svg => Success(warnings, builder.ExportToSvg(options)),
                _ => Failure($"Unsupported output format '{request.Format}' for DOCX.")
            };
        }
        catch (Exception ex)
        {
            return Failure($"DOCX render failed: {ex.Message}");
        }
    }

    private static DocumentBuilder OpenBuilder(DocumentRenderRequest request, out IReadOnlyList<string> warnings)
    {
        if (!string.Equals(Path.GetExtension(request.SourcePath), ".json", StringComparison.OrdinalIgnoreCase))
        {
            warnings = Array.Empty<string>();
            return (DocumentBuilder)DocumentBuilder.Open(request.SourcePath);
        }

        var generated = new DocxGenerator().GenerateToBytes(File.ReadAllText(request.SourcePath));
        warnings = generated.Result.Warnings
            .Select(warning => warning.ToString())
            .ToList();
        return (DocumentBuilder)DocumentBuilder.Open(generated.Content);
    }

    private static DocumentRenderResult Success(IReadOnlyList<string> warnings, params byte[][] pages) => new()
    {
        Success = true,
        Pages = pages,
        Warnings = warnings
    };

    private static DocumentRenderResult Failure(string message) => new()
    {
        Success = false,
        ErrorMessage = message
    };
}
