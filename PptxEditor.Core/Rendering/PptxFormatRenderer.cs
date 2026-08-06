using System.Text;
using OfficeEditor.Core.Rendering;
using PptxEditor.Core.Builders;
using PptxEditor.Core.Generation.Archetypes;
using PptxEditor.Core.Generation.Components;
using PptxEditor.Core.Generation.Emit.Ooxml;
using PptxEditor.Core.Generation.Layout;
using PptxEditor.Core.Generation.Schema;
using PptxEditor.Core.Services;

namespace PptxEditor.Core.Rendering;

/// <summary>
/// PPTX format renderer for the shared facade (<see cref="IDocumentRenderer"/>). Renders an
/// existing .pptx (or a .pptx generated from PPTX generation JSON) to PDF/PNG/SVG/Typst by
/// reusing <see cref="PresentationBuilder.ExportToPdf"/>, <see cref="PresentationBuilder.ExportThumbnails"/>
/// (PNG and SVG) and <see cref="PresentationBuilder.ExportToTypst"/>. No office suite is involved;
/// everything compiles through the repository's own Typst pipeline.
/// </summary>
[DocumentSourceFormat(".pptx", 2)]
internal sealed class PptxFormatRenderer : IFormatRenderer
{
    public PptxFormatRenderer()
    {
    }

    public bool CanRenderSource(string sourcePath)
    {
        var extension = Path.GetExtension(sourcePath);
        if (string.Equals(extension, ".pptx", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (!string.Equals(extension, ".json", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        try
        {
            return new GenerationDocumentParser().Validate(File.ReadAllText(sourcePath)).IsValid;
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
            byte[]? generatedPptx = IsJsonSource(request) ? GeneratePptxFromJson(request.SourcePath) : null;

            if (request.Format == DocumentOutputFormat.Typ)
            {
                using var builder = Open(request, generatedPptx);
                return Success(Encoding.UTF8.GetBytes(builder.ExportToTypst()));
            }

            var ppi = request.Ppi > 0 ? request.Ppi : 150f;
            using (var builder = Open(request, generatedPptx))
            {
                return request.Format switch
                {
                    DocumentOutputFormat.Pdf => CompilePdf(builder, request.FontPath),
                    DocumentOutputFormat.Png => CompileImages(builder, "png", ppi, request.FontPath),
                    DocumentOutputFormat.Svg => CompileImages(builder, "svg", ppi, request.FontPath),
                    _ => Failure($"Unsupported output format '{request.Format}' for PPTX.")
                };
            }
        }
        catch (Exception ex)
        {
            return Failure($"PPTX render failed: {ex.Message}");
        }
    }

    private static IPresentationBuilder Open(DocumentRenderRequest request, byte[]? generatedPptx)
        => generatedPptx is not null
            ? PresentationBuilder.Open(generatedPptx)
            // Rendering must never mutate the source deck: open from an in-memory copy so the
            // read-write OpenXML document is discarded when the builder is disposed.
            : PresentationBuilder.Open(File.ReadAllBytes(request.SourcePath));

    private static DocumentRenderResult CompilePdf(IPresentationBuilder builder, string? fontPath)
    {
        var pdf = builder.ExportToPdf(new PdfOptions { FontDirectory = fontPath });
        return pdf.Length == 0 ? Failure("PDF export produced no output.") : Success(pdf);
    }

    private static DocumentRenderResult CompileImages(IPresentationBuilder builder, string format, float ppi, string? fontPath)
    {
        var pages = builder.ExportThumbnails(new ThumbnailOptions
        {
            Ppi = ppi,
            Format = format,
            FontDirectory = fontPath
        });

        return pages.Length == 0
            ? Failure($"{format.ToUpperInvariant()} export produced no output.")
            : Success(pages);
    }

    private static bool IsJsonSource(DocumentRenderRequest request)
        => string.Equals(Path.GetExtension(request.SourcePath), ".json", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Runs the declarative generation pipeline (validate → archetype/component expand → layout →
    /// OOXML emit) to produce an in-memory .pptx from PPTX generation JSON, mirroring the API's
    /// deck-generation service but without its HTTP/repository-confinement policy layer.
    /// </summary>
    private static byte[] GeneratePptxFromJson(string jsonPath)
    {
        var validation = new GenerationDocumentParser().Validate(File.ReadAllText(jsonPath));
        if (!validation.IsValid)
        {
            throw new InvalidOperationException(
                $"Invalid PPTX generation JSON: {string.Join("; ", validation.Errors.Select(error => error.Message))}");
        }

        LayoutResult layout;
        try
        {
            var archetyped = ArchetypeExpander.Expand(validation.Document!);
            var expanded = ComponentExpander.Expand(archetyped);
            layout = new LayoutResolver(new TextMeasure(new FontMetricsCatalog())).Resolve(expanded);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"PPTX generation layout failed: {ex.Message}");
        }

        // The emitter resolves relative image sources against the JSON document's
        // directory only (canonical, symlink-aware containment), so the directory of the
        // source file is threaded through.
        return new OoxmlEmitter(new OoxmlEmitOptions
        {
            DocumentDirectory = Path.GetDirectoryName(Path.GetFullPath(jsonPath))
        }).Emit(layout).Bytes;
    }

    private static DocumentRenderResult Success(params byte[][] pages) => new()
    {
        Success = true,
        Pages = pages
    };

    private static DocumentRenderResult Failure(string message) => new()
    {
        Success = false,
        ErrorMessage = message
    };
}
