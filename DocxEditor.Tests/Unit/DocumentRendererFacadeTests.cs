using System.Text;
using DocxEditor.Core.Builders;
using OfficeEditor.Core.Rendering;
using PptxEditor.Core.Builders;
using XlsxEditor.Core.Builders;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Facade tests for the shared document render API (<see cref="DocumentRenderer"/> /
/// <see cref="IDocumentRenderer"/>): one request type rendering PPTX, DOCX and XLSX to
/// PDF/PNG/SVG/Typst, JSON vocabulary dispatch, and clean (non-throwing) failure results.
/// Follows the repo's compile smoke-test convention (see <c>ExportThumbnailTests</c>): tiny
/// in-memory documents compiled through the repository's own Typst pipeline (TypstBridge
/// primary, typst CLI fallback) — no office suite is involved.
/// </summary>
public sealed class DocumentRendererFacadeTests : IDisposable
{
    private readonly string _tempDirectory;

    public DocumentRendererFacadeTests()
    {
        _tempDirectory = Path.Combine(Path.GetTempPath(), "officeeditor-facade-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDirectory);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDirectory))
            {
                Directory.Delete(_tempDirectory, recursive: true);
            }
        }
        catch
        {
            // Best-effort cleanup of the test scratch directory.
        }
    }

    // ─── PPTX → all formats ─────────────────────────────────────────

    [Fact]
    public void Render_Pptx_ToPdf_ReturnsSinglePdfBuffer()
    {
        using var renderer = new RendererScope();
        var result = renderer.Renderer.Render(new DocumentRenderRequest
        {
            SourcePath = WriteTempFile(".pptx", CreatePptxBytes()),
            Format = DocumentOutputFormat.Pdf
        });

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Single(result.Pages);
        AssertPdfSignature(result.Pages[0]);
    }

    [Fact]
    public void Render_Pptx_ToPng_ReturnsOneBufferPerSlide()
    {
        using var renderer = new RendererScope();
        var result = renderer.Renderer.Render(new DocumentRenderRequest
        {
            SourcePath = WriteTempFile(".pptx", CreatePptxBytes(2)),
            Format = DocumentOutputFormat.Png,
            Ppi = 72
        });

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(2, result.Pages.Length);
        foreach (var page in result.Pages)
        {
            AssertPngSignature(page);
        }
    }

    [Fact]
    public void Render_Pptx_ToSvg_ReturnsOneBufferPerSlide()
    {
        using var renderer = new RendererScope();
        var result = renderer.Renderer.Render(new DocumentRenderRequest
        {
            SourcePath = WriteTempFile(".pptx", CreatePptxBytes(2)),
            Format = DocumentOutputFormat.Svg
        });

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(2, result.Pages.Length);
        foreach (var page in result.Pages)
        {
            var head = Encoding.ASCII.GetString(page, 0, Math.Min(page.Length, 512));
            Assert.Contains("<svg", head);
        }
    }

    [Fact]
    public void Render_Pptx_ToTyp_ReturnsTypstSource()
    {
        using var renderer = new RendererScope();
        var result = renderer.Renderer.Render(new DocumentRenderRequest
        {
            SourcePath = WriteTempFile(".pptx", CreatePptxBytes(1)),
            Format = DocumentOutputFormat.Typ
        });

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Single(result.Pages);
        AssertTypstSource(result.Pages[0]);
    }

    // ─── DOCX → all formats ─────────────────────────────────────────

    [Fact]
    public void Render_Docx_ToPdf_ReturnsSinglePdfBuffer()
    {
        using var renderer = new RendererScope();
        var result = renderer.Renderer.Render(new DocumentRenderRequest
        {
            SourcePath = WriteTempFile(".docx", CreateDocxBytes()),
            Format = DocumentOutputFormat.Pdf
        });

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Single(result.Pages);
        AssertPdfSignature(result.Pages[0]);
    }

    [Fact]
    public void Render_Docx_ToPng_ReturnsImagePages()
    {
        using var renderer = new RendererScope();
        var result = renderer.Renderer.Render(new DocumentRenderRequest
        {
            SourcePath = WriteTempFile(".docx", CreateDocxBytes()),
            Format = DocumentOutputFormat.Png
        });

        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.Pages);
        foreach (var page in result.Pages)
        {
            AssertPngSignature(page);
        }
    }

    [Fact]
    public void Render_Docx_ToSvg_ReturnsSvgPages()
    {
        using var renderer = new RendererScope();
        var result = renderer.Renderer.Render(new DocumentRenderRequest
        {
            SourcePath = WriteTempFile(".docx", CreateDocxBytes()),
            Format = DocumentOutputFormat.Svg
        });

        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.Pages);
        foreach (var page in result.Pages)
        {
            var head = Encoding.ASCII.GetString(page, 0, Math.Min(page.Length, 512));
            Assert.Contains("<svg", head);
        }
    }

    [Fact]
    public void Render_Docx_ToTyp_ReturnsTypstSource()
    {
        using var renderer = new RendererScope();
        var result = renderer.Renderer.Render(new DocumentRenderRequest
        {
            SourcePath = WriteTempFile(".docx", CreateDocxBytes()),
            Format = DocumentOutputFormat.Typ
        });

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Single(result.Pages);
        AssertTypstSource(result.Pages[0]);
    }

    // ─── XLSX → all formats ─────────────────────────────────────────

    [Fact]
    public void Render_Xlsx_ToPdf_ReturnsSinglePdfBuffer()
    {
        using var renderer = new RendererScope();
        var result = renderer.Renderer.Render(new DocumentRenderRequest
        {
            SourcePath = WriteTempFile(".xlsx", CreateXlsxBytes()),
            Format = DocumentOutputFormat.Pdf
        });

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Single(result.Pages);
        AssertPdfSignature(result.Pages[0]);
    }

    [Fact]
    public void Render_Xlsx_ToPng_ReturnsImagePages()
    {
        using var renderer = new RendererScope();
        var result = renderer.Renderer.Render(new DocumentRenderRequest
        {
            SourcePath = WriteTempFile(".xlsx", CreateXlsxBytes()),
            Format = DocumentOutputFormat.Png
        });

        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.Pages);
        foreach (var page in result.Pages)
        {
            AssertPngSignature(page);
        }
    }

    [Fact]
    public void Render_Xlsx_ToSvg_ReturnsSvgPages()
    {
        using var renderer = new RendererScope();
        var result = renderer.Renderer.Render(new DocumentRenderRequest
        {
            SourcePath = WriteTempFile(".xlsx", CreateXlsxBytes()),
            Format = DocumentOutputFormat.Svg
        });

        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.Pages);
        foreach (var page in result.Pages)
        {
            var head = Encoding.ASCII.GetString(page, 0, Math.Min(page.Length, 512));
            Assert.Contains("<svg", head);
        }
    }

    [Fact]
    public void Render_Xlsx_ToTyp_ReturnsTypstSource()
    {
        using var renderer = new RendererScope();
        var result = renderer.Renderer.Render(new DocumentRenderRequest
        {
            SourcePath = WriteTempFile(".xlsx", CreateXlsxBytes()),
            Format = DocumentOutputFormat.Typ
        });

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Single(result.Pages);
        AssertTypstSource(result.Pages[0]);
    }

    // ─── JSON dispatch ──────────────────────────────────────────────

    [Fact]
    public void Render_JsonXlsxVocabulary_ToPdf_RunsUnconditionally()
    {
        // The .json → XLSX → PDF path is fast (no document conversion) and must run without
        // the Typst compile opt-in gate: it exercises JSON detection, generation, and the
        // full render pipeline from one request type.
        using var renderer = new RendererScope();
        var result = renderer.Renderer.Render(new DocumentRenderRequest
        {
            SourcePath = WriteTempFile(".json", Encoding.UTF8.GetBytes(XlsxJson)),
            Format = DocumentOutputFormat.Pdf
        });

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Single(result.Pages);
        AssertPdfSignature(result.Pages[0]);
    }

    [Fact]
    public void Render_JsonXlsxVocabulary_ToTyp_RoutesToXlsxRenderer()
    {
        using var renderer = new RendererScope();
        var result = renderer.Renderer.Render(new DocumentRenderRequest
        {
            SourcePath = WriteTempFile(".json", Encoding.UTF8.GetBytes(XlsxJson)),
            Format = DocumentOutputFormat.Typ
        });

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Single(result.Pages);
        // The Typst table source references the worksheet content.
        Assert.Contains("Sheet1", Encoding.UTF8.GetString(result.Pages[0]));
    }

    [Fact]
    public void Render_JsonDocxVocabulary_ToPdf_RoutesToDocxRenderer()
    {
        using var renderer = new RendererScope();
        var result = renderer.Renderer.Render(new DocumentRenderRequest
        {
            SourcePath = WriteTempFile(".json", Encoding.UTF8.GetBytes(DocxJson)),
            Format = DocumentOutputFormat.Pdf
        });

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Single(result.Pages);
        AssertPdfSignature(result.Pages[0]);
    }

    [Fact]
    public void Render_JsonDocxVocabulary_ToTyp_RoutesToDocxRenderer()
    {
        using var renderer = new RendererScope();
        var result = renderer.Renderer.Render(new DocumentRenderRequest
        {
            SourcePath = WriteTempFile(".json", Encoding.UTF8.GetBytes(DocxJson)),
            Format = DocumentOutputFormat.Typ
        });

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Single(result.Pages);
        Assert.Contains("Hello facade", Encoding.UTF8.GetString(result.Pages[0]));
    }

    [Fact]
    public void Render_JsonPptxVocabulary_ToSvg_RoutesToPptxRenderer()
    {
        using var renderer = new RendererScope();
        var result = renderer.Renderer.Render(new DocumentRenderRequest
        {
            SourcePath = WriteTempFile(".json", Encoding.UTF8.GetBytes(PptxJson)),
            Format = DocumentOutputFormat.Svg
        });

        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.Pages);
        var head = Encoding.ASCII.GetString(result.Pages[0], 0, Math.Min(result.Pages[0].Length, 512));
        Assert.Contains("<svg", head);
    }

    [Fact]
    public void Render_JsonPptxVocabulary_ToTyp_RoutesToPptxRenderer()
    {
        using var renderer = new RendererScope();
        var result = renderer.Renderer.Render(new DocumentRenderRequest
        {
            SourcePath = WriteTempFile(".json", Encoding.UTF8.GetBytes(PptxJson)),
            Format = DocumentOutputFormat.Typ
        });

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Single(result.Pages);
        AssertTypstSource(result.Pages[0]);
    }

    // ─── Failure results, never crashes ─────────────────────────────

    [Fact]
    public void Render_UnsupportedExtension_ReturnsCleanError()
    {
        using var renderer = new RendererScope();
        var result = renderer.Renderer.Render(new DocumentRenderRequest
        {
            SourcePath = WriteTempFile(".txt", Encoding.UTF8.GetBytes("not a document")),
            Format = DocumentOutputFormat.Pdf
        });

        Assert.False(result.Success);
        Assert.False(string.IsNullOrWhiteSpace(result.ErrorMessage));
        Assert.Contains("Unsupported source type", result.ErrorMessage);
        Assert.Empty(result.Pages);
    }

    [Fact]
    public void Render_MissingFile_ReturnsCleanError()
    {
        using var renderer = new RendererScope();
        var result = renderer.Renderer.Render(new DocumentRenderRequest
        {
            SourcePath = Path.Combine(_tempDirectory, "does-not-exist.pptx"),
            Format = DocumentOutputFormat.Pdf
        });

        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage);
    }

    [Fact]
    public void Render_UnrecognizedJson_ReturnsCleanError()
    {
        using var renderer = new RendererScope();
        var result = renderer.Renderer.Render(new DocumentRenderRequest
        {
            SourcePath = WriteTempFile(".json", Encoding.UTF8.GetBytes("{}")),
            Format = DocumentOutputFormat.Pdf
        });

        Assert.False(result.Success);
        Assert.Contains("generation vocabulary", result.ErrorMessage);
    }

    [Fact]
    public void Render_MalformedJson_ReturnsCleanError()
    {
        using var renderer = new RendererScope();
        var result = renderer.Renderer.Render(new DocumentRenderRequest
        {
            SourcePath = WriteTempFile(".json", Encoding.UTF8.GetBytes("{ not valid json !")),
            Format = DocumentOutputFormat.Pdf
        });

        Assert.False(result.Success);
        Assert.Contains("Malformed JSON", result.ErrorMessage);
    }

    // ─── Fixtures ───────────────────────────────────────────────────

    private const string XlsxJson = """
        {
          "version": "1.0",
          "worksheets": [
            { "name": "Sheet1", "headers": ["A", "B"], "rows": [["1", "2"]] }
          ]
        }
        """;

    private const string DocxJson = """
        {
          "version": "1.0",
          "sections": [
            { "blocks": [ { "type": "paragraph", "text": "Hello facade" } ] }
          ]
        }
        """;

    private const string PptxJson = """
        {
          "version": "2.0",
          "design": { "palette": { "primary": "#0B3D91", "paper": "#FFFFFF" } },
          "slides": [
            {
              "type": "container",
              "children": [
                { "type": "text", "text": "Hello facade", "fontSize": 30, "color": "primary",
                  "at": { "x": 10, "y": 10 }, "size": { "w": 400, "h": 40 } }
              ]
            }
          ]
        }
        """;

    private static byte[] CreatePptxBytes(int slideCount = 1)
    {
        using var builder = PresentationBuilder.Create();
        for (var i = 0; i < slideCount; i++)
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTitle($"Slide {i + 1}");
        }

        return builder.SaveToBytes();
    }

    private static byte[] CreateDocxBytes()
    {
        using var builder = DocumentBuilder.Create();
        builder.AddParagraph("Hello facade document", "Normal");
        return builder.SaveToBytes();
    }

    private static byte[] CreateXlsxBytes()
    {
        using var builder = WorkbookBuilder.Create();
        var sheet = builder.AddWorksheet("Facade");
        sheet.AddHeaderRow(["Region", "Sales"], rowIndex: 1);
        sheet.AddDataRow(["North", "100"], rowIndex: 2);
        sheet.AddDataRow(["South", "200"], rowIndex: 3);
        return builder.SaveToBytes();
    }

    private string WriteTempFile(string extension, byte[] content)
    {
        var path = Path.Combine(_tempDirectory, $"{Guid.NewGuid():N}{extension}");
        File.WriteAllBytes(path, content);
        return path;
    }

    private static void AssertPdfSignature(byte[] bytes)
    {
        Assert.True(bytes.Length > 8, "PDF page is suspiciously small.");
        Assert.Equal(0x25, bytes[0]); // '%'
        Assert.Equal((byte)'P', bytes[1]);
        Assert.Equal((byte)'D', bytes[2]);
        Assert.Equal((byte)'F', bytes[3]);
    }

    private static void AssertPngSignature(byte[] bytes)
    {
        Assert.True(bytes.Length > 8, "PNG page is suspiciously small.");
        Assert.Equal(0x89, bytes[0]);
        Assert.Equal((byte)'P', bytes[1]);
        Assert.Equal((byte)'N', bytes[2]);
        Assert.Equal((byte)'G', bytes[3]);
    }

    private static void AssertTypstSource(byte[] bytes)
    {
        var source = Encoding.UTF8.GetString(bytes);
        Assert.False(string.IsNullOrWhiteSpace(source));
        // Typst sources are markup; a "#set" directive is the universal opener.
        Assert.Contains("#set", source);
    }

    /// <summary>
    /// Owns the shared facade and forces the format-core assemblies to be loaded before the
    /// discovery-based <see cref="DocumentRenderer"/> is constructed (discovery also loads them
    /// explicitly, this is belt-and-braces).
    /// </summary>
    private sealed class RendererScope : IDisposable
    {
        public RendererScope()
        {
            // Touch one type per format core so their assemblies are guaranteed resident.
            _ = typeof(PresentationBuilder);
            _ = typeof(DocumentBuilder);
            _ = typeof(WorkbookBuilder);
            Renderer = new DocumentRenderer();
        }

        public IDocumentRenderer Renderer { get; }

        public void Dispose()
        {
        }
    }
}
