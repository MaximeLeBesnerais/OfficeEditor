using System.Text;
using System.Text.Json;
using OfficeEditor.Api.Services;
using PptxEditor.Core.Builders;
using Xunit.Abstractions;

namespace OfficeEditor.Api.Tests.Unit;

/// <summary>
/// deck generation surface (P9, plan.md §7.1): JSON in → PPTX + per-slide previews out.
/// Preview rendering needs a Typst backend (TypstBridge/CLI), which this sandbox lacks, so
/// the render assertions are environment-invariant (previews xor previewError) and the
/// happy-path render check is opt-in via OE_RUN_TYPST_COMPILE_TESTS=1 (same convention as
/// TypstEmitterCompileTests). Everything else — validation verbatim, PPTX bytes, timing —
/// is fully deterministic.
/// </summary>
public sealed class DeckGenerationServiceTests
{
    private const string EnableRenderEnvVar = "OE_RUN_TYPST_COMPILE_TESTS";

    private const string ValidDocument = """
        {
          "version": "2.0",
          "design": {
            "palette": { "primary": "#0B3D91", "accent": "#FF6B00", "ink": "#1A1A1A", "paper": "#FFFFFF", "muted": "#8A94A6" },
            "fonts": { "display": "Aptos Display", "body": "Aptos" },
            "metrics": { "marginPt": 43, "gutterPt": 18, "titleSizePt": 30, "bodySizePt": 14 }
          },
          "slides": [
            {
              "type": "container",
              "layout": { "mode": "row", "gap": 18, "justify": "space-evenly", "align": "center" },
              "padding": 43,
              "fill": "paper",
              "children": [
                { "type": "card", "size": { "grow": 1, "aspect": "4:3" },
                  "content": { "title": "+34%", "subtitle": "Revenue" } },
                { "type": "card", "size": { "grow": 1, "aspect": "4:3" },
                  "content": { "title": "12k", "subtitle": "Users" } }
              ]
            },
            {
              "type": "container",
              "fill": "primary",
              "children": [
                { "type": "text", "text": "Slide two", "color": "paper", "fontSize": 30,
                  "anchor": "middle", "textAlign": "center",
                  "at": { "x": 0, "y": 250 }, "size": { "w": 960, "h": 40 } }
              ]
            }
          ]
        }
        """;

    private readonly ITestOutputHelper _output;

    public DeckGenerationServiceTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Fact]
    public void Generate_ValidDocument_ReturnsOpenablePptx()
    {
        var result = new DeckGenerationService().Generate(ValidDocument, "svg", 150);

        Assert.True(result.Success);
        Assert.Equal(2, result.SlideCount);
        Assert.Empty(result.Errors);
        Assert.NotNull(result.PptxBytes);
        Assert.True(result.PptxBytes!.Length > 0);
        // Zip magic: the PPTX is a real package, not an error payload.
        Assert.Equal((byte)'P', result.PptxBytes[0]);
        Assert.Equal((byte)'K', result.PptxBytes[1]);

        using var builder = PresentationBuilder.Open(result.PptxBytes);
        Assert.Equal(2, builder.SlideCount);
    }

    [Fact]
    public void Generate_ValidDocument_PreviewContractIsInvariant()
    {
        var result = new DeckGenerationService().Generate(ValidDocument, "svg", 150);

        Assert.True(result.Success);
        if (result.PreviewError is null)
        {
            // Backend available: every slide rendered, content types correct.
            Assert.Equal(result.SlideCount, result.Previews.Count);
            Assert.All(result.Previews, p =>
            {
                Assert.Equal("image/svg+xml", p.ContentType);
                Assert.Equal("svg", p.Format);
                Assert.True(p.Bytes.Length > 0);
            });
            Assert.Equal(Enumerable.Range(1, result.SlideCount), result.Previews.Select(p => p.Slide));
        }
        else
        {
            // No Typst backend (sandbox): explicit error, empty previews, deck still delivered.
            Assert.Empty(result.Previews);
            Assert.NotEmpty(result.PreviewError);
        }
    }

    [Fact]
    public void Generate_ValidDocument_PngFormatPreviewsUsePngContentType()
    {
        var result = new DeckGenerationService().Generate(ValidDocument, "png", 96);

        Assert.True(result.Success);
        if (result.PreviewError is null)
        {
            Assert.All(result.Previews, p => Assert.Equal("image/png", p.ContentType));
        }
    }

    [Fact]
    public void Generate_UnknownProperty_SurfacesValidatorErrorVerbatim()
    {
        var document = ValidDocument.Replace("\"text\": \"Slide two\"", "\"text\": \"Slide two\", \"colour\": \"paper\"");

        var result = new DeckGenerationService().Generate(document, "svg", 150);

        Assert.False(result.Success);
        Assert.Null(result.PptxBytes);
        var error = Assert.Single(result.Errors);
        Assert.Equal("$.slides[1].children[0].colour", error.Path);
        Assert.Contains("unknown property 'colour'", error.Message);
        Assert.Equal("Did you mean 'color'?", error.Suggestion);
    }

    [Fact]
    public void Generate_CssIsm_IsRejectedWithActionableError()
    {
        var document = """
            {
              "version": "2.0",
              "design": { "palette": { "primary": "#0B3D91" } },
              "slides": [
                { "type": "container", "flexWrap": "wrap", "children": [] }
              ]
            }
            """;

        var result = new DeckGenerationService().Generate(document, "svg", 150);

        Assert.False(result.Success);
        var error = Assert.Single(result.Errors);
        Assert.Equal("$.slides[0].flexWrap", error.Path);
        Assert.Contains("CSS 'flex-wrap' is not supported", error.Message);
    }

    [Fact]
    public void Generate_MalformedJson_ReturnsValidatorError()
    {
        var result = new DeckGenerationService().Generate("{ not json", "svg", 150);

        Assert.False(result.Success);
        var error = Assert.Single(result.Errors);
        Assert.Equal("$", error.Path);
        Assert.Contains("malformed JSON", error.Message);
    }

    [Fact]
    public void Generate_ComponentContentError_SurfacesPathAndSuggestion()
    {
        var document = """
            {
              "version": "2.0",
              "design": {
                "palette": { "primary": "#0B3D91", "accent": "#FF6B00", "ink": "#1A1A1A", "paper": "#FFFFFF", "muted": "#8A94A6" }
              },
              "slides": [
                {
                  "type": "container",
                  "layout": { "mode": "row" },
                  "children": [
                    { "type": "kpi", "size": { "grow": 1 },
                      "content": { "value": "+34%", "lable": "Revenue" } }
                  ]
                }
              ]
            }
            """;

        var result = new DeckGenerationService().Generate(document, "svg", 150);

        Assert.False(result.Success);
        var error = Assert.Single(result.Errors);
        Assert.Equal("slides[0].children[0]", error.Path);
        Assert.Contains("unknown property 'lable'", error.Message);
        Assert.Contains("Did you mean 'label'?", error.Message);
    }

    [Fact]
    public void Generate_LayoutOverflowError_SurfacesAsErrorWithPath()
    {
        // Two 600pt rects in a 960pt-wide row, default overflow=error (plan.md §3.2).
        var document = """
            {
              "version": "2.0",
              "design": { "palette": { "primary": "#0B3D91" } },
              "slides": [
                {
                  "type": "container",
                  "layout": { "mode": "row" },
                  "children": [
                    { "type": "rect", "size": { "w": 600, "h": 100 }, "fill": "primary" },
                    { "type": "rect", "size": { "w": 600, "h": 100 }, "fill": "primary" }
                  ]
                }
              ]
            }
            """;

        var result = new DeckGenerationService().Generate(document, "svg", 150);

        Assert.False(result.Success);
        var error = Assert.Single(result.Errors);
        Assert.Equal("slides[0]", error.Path);
        Assert.Contains("overflow", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Generate_ArchetypeSlide_ExpandsBeforeComponentPass()
    {
        // Archetypes (table_slide, two_col, …) are a documented layer of the vocabulary:
        // the pipeline must run archetype expansion before component expansion, like the CLI.
        var document = """
            {
              "version": "2.0",
              "design": {
                "palette": { "primary": "#0B3D91", "accent": "#FF6B00", "ink": "#1A1A1A", "paper": "#FFFFFF", "muted": "#8A94A6" },
                "fonts": { "display": "Aptos Display", "body": "Aptos" },
                "metrics": { "marginPt": 43, "gutterPt": 18, "titleSizePt": 30, "bodySizePt": 14 }
              },
              "slides": [
                {
                  "type": "table_slide",
                  "content": {
                    "title": "Regional performance",
                    "columns": ["Region", "Revenue", "Growth", "Pipeline"],
                    "rows": [
                      ["North America", "$8.2M", "+38%", "$5.1M"],
                      ["EMEA", "$4.6M", "+41%", "$3.8M"]
                    ],
                    "columnWeights": [2, 1, 1, 1]
                  }
                }
              ]
            }
            """;

        var result = new DeckGenerationService().Generate(document, "svg", 150);

        Assert.True(result.Success);
        Assert.Equal(1, result.SlideCount);
        Assert.Empty(result.Errors);
        Assert.NotNull(result.PptxBytes);

        using var builder = PresentationBuilder.Open(result.PptxBytes);
        Assert.Equal(1, builder.SlideCount);
    }

    [Fact]
    public void Generate_ArchetypeContentError_SurfacesAsRejectedResult()
    {
        // Archetype expansion throws ComponentException — it must map to the same rejected
        // result path as component errors, not escape.
        var document = """
            {
              "version": "2.0",
              "design": { "palette": { "primary": "#0B3D91" } },
              "slides": [
                {
                  "type": "table_slide",
                  "content": { "columns": ["A", "B"] }
                }
              ]
            }
            """;

        var result = new DeckGenerationService().Generate(document, "svg", 150);

        Assert.False(result.Success);
        Assert.Null(result.PptxBytes);
        var error = Assert.Single(result.Errors);
        Assert.Equal("slides[0]", error.Path);
        Assert.Contains("rows", error.Message);
    }

    [Fact]
    public void Generate_ImageSlideWithRepoRelativeSrc_GeneratesPptxAndPreviews()
    {
        // Regression test: the src is repo-root-relative and the file ships with the repo.
        // The OOXML pass resolves it against the repository root; the Typst preview compile
        // uses the repository root as its project root (CompileOptions.WorkingDirectory), so
        // the same relative path resolves there too — previously the preview compile failed
        // with "file not found (searched at <api-cwd>/<absolute-src>)".
        var document = """
            {
              "version": "2.0",
              "design": {
                "palette": { "primary": "#0B3D91", "ink": "#1A1A1A", "paper": "#FFFFFF", "muted": "#8A94A6" }
              },
              "slides": [
                {
                  "type": "container",
                  "fill": "paper",
                  "layout": { "mode": "column" },
                  "padding": 43,
                  "children": [
                    { "type": "image", "src": "demo/assets/dashboard.png", "fit": "contain", "size": { "grow": 1 } }
                  ]
                }
              ]
            }
            """;

        var result = new DeckGenerationService().Generate(document, "svg", 150);

        Assert.True(result.Success,
            $"generation failed: {string.Join("; ", result.Errors.Select(e => $"{e.Path}: {e.Message}"))}");
        Assert.Equal(1, result.SlideCount);
        Assert.NotNull(result.PptxBytes);
        Assert.True(result.PptxBytes!.Length > 0);

        using var builder = PresentationBuilder.Open(result.PptxBytes);
        Assert.Equal(1, builder.SlideCount);

        if (Environment.GetEnvironmentVariable(EnableRenderEnvVar) == "1")
        {
            Assert.Null(result.PreviewError);
            Assert.Equal(result.SlideCount, result.Previews.Count);
        }
    }

    [Fact]
    public void Generate_InvalidFormat_Throws()
    {
        Assert.Throws<ArgumentException>(
            () => new DeckGenerationService().Generate(ValidDocument, "jpeg", 150));
    }

    [Fact]
    public void Generate_ImageSlideWithTraversalSrc_IsRejected()
    {
        // The generate endpoint is unauthenticated: a "../../…" src must never resolve
        // outside the repository and be embedded into the returned PPTX.
        var document = ImageDocument("../../../../../../etc/passwd");

        var result = new DeckGenerationService().Generate(document, "svg", 150);

        Assert.False(result.Success);
        Assert.Null(result.PptxBytes);
        var error = Assert.Single(result.Errors);
        Assert.Contains("outside the allowed root", error.Message);
    }

    [Fact]
    public void Generate_ImageSlideWithAbsoluteSrcOutsideRepoRoot_IsRejected()
    {
        // A real file outside the repository: without the containment check it would be
        // embedded into the returned PPTX byte-for-byte.
        var outsideDir = Path.Combine(Path.GetTempPath(), $"deck-generation-tests-{Guid.NewGuid():N}");
        Directory.CreateDirectory(outsideDir);
        try
        {
            var outsidePath = Path.Combine(outsideDir, "secret.png");
            File.WriteAllBytes(outsidePath, [0x89, 0x50, 0x4E, 0x47]); // PNG magic

            var result = new DeckGenerationService().Generate(ImageDocument(outsidePath), "svg", 150);

            Assert.False(result.Success);
            Assert.Null(result.PptxBytes);
            var error = Assert.Single(result.Errors);
            Assert.Contains("outside the allowed root", error.Message);
        }
        finally
        {
            Directory.Delete(outsideDir, recursive: true);
        }
    }

    private static string ImageDocument(string src) => $$"""
        {
          "version": "2.0",
          "design": {
            "palette": { "primary": "#0B3D91", "ink": "#1A1A1A", "paper": "#FFFFFF", "muted": "#8A94A6" }
          },
          "slides": [
            {
              "type": "container",
              "fill": "paper",
              "layout": { "mode": "column" },
              "padding": 43,
              "children": [
                { "type": "image", "src": {{JsonSerializer.Serialize(src)}}, "fit": "contain", "size": { "grow": 1 } }
              ]
            }
          ]
        }
        """;

    [Fact]
    public void Generate_SixteenSlides_WarmPathUnderTarget()
    {
        // plan.md §7.1: warm < 500ms/deck for ≤16 slides (PPTX path; preview render excluded
        // — it is backend-dependent and absent in this sandbox).
        var document = BuildDocument(slideCount: 16);
        var service = new DeckGenerationService();

        var cold = service.Generate(document, "svg", 150);
        Assert.True(cold.Success);

        var warm = service.Generate(document, "svg", 150);
        Assert.True(warm.Success);
        Assert.Equal(16, warm.SlideCount);

        _output.WriteLine(
            $"16-slide deck: cold={cold.GenerationMilliseconds:F1}ms warm={warm.GenerationMilliseconds:F1}ms " +
            $"(total incl. preview attempt: cold={cold.TotalMilliseconds:F1}ms warm={warm.TotalMilliseconds:F1}ms)");
        Assert.True(
            warm.GenerationMilliseconds < 500,
            $"warm generation took {warm.GenerationMilliseconds:F1}ms, target < 500ms");
    }

    [Fact]
    public void Generate_Previews_RenderedPerSlide_WhenBackendAvailable()
    {
        if (Environment.GetEnvironmentVariable(EnableRenderEnvVar) != "1")
        {
            return; // no Typst backend in this environment (see class summary)
        }

        var result = new DeckGenerationService().Generate(ValidDocument, "svg", 150);

        Assert.True(result.Success);
        Assert.Null(result.PreviewError);
        Assert.Equal(result.SlideCount, result.Previews.Count);
        Assert.All(result.Previews, p =>
        {
            var svg = Encoding.UTF8.GetString(p.Bytes);
            Assert.Contains("<svg", svg);
        });
    }

    private static string BuildDocument(int slideCount)
    {
        var sb = new StringBuilder();
        sb.Append("""
            {
              "version": "2.0",
              "design": {
                "palette": { "primary": "#0B3D91", "accent": "#FF6B00", "ink": "#1A1A1A", "paper": "#FFFFFF", "muted": "#8A94A6" },
                "fonts": { "display": "Aptos Display", "body": "Aptos" },
                "metrics": { "marginPt": 43, "gutterPt": 18, "titleSizePt": 30, "bodySizePt": 14 }
              },
              "slides": [
            """);
        for (var i = 0; i < slideCount; i++)
        {
            if (i > 0)
            {
                sb.Append(',');
            }
            sb.Append($$"""

                        {
                          "type": "container",
                          "padding": 43,
                          "fill": "paper",
                          "layout": { "mode": "column", "gap": 12 },
                          "children": [
                            { "type": "title_block", "size": { "h": 84 },
                              "content": { "title": "Slide {{i + 1}}", "kicker": "SECTION", "subtitle": "Generated" } },
                            { "type": "divider", "content": {} },
                            {
                              "type": "container",
                              "layout": { "mode": "row", "gap": 18 },
                              "size": { "grow": 3 },
                              "children": [
                                { "type": "kpi", "size": { "grow": 1 }, "content": { "value": "+34%", "label": "Revenue", "delta": "+12% vs LY" } },
                                { "type": "kpi", "size": { "grow": 1 }, "content": { "value": "12k", "label": "Users" } },
                                { "type": "kpi", "size": { "grow": 1 }, "content": { "value": "98.5%", "label": "Uptime" } }
                              ]
                            },
                            { "type": "bullet_list", "size": { "h": 76 },
                              "content": { "items": ["First point", "Second point", "Third point"] } }
                          ]
                        }
            """);
        }
        sb.Append("\n  ]\n}\n");
        return sb.ToString();
    }
}
