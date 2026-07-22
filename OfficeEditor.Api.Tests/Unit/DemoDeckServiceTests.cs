using OfficeEditor.Api.Services;
using OfficeEditor.Core.Services;
using PptxEditor.Core.Builders;

namespace OfficeEditor.Api.Tests.Unit;

/// <summary>
/// Demo deck catalog + timed render service. The whole-deck render needs a Typst backend,
/// so the happy-path render test is opt-in via OE_RUN_TYPST_COMPILE_TESTS=1 (same
/// convention as DeckGenerationServiceTests); everything else is environment-invariant.
/// </summary>
public sealed class DemoDeckServiceTests
{
    private const string EnableRenderEnvVar = "OE_RUN_TYPST_COMPILE_TESTS";

    [Fact]
    public void ListDecks_ReturnsExactlyTheWhitelistedNamesInOrder()
    {
        var service = new DemoDeckService(new StubDeckSessionStore());

        var decks = service.ListDecks();

        Assert.Equal(
            ["pres-pro", "aetherlink", "fusionfest", "pitch-deck"],
            decks.Select(d => d.Name));
        Assert.All(decks, d =>
        {
            Assert.False(string.IsNullOrWhiteSpace(d.FileName));
            Assert.False(string.IsNullOrWhiteSpace(d.Description));
            Assert.True(d.SlideCount >= 0);
        });
    }

    [Fact]
    public void ListDecks_ExistingRefFiles_ReportPositiveSlideCounts()
    {
        var service = new DemoDeckService(new StubDeckSessionStore());

        var decks = service.ListDecks();

        // All four REF decks ship with the repo; a 0 would mean the lazy open failed.
        Assert.All(decks, d => Assert.True(d.SlideCount > 0, $"{d.Name} reported 0 slides"));
    }

    [Fact]
    public void RenderDeck_UnknownName_ThrowsArgumentExceptionListingValidNames()
    {
        var service = new DemoDeckService(new StubDeckSessionStore());

        var ex = Assert.Throws<ArgumentException>(() => service.RenderDeck("no-such-deck", 110, "png"));

        Assert.Contains("pres-pro", ex.Message);
        Assert.Contains("aetherlink", ex.Message);
        Assert.Contains("fusionfest", ex.Message);
        Assert.Contains("pitch-deck", ex.Message);
    }

    [Theory]
    [InlineData("jpeg")]
    [InlineData("SVG")]
    [InlineData(" png ")]
    public void RenderDeck_NonNormalizedFormat_ThrowsArgumentException(string format)
    {
        // The endpoint normalizes via DeckPreviewValidators.TryNormalizeFormat; the service
        // requires already-normalized "svg"/"png" input (same defensive contract as
        // DeckGenerationService.Generate) and rejects anything else before rendering.
        var service = new DemoDeckService(new StubDeckSessionStore());

        var ex = Assert.Throws<ArgumentException>(() => service.RenderDeck("pres-pro", 110, format));

        Assert.Contains("normalized", ex.Message);
        Assert.Equal("format", ex.ParamName);
    }

    [Fact]
    public void TryGetDeckFile_KnownName_ResolvesToExistingFile()
    {
        var service = new DemoDeckService(new StubDeckSessionStore());

        Assert.True(service.TryGetDeckFile("pres-pro", out var path));
        Assert.True(File.Exists(path));
        Assert.EndsWith("pres-pro.pptx", path);
    }

    [Theory]
    [InlineData("no-such-deck")]
    [InlineData("")]
    [InlineData("   ")]
    public void TryGetDeckFile_UnknownOrBlankName_ReturnsFalse(string name)
    {
        var service = new DemoDeckService(new StubDeckSessionStore());

        Assert.False(service.TryGetDeckFile(name, out _));
    }

    [Fact]
    public void RewriteRelativeSrcPaths_RelativeSrc_BecomesRepoRootRelative()
    {
        const string json = """
            {
              "slides": [
                { "type": "image", "src": "demo/assets/dashboard.png" },
                { "type": "container", "children": [ { "type": "image", "src": "assets/nested.jpg" } ] }
              ]
            }
            """;

        var rewritten = DemoDeckService.RewriteRelativeSrcPaths(json);

        // Repo-root-relative forward-slash paths, NOT absolutized: the Typst preview
        // compile resolves them against the repository root (its project root).
        Assert.Contains("\"src\": \"demo/assets/dashboard.png\"", rewritten);
        Assert.Contains("\"src\": \"assets/nested.jpg\"", rewritten);
    }

    [Fact]
    public void RewriteRelativeSrcPaths_DotSlashAndBackslashes_AreNormalized()
    {
        const string json = """
            { "type": "image", "src": "./demo\\assets\\dashboard.png" }
            """;

        var rewritten = DemoDeckService.RewriteRelativeSrcPaths(json);

        Assert.Contains("\"src\": \"demo/assets/dashboard.png\"", rewritten);
    }

    [Fact]
    public void RewriteRelativeSrcPaths_AbsoluteSrc_IsLeftUntouched()
    {
        const string json = """
            { "type": "image", "src": "/already/absolute/logo.png" }
            """;

        var rewritten = DemoDeckService.RewriteRelativeSrcPaths(json);

        Assert.Contains("\"src\": \"/already/absolute/logo.png\"", rewritten);
    }

    [Fact]
    public void RewriteRelativeSrcPaths_SrcInsideArrays_IsRewritten()
    {
        const string json = """
            { "items": [ { "src": "a.png" }, { "src": "b.png" } ], "other": "src" }
            """;

        var rewritten = DemoDeckService.RewriteRelativeSrcPaths(json);

        Assert.Contains("\"src\": \"a.png\"", rewritten);
        Assert.Contains("\"src\": \"b.png\"", rewritten);
        // Only properties NAMED "src" are rewritten; a "src" value elsewhere is untouched.
        Assert.Contains("\"other\": \"src\"", rewritten);
    }

    [Fact]
    public void DeckGenerationService_FontDirectory_IsExposedForPreviewCompile()
    {
        Assert.Null(new DeckGenerationService().FontDirectory);
        Assert.Null(new DeckGenerationService("").FontDirectory);
        Assert.Null(new DeckGenerationService("   ").FontDirectory);
        Assert.Equal("/System/Library/Fonts:/Library/Fonts",
            new DeckGenerationService("/System/Library/Fonts:/Library/Fonts").FontDirectory);
        Assert.Equal("/fonts",
            new DeckGenerationService(new TypstCompilerService(), "/fonts").FontDirectory);
    }

    [Fact]
    public void RenderDeck_PresPro_RendersAllSlidesAndStoresSession()
    {
        if (Environment.GetEnvironmentVariable(EnableRenderEnvVar) != "1")
        {
            return; // no Typst backend in this environment (see class summary)
        }

        var store = new StubDeckSessionStore();
        var service = new DemoDeckService(store);

        var result = service.RenderDeck("pres-pro", 110, "png");

        Assert.True(result.SlideCount > 0);
        Assert.Equal("png", result.Format);
        Assert.Equal(result.SlideCount, result.Pages.Count);
        Assert.True(result.TotalMilliseconds > 0);
        Assert.All(result.Pages, page =>
        {
            // PNG magic: 0x89 'P' 'N' 'G'
            Assert.True(page.Length > 8);
            Assert.Equal(0x89, page[0]);
            Assert.Equal((byte)'P', page[1]);
            Assert.Equal((byte)'N', page[2]);
            Assert.Equal((byte)'G', page[3]);
        });

        Assert.True(store.TryGet(result.DeckId, out var session));
        Assert.Equal(result.SlideCount, session!.SlideCount);
        Assert.Equal("pres-pro.pptx", session.FileName);
    }

    [Fact]
    public void RenderDeck_PresPro_SvgFormat_RendersSvgPages()
    {
        if (Environment.GetEnvironmentVariable(EnableRenderEnvVar) != "1")
        {
            return; // no Typst backend in this environment (see class summary)
        }

        var store = new StubDeckSessionStore();
        var service = new DemoDeckService(store);

        var result = service.RenderDeck("pres-pro", 110, "svg");

        Assert.True(result.SlideCount > 0);
        Assert.Equal("svg", result.Format);
        Assert.Equal(result.SlideCount, result.Pages.Count);
        Assert.All(result.Pages, page =>
        {
            var text = System.Text.Encoding.UTF8.GetString(page);
            Assert.Contains("<svg", text);
        });

        Assert.True(store.TryGet(result.DeckId, out var session));
        Assert.Equal(result.SlideCount, session!.SlideCount);
    }

    [Fact]
    public void RenderUploadedDeck_GarbageBytes_ThrowsArgumentExceptionMentioningValidPptx()
    {
        var service = new DemoDeckService(new StubDeckSessionStore());
        var garbage = System.Text.Encoding.UTF8.GetBytes("this is definitely not a pptx file");

        var ex = Assert.Throws<ArgumentException>(() =>
            service.RenderUploadedDeck(garbage, "garbage.pptx", "png", 110));

        Assert.Contains("valid PPTX", ex.Message);
    }

    [Fact]
    public void RenderUploadedDeck_DeckOverSlideCap_ThrowsArgumentExceptionMentioning60()
    {
        var service = new DemoDeckService(new StubDeckSessionStore());
        var deckBytes = BuildDeckBytes(DemoDeckService.MaxUploadSlides + 1);

        var ex = Assert.Throws<ArgumentException>(() =>
            service.RenderUploadedDeck(deckBytes, "big-deck.pptx", "png", 110));

        Assert.Contains("60", ex.Message);
    }

    [Theory]
    [InlineData("jpeg")]
    [InlineData("SVG")]
    [InlineData(" png ")]
    public void RenderUploadedDeck_NonNormalizedFormat_ThrowsArgumentException(string format)
    {
        // Same defensive contract as RenderDeck: the endpoint normalizes first; the
        // service requires already-normalized "svg"/"png" input.
        var service = new DemoDeckService(new StubDeckSessionStore());

        var ex = Assert.Throws<ArgumentException>(() =>
            service.RenderUploadedDeck(BuildDeckBytes(1), "deck.pptx", format, 110));

        Assert.Contains("normalized", ex.Message);
        Assert.Equal("format", ex.ParamName);
    }

    [Fact]
    public void RenderUploadedDeck_InMemoryDeck_RendersAllSlidesAndStoresSession()
    {
        if (Environment.GetEnvironmentVariable(EnableRenderEnvVar) != "1")
        {
            return; // no Typst backend in this environment (see class summary)
        }

        var store = new StubDeckSessionStore();
        var service = new DemoDeckService(store);

        var result = service.RenderUploadedDeck(BuildDeckBytes(3), "upload.pptx", "png", 110);

        Assert.Equal(3, result.SlideCount);
        Assert.Equal("png", result.Format);
        Assert.Equal(3, result.Pages.Count);
        Assert.True(result.TotalMilliseconds > 0);
        Assert.All(result.Pages, page =>
        {
            // PNG magic: 0x89 'P' 'N' 'G'
            Assert.True(page.Length > 8);
            Assert.Equal(0x89, page[0]);
            Assert.Equal((byte)'P', page[1]);
        });

        Assert.True(store.TryGet(result.DeckId, out var session));
        Assert.Equal(3, session!.SlideCount);
        Assert.Equal("upload.pptx", session.FileName);
    }

    [Fact]
    public void RenderTypstLeg_GarbageBytes_ThrowsArgumentExceptionMentioningValidPptx()
    {
        var service = new DemoDeckService(new StubDeckSessionStore());
        var garbage = System.Text.Encoding.UTF8.GetBytes("this is definitely not a pptx file");

        var ex = Assert.Throws<ArgumentException>(() =>
            service.RenderTypstLeg(garbage, "garbage.pptx", 110));

        Assert.Contains("valid PPTX", ex.Message);
    }

    [Fact]
    public void RenderTypstLeg_EmptyDeck_ThrowsArgumentExceptionMentioningNoSlides()
    {
        var service = new DemoDeckService(new StubDeckSessionStore());

        var ex = Assert.Throws<ArgumentException>(() =>
            service.RenderTypstLeg(BuildDeckBytes(0), "empty.pptx", 110));

        Assert.Contains("no slides", ex.Message);
    }

    [Fact]
    public void RenderTypstLeg_DeckOverSlideCap_ThrowsArgumentExceptionMentioning60()
    {
        var service = new DemoDeckService(new StubDeckSessionStore());
        var deckBytes = BuildDeckBytes(DemoDeckService.MaxUploadSlides + 1);

        var ex = Assert.Throws<ArgumentException>(() =>
            service.RenderTypstLeg(deckBytes, "big-deck.pptx", 110));

        Assert.Contains("60", ex.Message);
    }

    [Fact]
    public void RenderTypstLeg_InMemoryDeck_RendersPngPagesAndBestEffortPdf()
    {
        if (Environment.GetEnvironmentVariable(EnableRenderEnvVar) != "1")
        {
            return; // no Typst backend in this environment (see class summary)
        }

        var service = new DemoDeckService(new StubDeckSessionStore());

        var result = service.RenderTypstLeg(BuildDeckBytes(3), "upload.pptx", 110);

        Assert.Equal(3, result.SlideCount);
        Assert.True(result.PngMilliseconds > 0);
        Assert.Equal(3, result.PngPages.Count);
        Assert.All(result.PngPages, page =>
        {
            // PNG magic: 0x89 'P' 'N' 'G'
            Assert.True(page.Length > 8);
            Assert.Equal(0x89, page[0]);
            Assert.Equal((byte)'P', page[1]);
        });

        // The PDF leg is best-effort: exactly one of PdfBytes / PdfError must be set.
        Assert.True(result.PdfBytes is not null || result.PdfError is not null,
            "PDF leg produced neither bytes nor an error");
        Assert.False(result.PdfBytes is not null && result.PdfError is not null,
            "PDF leg produced both bytes and an error");
        if (result.PdfBytes is not null)
        {
            // PDF magic: "%PDF-"
            Assert.True(result.PdfMilliseconds > 0);
            var header = System.Text.Encoding.ASCII.GetString(result.PdfBytes, 0, 5);
            Assert.Equal("%PDF-", header);
        }

        // Wall-clock semantics: the PNG and PDF legs run concurrently, so the total must
        // cover the slower leg (small epsilon for scheduling slop) and must not exceed the
        // sequential sum of the per-phase times.
        const double epsilonMs = 50;
        var slowestPhase = Math.Max(result.PngMilliseconds, result.PdfMilliseconds ?? 0);
        var sequentialSum = result.PngMilliseconds + (result.PdfMilliseconds ?? 0);
        Assert.True(result.TotalMilliseconds >= slowestPhase - epsilonMs,
            $"Total {result.TotalMilliseconds:F1}ms < slowest phase {slowestPhase:F1}ms");
        Assert.True(result.TotalMilliseconds <= sequentialSum + epsilonMs,
            $"Total {result.TotalMilliseconds:F1}ms > sequential sum {sequentialSum:F1}ms (legs did not overlap)");
    }

    private static byte[] BuildDeckBytes(int slideCount)
    {
        using var builder = PresentationBuilder.Create();
        for (var i = 0; i < slideCount; i++)
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTitle($"Slide {i + 1}");
        }

        return builder.SaveToBytes();
    }
}
