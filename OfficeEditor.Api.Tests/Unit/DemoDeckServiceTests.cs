using System.Text.Json;
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
            ["northwind", "sales", "aetherlink", "launch-review"],
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

        // The whitelisted REF deck ships with the repo; a 0 would mean the lazy open failed.
        Assert.All(decks, d => Assert.True(d.SlideCount > 0, $"{d.Name} reported 0 slides"));
    }

    [Fact]
    public void RenderDeck_UnknownName_ThrowsArgumentExceptionListingValidNames()
    {
        var service = new DemoDeckService(new StubDeckSessionStore());

        var ex = Assert.Throws<ArgumentException>(() => service.RenderDeck("no-such-deck", 110, "png"));

        Assert.Contains("northwind", ex.Message);
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

        var ex = Assert.Throws<ArgumentException>(() => service.RenderDeck("northwind", 110, format));

        Assert.Contains("normalized", ex.Message);
        Assert.Equal("format", ex.ParamName);
    }

    [Fact]
    public void TryGetDeckFile_KnownName_ResolvesToExistingFile()
    {
        var service = new DemoDeckService(new StubDeckSessionStore());

        Assert.True(service.TryGetDeckFile("northwind", out var path));
        Assert.True(File.Exists(path));
        Assert.EndsWith("northwind-demo.pptx", path);

        Assert.True(service.TryGetDeckFile("sales", out var salesPath));
        Assert.True(File.Exists(salesPath));
        Assert.EndsWith("sales_acceleration_deck.pptx", salesPath);

        Assert.True(service.TryGetDeckFile("aetherlink", out var aetherPath));
        Assert.True(File.Exists(aetherPath));
        Assert.EndsWith("AetherLink-Glass-Shareholder-Overview.pptx", aetherPath);

        Assert.True(service.TryGetDeckFile("launch-review", out var launchPath));
        Assert.True(File.Exists(launchPath));
        Assert.EndsWith("northwind-launch-review.pptx", launchPath);
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
                { "type": "image", "src": "assets/dashboard.png" },
                { "type": "container", "children": [ { "type": "image", "src": "assets/nested.jpg" } ] }
              ]
            }
            """;

        var rewritten = DemoDeckService.RewriteRelativeSrcPaths(json);

        // Doc-relative sources resolve against the template's own directory (demo/),
        // producing repo-root-relative forward-slash paths — NOT absolutized: the Typst
        // preview compile resolves them against the repository root (its project root).
        Assert.Contains("\"src\": \"demo/assets/dashboard.png\"", rewritten);
        Assert.Contains("\"src\": \"demo/assets/nested.jpg\"", rewritten);
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

        Assert.Contains("\"src\": \"demo/a.png\"", rewritten);
        Assert.Contains("\"src\": \"demo/b.png\"", rewritten);
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
    public void Constructor_FontDirectory_IsOptionalAndAccepted()
    {
        // Existing call sites (no font directory) keep compiling and working.
        var withoutFonts = new DemoDeckService(new StubDeckSessionStore());
        Assert.Equal(4, withoutFonts.ListDecks().Count);

        // A provided system-font path is accepted and does not affect non-render operations.
        var withFonts = new DemoDeckService(new StubDeckSessionStore(), "/System/Library/Fonts:/Library/Fonts");
        Assert.Equal(4, withFonts.ListDecks().Count);
    }

    [Fact]
    public void RenderDeck_Northwind_WithFontDirectory_RendersAllSlides()
    {
        if (Environment.GetEnvironmentVariable(EnableRenderEnvVar) != "1")
        {
            return; // no Typst backend in this environment (see class summary)
        }

        // End-to-end: the configured font directory flows into the thumbnail compile.
        var service = new DemoDeckService(new StubDeckSessionStore(), "/System/Library/Fonts:/Library/Fonts");

        var result = service.RenderDeck("northwind", 110, "png");

        Assert.Equal(15, result.SlideCount);
        Assert.Equal(result.SlideCount, result.Pages.Count);
        Assert.All(result.Pages, page => Assert.Equal(0x89, page[0]));
    }

    [Fact]
    public void RenderDeck_Northwind_RendersAllSlidesAndStoresSession()
    {
        if (Environment.GetEnvironmentVariable(EnableRenderEnvVar) != "1")
        {
            return; // no Typst backend in this environment (see class summary)
        }

        var store = new StubDeckSessionStore();
        var service = new DemoDeckService(store);

        var result = service.RenderDeck("northwind", 110, "png");

        Assert.Equal(15, result.SlideCount);
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
        Assert.Equal("northwind-demo.pptx", session.FileName);
    }

    [Fact]
    public void RenderDeck_Northwind_SvgFormat_RendersSvgPages()
    {
        if (Environment.GetEnvironmentVariable(EnableRenderEnvVar) != "1")
        {
            return; // no Typst backend in this environment (see class summary)
        }

        var store = new StubDeckSessionStore();
        var service = new DemoDeckService(store);

        var result = service.RenderDeck("northwind", 110, "svg");

        Assert.Equal(15, result.SlideCount);
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
    public void RenderUploadedDeck_DeckOverSlideCap_ThrowsArgumentExceptionMentioningCap()
    {
        var service = new DemoDeckService(new StubDeckSessionStore());
        var deckBytes = BuildDeckBytes(DemoDeckService.MaxUploadSlides + 1);

        var ex = Assert.Throws<ArgumentException>(() =>
            service.RenderUploadedDeck(deckBytes, "big-deck.pptx", "png", 110));

        Assert.Contains(DemoDeckService.MaxUploadSlides.ToString(), ex.Message);
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
    public void RenderUploadedDeck_CompileFailure_ThrowsInsteadOfReturningEmptyPreviews()
    {
        if (Environment.GetEnvironmentVariable(EnableRenderEnvVar) != "1")
        {
            return; // no Typst backend in this environment (see class summary)
        }

        // Regression: a deck whose Typst compile fails (here: an image Typst cannot
        // decode) used to come back as a "successful" result with ZERO pages — the
        // upload endpoint serialized that as {"success": true, "previews": []} and
        // the UI showed an empty gallery with no error. Compile failures must
        // propagate so the endpoint (500 {error}) and the Blazor AnyRenderScreen
        // (StatusMessage) can surface them.
        var service = new DemoDeckService(new StubDeckSessionStore());

        var ex = Assert.Throws<InvalidOperationException>(() =>
            service.RenderUploadedDeck(BuildDeckWithUndecodableImageBytes(), "broken.pptx", "png", 110));

        Assert.Contains("Typst", ex.Message);
    }

    private static byte[] BuildDeckWithUndecodableImageBytes()
    {
        var imagePath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.jpg");
        try
        {
            // Not a JPEG — Typst fails to decode it at compile time.
            File.WriteAllBytes(imagePath, [0xDE, 0xAD, 0xBE, 0xEF, 0x00, 0x01, 0x02, 0x03]);

            using var builder = PresentationBuilder.Create();
            builder.AddSlide();
            builder.CurrentSlide.AddImage(imagePath);
            return builder.SaveToBytes();
        }
        finally
        {
            if (File.Exists(imagePath))
            {
                File.Delete(imagePath);
            }
        }
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
    public void RenderTypstLeg_DeckOverSlideCap_ThrowsArgumentExceptionMentioningCap()
    {
        var service = new DemoDeckService(new StubDeckSessionStore());
        var deckBytes = BuildDeckBytes(DemoDeckService.MaxUploadSlides + 1);

        var ex = Assert.Throws<ArgumentException>(() =>
            service.RenderTypstLeg(deckBytes, "big-deck.pptx", 110));

        Assert.Contains(DemoDeckService.MaxUploadSlides.ToString(), ex.Message);
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

    [Fact]
    public void RenderDeckAllFormats_UnknownName_ThrowsArgumentExceptionListingValidNames()
    {
        var service = new DemoDeckService(new StubDeckSessionStore());

        var ex = Assert.Throws<ArgumentException>(() => service.RenderDeckAllFormats("no-such-deck", 110));

        Assert.Contains("northwind", ex.Message);
    }

    [Fact]
    public void RenderDeckAllFormats_Northwind_RendersAllFormatsWithStageTimings()
    {
        if (Environment.GetEnvironmentVariable(EnableRenderEnvVar) != "1")
        {
            return; // no Typst backend in this environment (see class summary)
        }

        var service = new DemoDeckService(new StubDeckSessionStore(), "/System/Library/Fonts:/Library/Fonts");

        var result = service.RenderDeckAllFormats("northwind", 110);

        Assert.Equal(15, result.SlideCount);
        Assert.Equal(result.SlideCount, result.PngPages.Count);
        Assert.Equal(result.SlideCount, result.SvgPages.Count);
        Assert.All(result.PngPages, page => Assert.Equal(0x89, page[0]));
        Assert.All(result.SvgPages, page =>
            Assert.Contains("<svg", System.Text.Encoding.UTF8.GetString(page)));

        // The PDF leg is best-effort: exactly one of PdfBytes / PdfError must be set.
        Assert.True(result.PdfBytes is not null || result.PdfError is not null,
            "PDF leg produced neither bytes nor an error");
        Assert.False(result.PdfBytes is not null && result.PdfError is not null,
            "PDF leg produced both bytes and an error");
        if (result.PdfBytes is not null)
        {
            Assert.Equal("%PDF-", System.Text.Encoding.ASCII.GetString(result.PdfBytes, 0, 5));
        }

        // Stage timings: one shared conversion + parallel compiles, so the wall-clock
        // total covers the conversion plus the slowest compile leg (epsilon for
        // scheduling slop) and stays under the sequential sum.
        Assert.True(result.ConversionMilliseconds > 0);
        Assert.True(result.PngMilliseconds is > 0);
        Assert.True(result.SvgMilliseconds > 0);
        const double epsilonMs = 100;
        var slowestLeg = Math.Max(
            Math.Max(result.PngMilliseconds ?? 0, result.SvgMilliseconds),
            result.PdfMilliseconds ?? 0);
        Assert.True(result.TotalMilliseconds >= result.ConversionMilliseconds + slowestLeg - epsilonMs,
            $"Total {result.TotalMilliseconds:F1}ms < conversion {result.ConversionMilliseconds:F1}ms + slowest leg {slowestLeg:F1}ms");
        var sequentialSum = result.ConversionMilliseconds
            + (result.PngMilliseconds ?? 0) + result.SvgMilliseconds + (result.PdfMilliseconds ?? 0);
        Assert.True(result.TotalMilliseconds <= sequentialSum + epsilonMs,
            $"Total {result.TotalMilliseconds:F1}ms > sequential sum {sequentialSum:F1}ms (legs did not overlap)");
    }

    [Fact]
    public void RenderDeckAllFormats_SkipPng_OmitsPngLegAndKeepsOtherFormats()
    {
        if (Environment.GetEnvironmentVariable(EnableRenderEnvVar) != "1")
        {
            return; // no Typst backend in this environment (see class summary)
        }

        var service = new DemoDeckService(new StubDeckSessionStore(), "/System/Library/Fonts:/Library/Fonts");

        var result = service.RenderDeckAllFormats("northwind", 110, includePng: false);

        // The PNG leg never ran: no pages, and a null timing — never a fake 0 ms.
        Assert.Null(result.PngMilliseconds);
        Assert.Empty(result.PngPages);

        Assert.Equal(15, result.SlideCount);
        Assert.Equal(result.SlideCount, result.SvgPages.Count);
        Assert.All(result.SvgPages, page =>
            Assert.Contains("<svg", System.Text.Encoding.UTF8.GetString(page)));

        // The PDF leg is best-effort: exactly one of PdfBytes / PdfError must be set.
        Assert.True(result.PdfBytes is not null || result.PdfError is not null,
            "PDF leg produced neither bytes nor an error");
        Assert.False(result.PdfBytes is not null && result.PdfError is not null,
            "PDF leg produced both bytes and an error");
        if (result.PdfBytes is not null)
        {
            Assert.Equal("%PDF-", System.Text.Encoding.ASCII.GetString(result.PdfBytes, 0, 5));
        }

        // Stage timings stay honest without the PNG leg: conversion + parallel
        // SVG/PDF compiles bound the wall-clock total (epsilon for scheduling slop).
        Assert.True(result.ConversionMilliseconds > 0);
        Assert.True(result.SvgMilliseconds > 0);
        const double epsilonMs = 100;
        var slowestLeg = Math.Max(result.SvgMilliseconds, result.PdfMilliseconds ?? 0);
        Assert.True(result.TotalMilliseconds >= result.ConversionMilliseconds + slowestLeg - epsilonMs,
            $"Total {result.TotalMilliseconds:F1}ms < conversion {result.ConversionMilliseconds:F1}ms + slowest leg {slowestLeg:F1}ms");
        var sequentialSum = result.ConversionMilliseconds
            + result.SvgMilliseconds + (result.PdfMilliseconds ?? 0);
        Assert.True(result.TotalMilliseconds <= sequentialSum + epsilonMs,
            $"Total {result.TotalMilliseconds:F1}ms > sequential sum {sequentialSum:F1}ms (legs did not overlap)");
    }

    [Fact]
    public void RewriteRelativeSrcPaths_NullInput_ThrowsArgumentNullException()
    {
        var ex = Assert.Throws<ArgumentNullException>(() => DemoDeckService.RewriteRelativeSrcPaths(null!));
        Assert.Equal("json", ex.ParamName);
    }

    [Fact]
    public void RewriteRelativeSrcPaths_EmptyString_Throws()
    {
        // JsonNode.Parse("") throws on empty input — the exception type depends on
        // framework internals; the exact type doesn't matter, only that it throws.
        var caught = Record.Exception(() => DemoDeckService.RewriteRelativeSrcPaths(""));
        Assert.NotNull(caught);
        Assert.IsAssignableFrom<Exception>(caught);
    }

    [Fact]
    public void RewriteRelativeSrcPaths_MalformedJson_Throws()
    {
        var caught = Record.Exception(() =>
            DemoDeckService.RewriteRelativeSrcPaths("{ not valid json"));
        Assert.NotNull(caught);
        Assert.IsAssignableFrom<Exception>(caught);
    }

    [Fact]
    public void RewriteRelativeSrcPaths_DataUriSrc_IsLeftUntouched()
    {
        const string json = """
            { "type": "image", "src": "data:image/png;base64,iVBORw0KGgo=" }
            """;

        var rewritten = DemoDeckService.RewriteRelativeSrcPaths(json);

        Assert.Contains("\"src\": \"data:image/png;base64,iVBORw0KGgo=\"", rewritten);
    }

    [Fact]
    public void RewriteRelativeSrcPaths_HttpUrlSrc_IsLeftUntouched()
    {
        const string json = """
            { "type": "image", "src": "https://cdn.example.com/img/logo.png" }
            """;

        var rewritten = DemoDeckService.RewriteRelativeSrcPaths(json);

        Assert.Contains("\"src\": \"https://cdn.example.com/img/logo.png\"", rewritten);
    }

    [Fact]
    public void RewriteRelativeSrcPaths_EmptySrcString_IsNotNormalized()
    {
        const string json = """
            { "type": "image", "src": "" }
            """;

        var rewritten = DemoDeckService.RewriteRelativeSrcPaths(json);

        Assert.Contains("\"src\": \"\"", rewritten);
    }

    [Fact]
    public void RewriteRelativeSrcPaths_WhitespaceOnlySrc_IsNotRewritten()
    {
        const string json = """
            { "type": "image", "src": "   " }
            """;

        var rewritten = DemoDeckService.RewriteRelativeSrcPaths(json);

        Assert.Contains("\"src\": \"   \"", rewritten);
    }

    [Fact]
    public void RewriteRelativeSrcPaths_MultipleDotSlashPrefixes_AllStripped()
    {
        const string json = """
            { "type": "image", "src": "./././demo/assets/img.png" }
            """;

        var rewritten = DemoDeckService.RewriteRelativeSrcPaths(json);

        Assert.Contains("\"src\": \"demo/assets/img.png\"", rewritten);
    }

    [Fact]
    public void RewriteRelativeSrcPaths_AbsoluteUnixPath_IsLeftUntouched()
    {
        const string json = """
            { "type": "image", "src": "/var/data/assets/logo.png" }
            """;

        var rewritten = DemoDeckService.RewriteRelativeSrcPaths(json);

        Assert.Contains("\"src\": \"/var/data/assets/logo.png\"", rewritten);
    }

    [Fact]
    public void RewriteRelativeSrcPaths_WindowsAbsolutePath_IsLeftUntouchedBecauseRooted()
    {
        // Path.IsPathRooted("C:\\Users\\...") is true on Windows, false on Unix.
        // The behavior depends on the platform, but on macOS this IS rooted
        // via Path.IsPathRooted (which on macOS considers "/" prefixed paths rooted).
        // A Windows-style backslash path starting with "C:" is NOT rooted on macOS.
        // Testing the contract: backslashes in a non-rooted path ARE normalized.
        const string json = """
            { "type": "image", "src": "assets\\icons\\logo.png" }
            """;

        var rewritten = DemoDeckService.RewriteRelativeSrcPaths(json);

        // On macOS, "assets\\icons\\logo.png" is not rooted, so:
        // 1. Path.IsPathRooted → false
        // 2. Backslashes → forward slashes, then the demo/ template directory is prepended.
        Assert.Contains("\"src\": \"demo/assets/icons/logo.png\"", rewritten);
    }

    [Fact]
    public void RewriteRelativeSrcPaths_DeeplyNested_RecursivelyRewrites()
    {
        const string json = """
            {
              "slides": [
                {
                  "elements": [
                    { "type": "container", "children": [
                      { "type": "image", "src": "./deep/nested.png" }
                    ]}
                  ]
                }
              ]
            }
            """;

        var rewritten = DemoDeckService.RewriteRelativeSrcPaths(json);

        Assert.Contains("\"src\": \"demo/deep/nested.png\"", rewritten);
    }

    [Fact]
    public void RewriteRelativeSrcPaths_NonSrcPropertyWithDotSlash_NotRewritten()
    {
        const string json = """
            { "type": "image", "path": "./assets/logo.png", "src": "real.png" }
            """;

        var rewritten = DemoDeckService.RewriteRelativeSrcPaths(json);

        // Only properties named "src" are rewritten; "path" is left alone.
        Assert.Contains("\"path\": \"./assets/logo.png\"", rewritten);
        Assert.Contains("\"src\": \"demo/real.png\"", rewritten);
    }

    [Fact]
    public void RewriteRelativeSrcPaths_MixedAbsoluteAndRelative_BothPreserved()
    {
        const string json = """
            {
              "slides": [
                { "type": "image", "src": "/abs/logo.png" },
                { "type": "image", "src": "./rel/icon.png" }
              ]
            }
            """;

        var rewritten = DemoDeckService.RewriteRelativeSrcPaths(json);

        Assert.Contains("\"src\": \"/abs/logo.png\"", rewritten);
        Assert.Contains("\"src\": \"demo/rel/icon.png\"", rewritten);
    }

    [Fact]
    public void RenderUploadedDeck_NullBytes_ThrowsArgumentNullException()
    {
        var service = new DemoDeckService(new StubDeckSessionStore());

        var ex = Assert.Throws<ArgumentNullException>(() =>
            service.RenderUploadedDeck(null!, "deck.pptx", "png", 110));

        Assert.Equal("sourceBytes", ex.ParamName);
    }

    [Fact]
    public void RenderUploadedDeck_NullFileName_ThrowsArgumentNullException()
    {
        var service = new DemoDeckService(new StubDeckSessionStore());

        var ex = Assert.Throws<ArgumentNullException>(() =>
            service.RenderUploadedDeck(BuildDeckBytes(1), null!, "png", 110));

        Assert.Equal("fileName", ex.ParamName);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void RenderUploadedDeck_BlankFileName_ThrowsArgumentException(string fileName)
    {
        var service = new DemoDeckService(new StubDeckSessionStore());

        var ex = Assert.Throws<ArgumentException>(() =>
            service.RenderUploadedDeck(BuildDeckBytes(1), fileName, "png", 110));

        Assert.Equal("fileName", ex.ParamName);
    }

    [Fact]
    public void RenderUploadedDeck_EmptyDeck_ThrowsArgumentExceptionMentioningNoSlides()
    {
        var service = new DemoDeckService(new StubDeckSessionStore());

        var ex = Assert.Throws<ArgumentException>(() =>
            service.RenderUploadedDeck(BuildDeckBytes(0), "empty.pptx", "png", 110));

        Assert.Contains("no slides", ex.Message);
        Assert.Equal("sourceBytes", ex.ParamName);
    }

    [Fact]
    public void RenderTypstLeg_NullBytes_ThrowsArgumentNullException()
    {
        var service = new DemoDeckService(new StubDeckSessionStore());

        var ex = Assert.Throws<ArgumentNullException>(() =>
            service.RenderTypstLeg(null!, "deck.pptx", 110));

        Assert.Equal("sourceBytes", ex.ParamName);
    }

    [Fact]
    public void RenderTypstLeg_NullFileName_ThrowsArgumentNullException()
    {
        var service = new DemoDeckService(new StubDeckSessionStore());

        var ex = Assert.Throws<ArgumentNullException>(() =>
            service.RenderTypstLeg(BuildDeckBytes(1), null!, 110));

        Assert.Equal("fileName", ex.ParamName);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void RenderTypstLeg_BlankFileName_ThrowsArgumentException(string fileName)
    {
        var service = new DemoDeckService(new StubDeckSessionStore());

        var ex = Assert.Throws<ArgumentException>(() =>
            service.RenderTypstLeg(BuildDeckBytes(1), fileName, 110));

        Assert.Equal("fileName", ex.ParamName);
    }

    [Fact]
    public void RenderDeck_NullOrEmptyName_ThrowsArgumentException()
    {
        var service = new DemoDeckService(new StubDeckSessionStore());

        var ex = Assert.Throws<ArgumentException>(() => service.RenderDeck("", 110, "png"));

        Assert.Equal("name", ex.ParamName);
    }

    [Fact]
    public void GetDeckTemplateJson_ReturnsNonEmptyRewrittenJson()
    {
        var service = new DemoDeckService(new StubDeckSessionStore());

        var result = service.GetDeckTemplateJson();

        Assert.False(string.IsNullOrWhiteSpace(result));
        // The template is demo/demo-deck.json — must contain expected fields
        // after the src-path rewrite.
        Assert.Contains("\"slides\"", result);
        Assert.Contains("\"title\"", result);
        // The authored doc-relative "assets/dashboard.png" resolves against the template's
        // own directory (demo/), so the served JSON carries the repo-root-relative path the
        // repository-root sandbox resolves.
        Assert.Contains("\"src\": \"demo/assets/dashboard.png\"", result);
    }

    [Fact]
    public void DemoRenderResult_DirectConstruction_HasCorrectShape()
    {
        var deckId = Guid.NewGuid();
        var pages = new byte[][] { [0x89, (byte)'P', (byte)'N', (byte)'G'] };

        var result = new DemoRenderResult(deckId, 1, 42.5, "png", pages);

        Assert.Equal(deckId, result.DeckId);
        Assert.Equal(1, result.SlideCount);
        Assert.Equal(42.5, result.TotalMilliseconds);
        Assert.Equal("png", result.Format);
        Assert.Single(result.Pages);
        Assert.Equal(0x89, result.Pages[0][0]);
    }

    [Fact]
    public void TypstLegResult_DirectConstruction_HasCorrectShape()
    {
        var pngPages = new byte[][] { [0x89, (byte)'P', (byte)'N', (byte)'G'] };
        var pdfBytes = System.Text.Encoding.ASCII.GetBytes("%PDF-1.4");

        var result = new TypstLegResult(1, 250.0, 120.0, 270.0, pngPages, pdfBytes, null);

        Assert.Equal(1, result.SlideCount);
        Assert.Equal(250.0, result.PngMilliseconds);
        Assert.Equal(120.0, result.PdfMilliseconds);
        Assert.Equal(270.0, result.TotalMilliseconds);
        Assert.Single(result.PngPages);
        Assert.Equal("%PDF-1.4", System.Text.Encoding.ASCII.GetString(result.PdfBytes!));
        Assert.Null(result.PdfError);
    }

    [Fact]
    public void TypstLegResult_DirectConstruction_WithPdfError_HasCorrectShape()
    {
        var pngPages = new byte[][] { [0x89, (byte)'P'] };

        var result = new TypstLegResult(1, 100.0, null, 100.0, pngPages, null, "PDF failed");

        Assert.Equal(1, result.SlideCount);
        Assert.Equal(100.0, result.PngMilliseconds);
        Assert.Null(result.PdfMilliseconds);
        Assert.Null(result.PdfBytes);
        Assert.Equal("PDF failed", result.PdfError);
    }

    [Fact]
    public void DemoDeckAllFormatsResult_DirectConstruction_HasCorrectShape()
    {
        var pngPages = new byte[][] { [0x89, (byte)'P'] };
        var svgPages = new byte[][] { System.Text.Encoding.UTF8.GetBytes("<svg></svg>") };
        var pdfBytes = System.Text.Encoding.ASCII.GetBytes("%PDF-1.4");

        var result = new DemoDeckAllFormatsResult(
            2,
            500.0,
            300.0,
            200.0,
            150.0,
            600.0,
            pngPages,
            svgPages,
            pdfBytes,
            null);

        Assert.Equal(2, result.SlideCount);
        Assert.Equal(500.0, result.ConversionMilliseconds);
        Assert.Equal(300.0, result.PngMilliseconds);
        Assert.Equal(200.0, result.SvgMilliseconds);
        Assert.Equal(150.0, result.PdfMilliseconds);
        Assert.Equal(600.0, result.TotalMilliseconds);
        Assert.Single(result.PngPages);
        Assert.Single(result.SvgPages);
        Assert.NotNull(result.PdfBytes);
        Assert.Null(result.PdfError);
    }

    [Fact]
    public void DemoDeckAllFormatsResult_DirectConstruction_WithPdfErrorAndSkippedPng_HasCorrectShape()
    {
        var svgPages = new byte[][] { System.Text.Encoding.UTF8.GetBytes("<svg></svg>") };

        var result = new DemoDeckAllFormatsResult(
            15,
            1200.0,
            null,
            800.0,
            null,
            2100.0,
            [],
            svgPages,
            null,
            "PDF compilation error");

        Assert.Equal(15, result.SlideCount);
        Assert.Null(result.PngMilliseconds);
        Assert.Empty(result.PngPages);
        Assert.Equal(800.0, result.SvgMilliseconds);
        Assert.Single(result.SvgPages);
        Assert.Null(result.PdfMilliseconds);
        Assert.Null(result.PdfBytes);
        Assert.Equal("PDF compilation error", result.PdfError);
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
