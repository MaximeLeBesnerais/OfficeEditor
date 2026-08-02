using System.Text;
using OfficeEditor.Api.Services;

namespace OfficeEditor.Api.Tests.Unit;

public sealed class DeckPreviewServiceTests
{
    private sealed class FakeSlideRenderer : ISlideRenderer
    {
        public int CallCount { get; private set; }
        public List<int> RequestedSlideIndexes { get; } = new();

        public byte[] RenderSlide(DeckSession session, int slideIndex, string normalizedFormat, int ppi)
        {
            CallCount++;
            RequestedSlideIndexes.Add(slideIndex);
            // Distinct bytes per cache key so cache correctness is observable.
            return Encoding.UTF8.GetBytes($"{normalizedFormat}:{slideIndex}:{ppi}");
        }
    }

    private readonly FakeSlideRenderer _renderer = new();
    private readonly DeckPreviewService _service;

    public DeckPreviewServiceTests()
    {
        _service = new DeckPreviewService(_renderer);
    }

    private static DeckSession BuildSession(int slideCount = 3) => new()
    {
        DeckId = Guid.NewGuid(),
        SourceBytes = [1, 2, 3],
        FileName = "deck.pptx",
        SlideCount = slideCount
    };

    [Fact]
    public async Task GetSlidePreviewAsync_FirstCall_RendersAndCachesPage()
    {
        var session = BuildSession();

        var preview = await _service.GetSlidePreviewAsync(session, slideNumber: 1, "png", 150);

        Assert.Equal(1, _renderer.CallCount);
        Assert.Equal("image/png", preview.ContentType);
        Assert.NotEmpty(preview.Bytes);
        Assert.Single(session.RenderedPages);
        Assert.True(session.RenderedPages.ContainsKey(new RenderedPageKey(0, "png", 150)));
    }

    [Fact]
    public async Task GetSlidePreviewAsync_SecondIdenticalCall_ServesFromCache()
    {
        var session = BuildSession();

        var first = await _service.GetSlidePreviewAsync(session, 1, "png", 150);
        var second = await _service.GetSlidePreviewAsync(session, 1, "png", 150);

        Assert.Equal(1, _renderer.CallCount);
        Assert.Same(first.Bytes, second.Bytes);
        Assert.Equal(first.ETag, second.ETag);
    }

    [Fact]
    public async Task GetSlidePreviewAsync_DifferentFormatOrPpi_RendersSeparately()
    {
        var session = BuildSession();

        await _service.GetSlidePreviewAsync(session, 1, "png", 150);
        await _service.GetSlidePreviewAsync(session, 1, "svg", 150);
        await _service.GetSlidePreviewAsync(session, 1, "png", 300);

        Assert.Equal(3, _renderer.CallCount);
        Assert.Equal(3, session.RenderedPages.Count);
        Assert.True(session.RenderedPages.ContainsKey(new RenderedPageKey(0, "svg", 150)));
        Assert.True(session.RenderedPages.ContainsKey(new RenderedPageKey(0, "png", 300)));
    }

    [Fact]
    public async Task GetSlidePreviewAsync_DifferentSlides_RenderedIndependently()
    {
        var session = BuildSession(slideCount: 3);

        await _service.GetSlidePreviewAsync(session, 1, "png", 150);
        await _service.GetSlidePreviewAsync(session, 3, "png", 150);

        Assert.Equal(2, _renderer.CallCount);
        Assert.Equal(new[] { 0, 2 }, _renderer.RequestedSlideIndexes);
    }

    [Fact]
    public async Task GetSlidePreviewAsync_ConvertsOneBasedSlideNumberToZeroBasedIndex()
    {
        var session = BuildSession();

        await _service.GetSlidePreviewAsync(session, slideNumber: 2, "png", 150);

        Assert.Equal(new[] { 1 }, _renderer.RequestedSlideIndexes);
    }

    [Fact]
    public async Task GetSlidePreviewAsync_SvgContentType_IsImageSvgXml()
    {
        var session = BuildSession();

        var preview = await _service.GetSlidePreviewAsync(session, 1, "svg", 150);

        Assert.Equal("image/svg+xml", preview.ContentType);
    }

    [Fact]
    public async Task GetSlidePreviewAsync_ETag_IsQuotedStableHash()
    {
        var session = BuildSession();

        var preview = await _service.GetSlidePreviewAsync(session, 1, "png", 150);

        Assert.StartsWith("\"", preview.ETag);
        Assert.EndsWith("\"", preview.ETag);
        Assert.Equal(34, preview.ETag.Length); // 32 hex chars + 2 quotes
    }

    [Fact]
    public async Task GetSlidePreviewAsync_Warnings_EmptyUntilW3Integration()
    {
        var session = BuildSession();

        var preview = await _service.GetSlidePreviewAsync(session, 1, "png", 150);

        Assert.NotNull(preview.Warnings);
        Assert.Empty(preview.Warnings);
    }

    [Fact]
    public async Task GetSlidePreviewAsync_CacheIsPerSession_NotSharedAcrossDecks()
    {
        var first = BuildSession();
        var second = BuildSession();

        await _service.GetSlidePreviewAsync(first, 1, "png", 150);
        await _service.GetSlidePreviewAsync(second, 1, "png", 150);

        Assert.Equal(2, _renderer.CallCount);
        Assert.Single(first.RenderedPages);
        Assert.Single(second.RenderedPages);
    }

    [Fact]
    public async Task GetSlidePreviewAsync_ConcurrentIdenticalCalls_RenderOnce()
    {
        var session = BuildSession();

        var results = await Task.WhenAll(
            _service.GetSlidePreviewAsync(session, 1, "png", 150),
            _service.GetSlidePreviewAsync(session, 1, "png", 150),
            _service.GetSlidePreviewAsync(session, 1, "png", 150));

        Assert.Equal(1, _renderer.CallCount);
        Assert.All(results, r => Assert.Same(results[0].Bytes, r.Bytes));
    }

    [Fact]
    public async Task GetSlidePreviewAsync_WhenCacheBudgetIsExhausted_DoesNotRetainPage()
    {
        var session = BuildSession();
        session.RenderedPageBytes = ApiResourceLimits.RenderedPagesPerDeckBytes;

        await _service.GetSlidePreviewAsync(session, 1, "png", 150);
        await _service.GetSlidePreviewAsync(session, 1, "png", 150);

        Assert.Equal(2, _renderer.CallCount);
        Assert.Empty(session.RenderedPages);
        Assert.Equal(ApiResourceLimits.RenderedPagesPerDeckBytes, session.RenderedPageBytes);
    }
}
