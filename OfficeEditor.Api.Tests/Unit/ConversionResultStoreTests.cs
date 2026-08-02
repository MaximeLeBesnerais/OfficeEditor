using Microsoft.Extensions.Caching.Memory;
using OfficeEditor.Api.Services;

namespace OfficeEditor.Api.Tests.Unit;

public sealed class ConversionResultStoreTests : IDisposable
{
    private readonly MemoryCache _cache = new(new MemoryCacheOptions());
    private readonly InMemoryConversionResultStore _store;

    public ConversionResultStoreTests()
    {
        _store = new InMemoryConversionResultStore(_cache);
    }

    public void Dispose() => _cache.Dispose();

    [Fact]
    public void Store_ThenTryGet_ReturnsSamePayload()
    {
        byte[] bytes = [1, 2, 3, 4];

        var id = _store.Store(bytes, "application/pdf", "out.pdf");
        var found = _store.TryGet(id, out var result);

        Assert.True(found);
        Assert.NotNull(result);
        Assert.Same(bytes, result!.Bytes);
        Assert.Equal("application/pdf", result.ContentType);
        Assert.Equal("out.pdf", result.FileName);
    }

    [Fact]
    public void TryGet_UnknownId_ReturnsFalseAndNullResult()
    {
        var found = _store.TryGet(Guid.NewGuid(), out var result);

        Assert.False(found);
        Assert.Null(result);
    }

    [Fact]
    public void Store_ReturnsUniqueIdsAcrossCalls()
    {
        var first = _store.Store([1], "application/pdf", "a.pdf");
        var second = _store.Store([2], "application/pdf", "b.pdf");

        Assert.NotEqual(Guid.Empty, first);
        Assert.NotEqual(first, second);
        Assert.True(_store.TryGet(first, out _));
        Assert.True(_store.TryGet(second, out _));
    }

    [Fact]
    public void Store_NullBytes_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>(() => _store.Store(null!, "application/pdf", "out.pdf"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Store_InvalidContentType_ThrowsArgumentException(string? contentType)
    {
        Assert.ThrowsAny<ArgumentException>(() => _store.Store([1], contentType!, "out.pdf"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Store_InvalidFileName_ThrowsArgumentException(string? fileName)
    {
        Assert.ThrowsAny<ArgumentException>(() => _store.Store([1], "application/pdf", fileName!));
    }

    [Fact]
    public void Store_RegistersEntriesWithThirtyMinuteSlidingExpiration()
    {
        // White-box check of the eviction policy: entries must renew their
        // lifetime on every access (sliding), with a 30-minute window.
        var options = InMemoryConversionResultStore.CreateCacheOptions(123);
        Assert.Equal(TimeSpan.FromMinutes(30), options.SlidingExpiration);
        Assert.Null(options.AbsoluteExpirationRelativeToNow);
        Assert.Equal(123, options.Size);
    }

    [Fact]
    public void TryGet_RepeatedAccess_KeepsEntryRetrievable()
    {
        // Sliding expiration means reads renew the entry; repeated access must
        // never make a live entry disappear.
        var id = _store.Store([1], "application/pdf", "out.pdf");

        for (var i = 0; i < 5; i++)
        {
            Assert.True(_store.TryGet(id, out var result));
            Assert.NotNull(result);
        }
    }

    [Fact]
    public void TryGet_AfterExternalCacheEviction_ReturnsFalse()
    {
        // Proves the store delegates to the injected IMemoryCache rather than
        // keeping its own private copy: evicting the cache removes entries.
        var id = _store.Store([1], "application/pdf", "out.pdf");

        _cache.Compact(1.0);

        Assert.False(_store.TryGet(id, out var result));
        Assert.Null(result);
    }
}
