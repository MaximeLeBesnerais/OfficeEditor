using Microsoft.Extensions.Caching.Memory;
using OfficeEditor.Api.Services;

namespace OfficeEditor.Api.Tests.Unit;

public sealed class DeckSessionStoreTests : IDisposable
{
    private readonly MemoryCache _cache = new(new MemoryCacheOptions());
    private readonly InMemoryDeckSessionStore _store;

    public DeckSessionStoreTests()
    {
        _store = new InMemoryDeckSessionStore(_cache);
    }

    public void Dispose() => _cache.Dispose();

    [Fact]
    public void Store_ThenTryGet_ReturnsSessionWithSlideCount()
    {
        byte[] bytes = [1, 2, 3];

        var deckId = _store.Store(bytes, "deck.pptx", 7);
        var found = _store.TryGet(deckId, out var session);

        Assert.True(found);
        Assert.NotNull(session);
        Assert.Equal(deckId, session!.DeckId);
        Assert.Same(bytes, session.SourceBytes);
        Assert.Equal("deck.pptx", session.FileName);
        Assert.Equal(7, session.SlideCount);
        Assert.Equal(0, session.Revision);
        Assert.Empty(session.RenderedPages);
        Assert.Empty(session.TempDirectories);
    }

    [Fact]
    public void TryGet_UnknownDeck_ReturnsFalseAndNullSession()
    {
        var found = _store.TryGet(Guid.NewGuid(), out var session);

        Assert.False(found);
        Assert.Null(session);
    }

    [Fact]
    public void Store_ReturnsUniqueDeckIdsAcrossCalls()
    {
        var first = _store.Store([1], "a.pptx", 1);
        var second = _store.Store([2], "b.pptx", 2);

        Assert.NotEqual(Guid.Empty, first);
        Assert.NotEqual(first, second);
        Assert.True(_store.TryGet(first, out _));
        Assert.True(_store.TryGet(second, out _));
    }

    [Fact]
    public void Store_NullBytes_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>(() => _store.Store(null!, "deck.pptx", 1));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Store_InvalidFileName_ThrowsArgumentException(string? fileName)
    {
        Assert.ThrowsAny<ArgumentException>(() => _store.Store([1], fileName!, 1));
    }

    [Fact]
    public void Store_NegativeSlideCount_ThrowsArgumentOutOfRangeException()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => _store.Store([1], "deck.pptx", -1));
    }

    [Fact]
    public void Remove_UnknownDeck_DoesNotThrow()
    {
        _store.Remove(Guid.NewGuid());
    }

    [Fact]
    public void Remove_SessionWithTempDir_SweeperDeletesTempDir()
    {
        var tempDir = CreateSessionTempDir();
        var session = BuildSessionWithTempDir(tempDir);
        _store.Store(session);

        _store.Remove(session.DeckId);

        // PostEvictionCallbacks run on a thread-pool thread: poll, don't assume sync.
        Assert.True(SpinWait.SpinUntil(() => !Directory.Exists(tempDir), TimeSpan.FromSeconds(5)),
            "sweeper did not delete the session temp dir after removal");
        Assert.False(_store.TryGet(session.DeckId, out _));
    }

    [Fact]
    public void CacheEviction_SessionWithTempDir_SweeperDeletesTempDir()
    {
        var tempDir = CreateSessionTempDir();
        var session = BuildSessionWithTempDir(tempDir);
        _store.Store(session);

        _cache.Compact(1.0);

        Assert.True(SpinWait.SpinUntil(() => !Directory.Exists(tempDir), TimeSpan.FromSeconds(5)),
            "sweeper did not delete the session temp dir after cache eviction");
    }

    [Fact]
    public void Sweep_ClearsTempDirectoryList()
    {
        var tempDir = CreateSessionTempDir();
        var session = BuildSessionWithTempDir(tempDir);
        _store.Store(session);

        _store.Remove(session.DeckId);

        Assert.True(SpinWait.SpinUntil(
            () =>
            {
                lock (session.TempDirectories)
                {
                    return session.TempDirectories.Count == 0;
                }
            }, TimeSpan.FromSeconds(5)));
    }

    [Fact]
    public void SweepSession_CalledDirectly_IsIdempotent()
    {
        var tempDir = CreateSessionTempDir();
        var session = BuildSessionWithTempDir(tempDir);

        _store.SweepSession(session);
        _store.SweepSession(session);

        Assert.False(Directory.Exists(tempDir));
        Assert.Empty(session.TempDirectories);
    }

    [Fact]
    public void Sweep_MissingTempDir_DoesNotThrow()
    {
        var session = new DeckSession
        {
            DeckId = Guid.NewGuid(),
            SourceBytes = [1],
            FileName = "deck.pptx",
            SlideCount = 1
        };
        session.TempDirectories.Add(Path.Combine(Path.GetTempPath(), "w6-nonexistent-" + Guid.NewGuid().ToString("N")));

        _store.SweepSession(session);

        Assert.Empty(session.TempDirectories);
    }

    private static string CreateSessionTempDir()
    {
        var dir = Path.Combine(Path.GetTempPath(), "w6-session-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        File.WriteAllText(Path.Combine(dir, "font.ttf"), "fake");
        return dir;
    }

    private static DeckSession BuildSessionWithTempDir(string tempDir)
    {
        var session = new DeckSession
        {
            DeckId = Guid.NewGuid(),
            SourceBytes = [1],
            FileName = "deck.pptx",
            SlideCount = 3
        };
        session.TempDirectories.Add(tempDir);
        return session;
    }
}
