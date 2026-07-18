using OfficeEditor.Mcp.Sessions;

namespace OfficeEditor.Mcp.Tests;

/// <summary>
/// Session-store semantics: sliding 30-minute lifetime (via an injected clock),
/// lazy expiry, temp-dir sweeping, revision bumps. No timers — deterministic.
/// </summary>
public sealed class DeckSessionStoreTests
{
    [Fact]
    public void Store_ThenTryGet_ReturnsSession()
    {
        using var store = new DeckSessionStore(enableSweepTimer: false);

        var session = store.Store(new byte[] { 1, 2, 3 }, "deck.pptx", 5);

        Assert.True(store.TryGet(session.Handle, out var fetched));
        Assert.Equal(session.Handle, fetched!.Handle);
        Assert.Equal(5, fetched.SlideCount);
        Assert.Equal(0, fetched.Revision);
        Assert.True(Directory.Exists(session.TempDirectory));
    }

    [Fact]
    public void TryGet_SlidesTheExpiryForward()
    {
        var now = DateTimeOffset.UtcNow;
        using var store = new DeckSessionStore(
            clock: () => now,
            slidingLifetime: TimeSpan.FromMinutes(30),
            enableSweepTimer: false);
        var session = store.Store(new byte[] { 1 }, "deck.pptx", 1);

        // 29 minutes pass, then an access: the handle survives another full window.
        now = now.AddMinutes(29);
        Assert.True(store.TryGet(session.Handle, out _));
        now = now.AddMinutes(29);
        Assert.True(store.TryGet(session.Handle, out _));

        // No access for 31 minutes → expired (lazy sweep on access).
        now = now.AddMinutes(31);
        Assert.False(store.TryGet(session.Handle, out _));
        // The lazy sweep also deletes the session temp dir.
        Assert.False(Directory.Exists(session.TempDirectory));
    }

    [Fact]
    public void Expiry_SweepsTheSessionTempDirectory()
    {
        var now = DateTimeOffset.UtcNow;
        using var store = new DeckSessionStore(
            clock: () => now,
            slidingLifetime: TimeSpan.FromMinutes(30),
            enableSweepTimer: false);
        var session = store.Store(new byte[] { 1 }, "deck.pptx", 1);
        Assert.True(Directory.Exists(session.TempDirectory));

        now = now.AddMinutes(31);
        store.SweepExpired();

        Assert.False(Directory.Exists(session.TempDirectory));
        Assert.Equal(0, store.Count);
    }

    [Fact]
    public void UpdateBytes_BumpsRevisionAndSlideCount()
    {
        using var store = new DeckSessionStore(enableSweepTimer: false);
        var session = store.Store(new byte[] { 1 }, "deck.pptx", 3);

        store.UpdateBytes(session.Handle, new byte[] { 1, 2 }, 4);

        Assert.True(store.TryGet(session.Handle, out var updated));
        Assert.Equal(1, updated!.Revision);
        Assert.Equal(4, updated.SlideCount);
        Assert.Equal(2, updated.DeckBytes.Length);
    }

    [Fact]
    public void UpdateBytes_UnknownHandle_Throws()
    {
        using var store = new DeckSessionStore(enableSweepTimer: false);

        Assert.Throws<KeyNotFoundException>(() =>
            store.UpdateBytes(Guid.NewGuid(), new byte[] { 1 }, 1));
    }

    [Fact]
    public void Dispose_SweepsAllSessionDirectories()
    {
        string tempDirectory;
        using (var store = new DeckSessionStore(enableSweepTimer: false))
        {
            var session = store.Store(new byte[] { 1 }, "deck.pptx", 1);
            tempDirectory = session.TempDirectory;
            Assert.True(Directory.Exists(tempDirectory));
        }

        Assert.False(Directory.Exists(tempDirectory));
    }
}
