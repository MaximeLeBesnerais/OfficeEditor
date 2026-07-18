using Microsoft.Extensions.Caching.Memory;
using PptxEditor.Core.Builders;

namespace OfficeEditor.Api.Services;

/// <summary>
/// Cache key for a single rendered page inside a deck session. The deck revision is
/// part of the session itself (a new revision = a new session entry), so cache-key
/// correctness is guaranteed by keying pages within the owning session.
/// </summary>
public sealed record RenderedPageKey(int SlideIndex, string Format, int Ppi);

/// <summary>
/// Server-side state for one uploaded deck. Lives ONLY in the API layer — no session
/// state may leak into PptxEditor.Core / OfficeEditor.Core (stateless-core rule).
/// Memory-budget conscious: keeps the raw PPTX bytes and re-opens a PresentationBuilder
/// lazily per render instead of holding a permanently open PresentationDocument.
/// </summary>
public sealed class DeckSession
{
    public required Guid DeckId { get; init; }
    public required byte[] SourceBytes { get; init; }
    public required string FileName { get; init; }
    public required int SlideCount { get; init; }

    /// <summary>
    /// Monotonic revision, bumped by edit operations (W7). Rendered pages are cached
    /// per session instance, so a revision bump (new session) naturally invalidates them.
    /// </summary>
    public int Revision { get; set; }

    /// <summary>
    /// Serializes edit vs. render work for this deck. Never disposed by the store:
    /// eviction can race an in-flight render, and SemaphoreSlim needs no disposal
    /// as long as AvailableWaitHandle is never accessed (GC reclaims it).
    /// </summary>
    public SemaphoreSlim RenderLock { get; } = new(1, 1);

    /// <summary>Rendered pages keyed by (slideIndex, format, ppi) within this revision.</summary>
    public Dictionary<RenderedPageKey, byte[]> RenderedPages { get; } = new();

    /// <summary>
    /// Session-owned temp directories (fonts, extracted assets). Deleted by the
    /// session-store sweeper when the session is evicted or removed. Lock on this
    /// list instance when mutating it after the session is stored.
    /// </summary>
    public List<string> TempDirectories { get; } = new();
}

public sealed class InMemoryDeckSessionStore : IDeckSessionStore
{
    private static readonly TimeSpan SlidingLifetime = TimeSpan.FromMinutes(30);

    private readonly IMemoryCache _cache;

    public InMemoryDeckSessionStore(IMemoryCache cache)
    {
        _cache = cache;
    }

    public Guid Store(byte[] sourceBytes, string fileName, int slideCount)
    {
        ArgumentNullException.ThrowIfNull(sourceBytes);
        ArgumentException.ThrowIfNullOrWhiteSpace(fileName);
        ArgumentOutOfRangeException.ThrowIfNegative(slideCount);

        var session = new DeckSession
        {
            DeckId = Guid.NewGuid(),
            SourceBytes = sourceBytes,
            FileName = fileName,
            SlideCount = slideCount,
            Revision = 0
        };

        Store(session);
        return session.DeckId;
    }

    public void Store(DeckSession session)
    {
        ArgumentNullException.ThrowIfNull(session);

        var options = new MemoryCacheEntryOptions()
            .SetSlidingExpiration(SlidingLifetime)
            // Sweeper: session-owned temp dirs must not outlive the cache entry.
            .RegisterPostEvictionCallback((_, value, _, _) =>
            {
                if (value is DeckSession evicted)
                {
                    SweepSession(evicted);
                }
            });

        _cache.Set(session.DeckId, session, options);
    }

    public bool TryGet(Guid deckId, out DeckSession? session)
    {
        return _cache.TryGetValue(deckId, out session);
    }

    public void Remove(Guid deckId)
    {
        _cache.Remove(deckId);
    }

    public byte[]? GetBytes(Guid deckId)
    {
        return TryGet(deckId, out var session) ? session!.SourceBytes : null;
    }

    public T WithDeckLock<T>(Guid deckId, Func<T> action)
    {
        ArgumentNullException.ThrowIfNull(action);

        if (!TryGet(deckId, out var session) || session is null)
        {
            // Unknown/expired deck: no lock to hold. Edit callers re-check GetBytes
            // inside the action, so this still maps to 404 rather than throwing.
            return action();
        }

        // Serialize against renders and other edits of this deck — the preview
        // pipeline takes the same per-deck semaphore (DeckPreviewService).
        session.RenderLock.Wait();
        try
        {
            return action();
        }
        finally
        {
            session.RenderLock.Release();
        }
    }

    public void UpdateBytes(Guid deckId, byte[] bytes)
    {
        ArgumentNullException.ThrowIfNull(bytes);

        if (!TryGet(deckId, out var existing) || existing is null)
        {
            throw new KeyNotFoundException($"Unknown deck id '{deckId}'.");
        }

        // Slide-ops instructions (duplicate/move/delete) can change the slide count,
        // and preview validation trusts SlideCount: re-derive it from the new bytes
        // instead of carrying a stale count into the new revision.
        int slideCount;
        using (var builder = PresentationBuilder.Open(bytes))
        {
            slideCount = builder.SlideCount;
        }

        // Revision bump = a NEW session entry (see DeckSession.Revision): pages cached
        // on the old revision are discarded with it, and evicting the old entry sweeps
        // its session-owned temp dirs. The fresh RenderLock is safe: callers already
        // inside WithDeckLock hold the OLD session's lock for their whole critical
        // section, so old and new sessions are coherent snapshots.
        var revision = new DeckSession
        {
            DeckId = existing.DeckId,
            SourceBytes = bytes,
            FileName = existing.FileName,
            SlideCount = slideCount,
            Revision = existing.Revision + 1
        };

        Store(revision);
    }

    public void SweepSession(DeckSession session)
    {
        // PostEvictionCallbacks fire on a thread-pool thread, so sweep must be
        // re-entrant and must not enumerate a list another thread could touch.
        string[] directories;
        lock (session.TempDirectories)
        {
            directories = session.TempDirectories.ToArray();
            session.TempDirectories.Clear();
        }

        foreach (var directory in directories)
        {
            try
            {
                if (Directory.Exists(directory))
                {
                    Directory.Delete(directory, recursive: true);
                }
            }
            catch (IOException)
            {
                // Temp-dir cleanup is best-effort; the OS temp cleaner is the backstop.
            }
            catch (UnauthorizedAccessException)
            {
                // Temp-dir cleanup is best-effort; the OS temp cleaner is the backstop.
            }
        }
    }
}
