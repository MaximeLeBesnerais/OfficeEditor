using System.Collections.Concurrent;
using OfficeEditor.Api.Services;

namespace OfficeEditor.Api.Tests.Unit;

/// <summary>
/// Simple in-memory <see cref="IDeckSessionStore"/> for service tests. Keeps the
/// bytes/lock semantics the edit/anatomy services rely on; the session-object members
/// are minimal (no cache eviction, no temp-dir sweeping).
/// </summary>
internal sealed class StubDeckSessionStore : IDeckSessionStore
{
    private readonly ConcurrentDictionary<Guid, DeckSession> _decks = new();
    private readonly ConcurrentDictionary<Guid, object> _locks = new();

    public Guid Add(byte[] bytes)
    {
        var id = Guid.NewGuid();
        _decks[id] = NewSession(id, bytes);
        return id;
    }

    public Guid Store(byte[] sourceBytes, string fileName, int slideCount)
    {
        var session = new DeckSession
        {
            DeckId = Guid.NewGuid(),
            SourceBytes = sourceBytes,
            FileName = fileName,
            SlideCount = slideCount
        };
        _decks[session.DeckId] = session;
        return session.DeckId;
    }

    public void Store(DeckSession session)
    {
        _decks[session.DeckId] = session;
    }

    public bool TryGet(Guid deckId, out DeckSession? session)
        => _decks.TryGetValue(deckId, out session);

    public void Remove(Guid deckId)
    {
        _decks.TryRemove(deckId, out _);
    }

    public void SweepSession(DeckSession session)
    {
        // No-op: the stub owns no temp directories.
    }

    public byte[]? GetBytes(Guid deckId)
        => _decks.TryGetValue(deckId, out var session) ? session.SourceBytes : null;

    public T WithDeckLock<T>(Guid deckId, Func<T> action)
    {
        lock (_locks.GetOrAdd(deckId, _ => new object()))
        {
            return action();
        }
    }

    public void UpdateBytes(Guid deckId, byte[] bytes)
    {
        if (!_decks.TryGetValue(deckId, out var existing))
        {
            throw new KeyNotFoundException($"Unknown deck id {deckId}.");
        }

        // Revision bump: replace the session, carrying the stored slide count
        // (service tests re-open the bytes themselves when they need it).
        _decks[deckId] = new DeckSession
        {
            DeckId = existing.DeckId,
            SourceBytes = bytes,
            FileName = existing.FileName,
            SlideCount = existing.SlideCount,
            Revision = existing.Revision + 1
        };
    }

    private static DeckSession NewSession(Guid deckId, byte[] bytes)
        => new()
        {
            DeckId = deckId,
            SourceBytes = bytes,
            FileName = "stub.pptx",
            SlideCount = 0
        };
}
