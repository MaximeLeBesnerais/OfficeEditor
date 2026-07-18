using System.Collections.Concurrent;
using OfficeEditor.Api.Services;

namespace OfficeEditor.Api.Tests.Unit;

/// <summary>
/// Simple in-memory <see cref="IDeckSessionStore"/> for service tests.
/// The production implementation is owned by the deck-session workstream and
/// wired at integration time.
/// </summary>
internal sealed class StubDeckSessionStore : IDeckSessionStore
{
    private readonly ConcurrentDictionary<Guid, byte[]> _decks = new();
    private readonly ConcurrentDictionary<Guid, object> _locks = new();

    public Guid Add(byte[] bytes)
    {
        var id = Guid.NewGuid();
        _decks[id] = bytes;
        return id;
    }

    public byte[]? GetBytes(Guid deckId)
        => _decks.TryGetValue(deckId, out var bytes) ? bytes : null;

    public T WithDeckLock<T>(Guid deckId, Func<T> action)
    {
        lock (_locks.GetOrAdd(deckId, _ => new object()))
        {
            return action();
        }
    }

    public void UpdateBytes(Guid deckId, byte[] bytes)
    {
        if (!_decks.ContainsKey(deckId))
        {
            throw new KeyNotFoundException($"Unknown deck id {deckId}.");
        }
        _decks[deckId] = bytes;
    }
}
