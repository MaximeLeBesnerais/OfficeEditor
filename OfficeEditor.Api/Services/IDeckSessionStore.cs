namespace OfficeEditor.Api.Services;

/// <summary>
/// Session store for uploaded decks (PPTX bytes keyed by deck id).
/// <para>
/// This interface is defined by the W7 edit/anatomy services for exactly what they
/// need. The production implementation is owned by the deck-session workstream (W6);
/// final DI wiring happens at integration time. Tests use a simple in-memory stub.
/// </para>
/// </summary>
public interface IDeckSessionStore
{
    /// <summary>Returns the current deck bytes, or null when the deck id is unknown/expired.</summary>
    byte[]? GetBytes(Guid deckId);

    /// <summary>Runs <paramref name="action"/> while holding the deck's mutation lock.</summary>
    T WithDeckLock<T>(Guid deckId, Func<T> action);

    /// <summary>Replaces the deck's bytes. Callers must hold the deck lock (see <see cref="WithDeckLock{T}"/>).</summary>
    void UpdateBytes(Guid deckId, byte[] bytes);
}
