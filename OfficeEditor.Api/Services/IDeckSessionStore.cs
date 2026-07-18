namespace OfficeEditor.Api.Services;

/// <summary>
/// Session store for uploaded decks (PPTX bytes keyed by deck id). Lives ONLY in the
/// API layer — no session state may leak into PptxEditor.Core / OfficeEditor.Core
/// (stateless-core rule).
/// <para>
/// Single unified contract: the session-object members serve the upload/preview
/// pipeline (W6); the bytes/lock members serve the edit/anatomy pipeline (W7).
/// </para>
/// </summary>
public interface IDeckSessionStore
{
    /// <summary>Stores a new session and returns its generated deck id.</summary>
    Guid Store(byte[] sourceBytes, string fileName, int slideCount);

    /// <summary>Stores an already-built session (used by revision-bump edits).</summary>
    void Store(DeckSession session);

    bool TryGet(Guid deckId, out DeckSession? session);
    void Remove(Guid deckId);

    /// <summary>Deletes the session-owned temp directories; called on eviction/removal.</summary>
    void SweepSession(DeckSession session);

    /// <summary>Returns the current deck bytes, or null when the deck id is unknown/expired.</summary>
    byte[]? GetBytes(Guid deckId);

    /// <summary>Runs <paramref name="action"/> while holding the deck's mutation lock.</summary>
    T WithDeckLock<T>(Guid deckId, Func<T> action);

    /// <summary>Replaces the deck's bytes. Callers must hold the deck lock (see <see cref="WithDeckLock{T}"/>).</summary>
    void UpdateBytes(Guid deckId, byte[] bytes);
}
