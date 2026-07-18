using OfficeEditor.Api.Models;
using PptxEditor.Core.Builders;
using PptxEditor.Core.Instructions;
using PptxEditor.Core.Serialization;

namespace OfficeEditor.Api.Services;

public interface IDeckEditService
{
    /// <summary>
    /// Applies a JSON instruction batch ({ "operations": [...] }) to the deck.
    /// Returns null when the deck id is unknown (endpoint maps that to 404).
    /// </summary>
    DeckEditResponse? ApplyInstructions(Guid deckId, string instructionsJson);
}

public class DeckEditService : IDeckEditService
{
    private readonly IDeckSessionStore _store;
    private readonly PptxJsonInstructionParser _parser = new();
    private readonly PptxInstructionEngine _engine = new();

    public DeckEditService(IDeckSessionStore store)
    {
        _store = store;
    }

    public DeckEditResponse? ApplyInstructions(Guid deckId, string instructionsJson)
    {
        if (_store.GetBytes(deckId) is null)
        {
            return null;
        }

        PptxEditor.Core.Models.PptxInstructionSet instructionSet;
        try
        {
            instructionSet = _parser.Parse(instructionsJson);
        }
        catch (ArgumentException ex)
        {
            return Failure(deckId, new DeckEditErrorDto(-1, "parse", ex.Message));
        }

        // Re-read and persist under the deck lock so concurrent edits cannot
        // interleave between read-modify-write of the deck bytes.
        return _store.WithDeckLock(deckId, () =>
        {
            var bytes = _store.GetBytes(deckId);
            if (bytes is null)
            {
                return null;
            }

            using var builder = PresentationBuilder.Open(bytes);
            var result = _engine.Apply(builder, instructionSet);

            if (result.AppliedOps > 0)
            {
                _store.UpdateBytes(deckId, builder.SaveToBytes());
            }

            return ToResponse(deckId, result);
        });
    }

    private static DeckEditResponse ToResponse(Guid deckId, PptxEditResult result)
    {
        return new DeckEditResponse(
            Success: result.Success,
            Revision: result.Revision,
            DownloadUrl: $"/api/decks/{deckId}/file",
            ChangedSlides: result.ChangedSlides
                .Select(slideIndex => new ChangedSlideDto(
                    slideIndex,
                    $"/api/decks/{deckId}/slides/{slideIndex}/preview"))
                .ToList(),
            Errors: result.FailedOps
                .Select(error => new DeckEditErrorDto(error.Index, error.Type, error.Error))
                .ToList());
    }

    private static DeckEditResponse Failure(Guid deckId, params DeckEditErrorDto[] errors)
    {
        return new DeckEditResponse(
            Success: false,
            Revision: null,
            DownloadUrl: $"/api/decks/{deckId}/file",
            ChangedSlides: Array.Empty<ChangedSlideDto>(),
            Errors: errors);
    }
}
