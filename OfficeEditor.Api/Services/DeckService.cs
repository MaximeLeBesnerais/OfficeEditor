using OfficeEditor.Api.Models;
using PptxEditor.Core.Builders;

namespace OfficeEditor.Api.Services;

public interface IDeckService
{
    /// <summary>Deck anatomy for GET /api/decks/{id}/anatomy; null when the deck is unknown.</summary>
    DeckAnatomyDto? GetAnatomy(Guid deckId);

    /// <summary>Raw deck bytes for GET /api/decks/{id}/file; null when the deck is unknown.</summary>
    byte[]? GetDeckBytes(Guid deckId);
}

public class DeckService : IDeckService
{
    private readonly IDeckSessionStore _store;

    public DeckService(IDeckSessionStore store)
    {
        _store = store;
    }

    public DeckAnatomyDto? GetAnatomy(Guid deckId)
    {
        var bytes = _store.GetBytes(deckId);
        if (bytes is null)
        {
            return null;
        }

        using var builder = PresentationBuilder.Open(bytes);
        var anatomy = builder.Analyze();

        var slides = anatomy.Select(slide => new DeckSlideDto(
            slide.SlideIndex,
            slide.Elements.Select(ToDto).ToList())).ToList();

        return new DeckAnatomyDto(deckId, slides.Count, slides);
    }

    public byte[]? GetDeckBytes(Guid deckId) => _store.GetBytes(deckId);

    private static DeckElementDto ToDto(PptxEditor.Core.Models.SlideElement element)
    {
        // Position is only emitted when the full xfrm (offset + extents) is present;
        // partial transforms (e.g. off without ext) are reported as null.
        var position = element is { X: not null, Y: not null, Cx: not null, Cy: not null }
            ? new PositionDto(element.X.Value, element.Y.Value, element.Cx.Value, element.Cy.Value)
            : null;

        return new DeckElementDto(
            element.Id,
            element.Type,
            element.Name,
            element.Location,
            position,
            element.Text,
            element.TableData,
            element.ImagePath);
    }
}
