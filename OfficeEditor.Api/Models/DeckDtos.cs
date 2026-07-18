namespace OfficeEditor.Api.Models;

/// <summary>Element position/extent in EMU (from the element's a:xfrm).</summary>
public record PositionDto(long X, long Y, long Cx, long Cy);

public record DeckElementDto(
    uint Id,
    string Type,
    string Name,
    string Location,
    PositionDto? Position,
    string? Text,
    List<List<string>>? TableData,
    string? ImagePath);

public record DeckSlideDto(int SlideIndex, IReadOnlyList<DeckElementDto> Elements);

/// <summary>Response for GET /api/decks/{id}/anatomy.</summary>
public record DeckAnatomyDto(Guid DeckId, int SlideCount, IReadOnlyList<DeckSlideDto> Slides);

public record ChangedSlideDto(int SlideIndex, string PreviewUrl);

/// <summary>Per-op failure: operation index (0-based), op type, message.</summary>
public record DeckEditErrorDto(int Index, string Type, string Error);

/// <summary>Response for POST /api/decks/{id}/instructions.</summary>
public record DeckEditResponse(
    bool Success,
    Guid? Revision,
    string? DownloadUrl,
    IReadOnlyList<ChangedSlideDto> ChangedSlides,
    IReadOnlyList<DeckEditErrorDto> Errors);
