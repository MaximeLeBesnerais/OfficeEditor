namespace OfficeEditor.Api.Models;

/// <summary>One demo deck catalog entry (GET /api/demo/decks).</summary>
public record DemoDeckDto(string Name, string FileName, string Description, int SlideCount);

/// <summary>
/// Response for POST /api/demo/render: the timed whole-deck preview render (PNG or SVG)
/// plus the deck
/// session id (the source .pptx is downloadable via /api/decks/{deckId}/file).
/// </summary>
public record DemoRenderResponse(
    bool Success,
    Guid DeckId,
    int SlideCount,
    double TotalMilliseconds,
    IReadOnlyList<GeneratedSlidePreviewDto> Previews);

/// <summary>Capabilities of the LibreOffice compare leg (GET /api/demo/compare/capabilities).</summary>
public record CompareCapabilitiesDto(bool Available, string? Version, bool PdfToPpmAvailable, string? SkipReason);

/// <summary>
/// Flat response for POST /api/demo/compare/libreoffice and
/// POST /api/demo/compare/libreoffice-upload: the headless LibreOffice render of one deck
/// (PDF conversion + pdftoppm PNG pages). Timings are null when soffice is unavailable;
/// SlideCount is the real deck slide count (independent of whether pdftoppm ran).
/// </summary>
public record LibreOfficeLegResponse(
    bool Success,
    string Deck,
    int SlideCount,
    bool Available,
    string? Version,
    bool PdfToPpmAvailable,
    double? ConversionMilliseconds,
    double? RasterizationMilliseconds,
    double? TotalMilliseconds,
    IReadOnlyList<GeneratedSlidePreviewDto>? Previews,
    string? PdfDownloadUrl,
    string? Error);

/// <summary>
/// Flat response for POST /api/demo/compare/typst and
/// POST /api/demo/compare/typst-upload: the Typst render of one deck (timed whole-deck
/// PNG previews plus a best-effort whole-deck PDF export). PdfMilliseconds/PdfDownloadUrl
/// are null when the PDF export failed — PdfError then carries the message; the PNG
/// result is always present. TotalMilliseconds is wall-clock; the PNG and PDF legs run
/// concurrently, so it is ~max(PngMilliseconds, PdfMilliseconds), not their sum.
/// </summary>
public record TypstLegResponse(
    bool Success,
    string Deck,
    int SlideCount,
    double PngMilliseconds,
    double? PdfMilliseconds,
    double TotalMilliseconds,
    IReadOnlyList<GeneratedSlidePreviewDto> Previews,
    string? PdfDownloadUrl,
    string? PdfError);
