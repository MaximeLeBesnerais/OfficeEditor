namespace OfficeEditor.Api.Models;

/// <summary>One validation finding, verbatim from the P1 validator (path, message, suggestion, severity).</summary>
public record GenerationIssueDto(string Path, string Message, string? Suggestion, string Severity);

/// <summary>One per-slide preview of a generated deck, base64-encoded.</summary>
public record GeneratedSlidePreviewDto(int Slide, string Format, string ContentType, string ContentBase64);

/// <summary>
/// Response for POST /api/decks/generate. Mirrors the MCP deck_generate
/// payload: same fields, transport-appropriate deck reference (deckId + downloadUrl here,
/// deckHandle + pptxBase64 in MCP).
/// </summary>
public record GenerateDeckResponse(
    bool Success,
    Guid? DeckId = null,
    int SlideCount = 0,
    string? DownloadUrl = null,
    IReadOnlyList<GeneratedSlidePreviewDto>? Previews = null,
    IReadOnlyList<GenerationIssueDto>? Errors = null,
    IReadOnlyList<GenerationIssueDto>? Warnings = null,
    IReadOnlyList<string>? PipelineWarnings = null,
    string? PreviewError = null,
    string? ErrorMessage = null,
    double GenerationMilliseconds = 0,
    double TotalMilliseconds = 0);
