namespace OfficeEditor.Api.Models;

public record ConvertRequest(
    string TargetFormat,
    string? SampleName = null,
    IReadOnlyDictionary<string, string>? Options = null);

public record ConvertResponse(
    bool Success,
    string? OutputFileName = null,
    string? ContentType = null,
    string? DownloadUrl = null,
    string? PreviewUrl = null,
    string? ErrorMessage = null,
    IReadOnlyList<string>? Messages = null);

public record SampleFileDto(
    string Name,
    string Format,
    string Description);
