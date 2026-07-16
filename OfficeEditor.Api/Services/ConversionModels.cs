namespace OfficeEditor.Api.Services;

public record ConversionRequest(
    byte[] SourceBytes,
    string SourceFileName,
    ConversionTargetFormat TargetFormat,
    IReadOnlyDictionary<string, string>? Options = null);

public record ConversionResult(
    bool Success,
    byte[]? OutputBytes,
    string OutputFileName,
    string ContentType,
    string? ErrorMessage = null,
    IReadOnlyList<string>? Messages = null);
