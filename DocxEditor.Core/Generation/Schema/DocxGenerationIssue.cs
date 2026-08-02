namespace DocxEditor.Core.Generation.Schema;

/// <summary>Severity of a <see cref="DocxGenerationIssue"/>.</summary>
public enum DocxGenerationIssueSeverity
{
    /// <summary>Non-fatal caveat (e.g. raw hex color off-palette). Document still validates.</summary>
    Warning,

    /// <summary>Contract violation. Document is rejected.</summary>
    Error
}

/// <summary>
/// One validation finding: a JSON path, an actionable message and an optional suggestion
/// ("Did you mean 'list'?"). Formats as "sections[0].blocks[1].text: message".
/// </summary>
public sealed record DocxGenerationIssue(
    string Path,
    string Message,
    string? Suggestion,
    DocxGenerationIssueSeverity Severity)
{
    /// <summary>Path-prefixed rendering, suggestion appended when present.</summary>
    public override string ToString() =>
        Suggestion is null
            ? $"{Path}: {Message}"
            : $"{Path}: {Message} {Suggestion}";
}
