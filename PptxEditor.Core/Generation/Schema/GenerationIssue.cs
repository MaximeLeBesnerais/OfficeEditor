namespace PptxEditor.Core.Generation.Schema;

/// <summary>Severity of a <see cref="GenerationIssue"/>.</summary>
public enum GenerationIssueSeverity
{
    /// <summary>Non-fatal caveat (e.g. raw hex color off-palette, plan.md §3.1). Document still validates.</summary>
    Warning,

    /// <summary>Contract violation. Document is rejected.</summary>
    Error
}

/// <summary>
/// One validation finding: a JSON path, an actionable message and an optional suggestion
/// ("Did you mean 'label'?"). Formats as "slides[0].children[1]: message".
/// </summary>
public sealed record GenerationIssue(
    string Path,
    string Message,
    string? Suggestion,
    GenerationIssueSeverity Severity)
{
    /// <summary>Path-prefixed rendering, suggestion appended when present.</summary>
    public override string ToString() =>
        Suggestion is null
            ? $"{Path}: {Message}"
            : $"{Path}: {Message} {Suggestion}";
}
