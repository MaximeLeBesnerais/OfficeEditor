namespace PptxEditor.Core.Generation.Schema;

/// <summary>
/// Outcome of validating a generation JSON document: the parsed
/// <see cref="Model.GenerationDocument"/> when <see cref="IsValid"/>, plus every error
/// and warning collected in a single pass.
/// </summary>
public sealed record GenerationValidationResult
{
    internal GenerationValidationResult(
        Model.GenerationDocument? document,
        IReadOnlyList<GenerationIssue> errors,
        IReadOnlyList<GenerationIssue> warnings)
    {
        Document = document;
        Errors = errors;
        Warnings = warnings;
    }

    /// <summary>The parsed document. Null when <see cref="Errors"/> is non-empty.</summary>
    public Model.GenerationDocument? Document { get; }

    /// <summary>Contract violations, in document order. Empty when the document is valid.</summary>
    public IReadOnlyList<GenerationIssue> Errors { get; }

    /// <summary>Non-fatal caveats (e.g. off-palette raw hex colors).</summary>
    public IReadOnlyList<GenerationIssue> Warnings { get; }

    /// <summary>True when no errors were collected.</summary>
    public bool IsValid => Errors.Count == 0;

    /// <summary>Returns this result when valid; otherwise throws <see cref="GenerationValidationException"/>.</summary>
    public GenerationValidationResult ThrowIfInvalid()
    {
        if (!IsValid)
        {
            throw new GenerationValidationException(Errors);
        }
        return this;
    }
}
