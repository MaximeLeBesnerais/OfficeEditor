using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Schema;

/// <summary>
/// Outcome of validating a generation JSON document: the parsed
/// <see cref="Model.DocxGenerationDocument"/> when <see cref="IsValid"/>, plus every error
/// and warning collected in a single pass.
/// </summary>
public sealed record DocxGenerationValidationResult
{
    internal DocxGenerationValidationResult(
        DocxGenerationDocument? document,
        IReadOnlyList<DocxGenerationIssue> errors,
        IReadOnlyList<DocxGenerationIssue> warnings)
    {
        Document = document;
        Errors = errors;
        Warnings = warnings;
    }

    /// <summary>The parsed document. Null when <see cref="Errors"/> is non-empty.</summary>
    public DocxGenerationDocument? Document { get; }

    /// <summary>Contract violations, in document order. Empty when the document is valid.</summary>
    public IReadOnlyList<DocxGenerationIssue> Errors { get; }

    /// <summary>Non-fatal caveats (e.g. off-palette raw hex colors).</summary>
    public IReadOnlyList<DocxGenerationIssue> Warnings { get; }

    /// <summary>True when no errors were collected.</summary>
    public bool IsValid => Errors.Count == 0;

    /// <summary>Returns this result when valid; otherwise throws <see cref="DocxGenerationValidationException"/>.</summary>
    public DocxGenerationValidationResult ThrowIfInvalid()
    {
        if (!IsValid)
        {
            throw new DocxGenerationValidationException(Errors);
        }
        return this;
    }
}
