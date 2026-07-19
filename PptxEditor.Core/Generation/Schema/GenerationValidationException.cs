using OfficeEditor.Core.Exceptions;

namespace PptxEditor.Core.Generation.Schema;

/// <summary>
/// Domain error raised when a generation document violates the vocabulary contract.
/// Carries every collected <see cref="GenerationIssue"/> (JSON path + suggestion) so AI
/// and human authors can fix all problems at once.
/// </summary>
public sealed class GenerationValidationException : OfficeEditorException
{
    /// <summary>All validation errors, in document order.</summary>
    public IReadOnlyList<GenerationIssue> Issues { get; }

    public GenerationValidationException(IReadOnlyList<GenerationIssue> issues)
        : base(BuildMessage(issues))
    {
        Issues = issues;
    }

    private static string BuildMessage(IReadOnlyList<GenerationIssue> issues)
    {
        var header = issues.Count == 1
            ? "Invalid generation document (1 error):"
            : $"Invalid generation document ({issues.Count} errors):";
        return header + Environment.NewLine + string.Join(Environment.NewLine, issues.Select(i => "  " + i));
    }
}
