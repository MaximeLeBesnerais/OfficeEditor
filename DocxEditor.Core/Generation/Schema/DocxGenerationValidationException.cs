using OfficeEditor.Core.Exceptions;

namespace DocxEditor.Core.Generation.Schema;

/// <summary>
/// Domain error raised when a generation document violates the vocabulary contract.
/// Carries every collected <see cref="DocxGenerationIssue"/> (JSON path + suggestion) so AI
/// and human authors can fix all problems at once.
/// </summary>
public sealed class DocxGenerationValidationException : OfficeEditorException
{
    /// <summary>All validation errors, in document order.</summary>
    public IReadOnlyList<DocxGenerationIssue> Issues { get; }

    public DocxGenerationValidationException(IReadOnlyList<DocxGenerationIssue> issues)
        : base(BuildMessage(issues))
    {
        Issues = issues;
    }

    private static string BuildMessage(IReadOnlyList<DocxGenerationIssue> issues)
    {
        var header = issues.Count == 1
            ? "Invalid DOCX generation document (1 error):"
            : $"Invalid DOCX generation document ({issues.Count} errors):";
        return header + Environment.NewLine + string.Join(Environment.NewLine, issues.Select(i => "  " + i));
    }
}
