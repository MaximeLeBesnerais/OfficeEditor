namespace DocxEditor.Core.Markdown;

/// <summary>
/// Thrown when markdown cannot be parsed at all. Source-level issues are never
/// swallowed: they surface either as structured nodes with diagnostics or as this
/// domain exception for hard failures.
/// </summary>
public sealed class MarkdownParseException : Exception
{
    public MarkdownParseException(string message)
        : base(message)
    {
    }

    public MarkdownParseException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}
