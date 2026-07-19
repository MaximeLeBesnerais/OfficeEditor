namespace PptxEditor.Core.Generation.Layout;

/// <summary>
/// Loud failure for invalid or over-constrained layout input (plan.md §3.2 overflow=error,
/// §3.4 loud errors). Carries the element path (e.g. "slides[0].children[2]").
/// </summary>
public sealed class LayoutException : Exception
{
    /// <summary>Element path the failure is attached to.</summary>
    public string Path { get; }

    /// <summary>Creates the exception with a path-prefixed message.</summary>
    public LayoutException(string path, string message)
        : base($"{path}: {message}")
    {
        Path = path;
    }
}
