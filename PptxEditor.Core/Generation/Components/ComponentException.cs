namespace PptxEditor.Core.Generation.Components;

/// <summary>
/// Loud failure for invalid component content or expansion input (plan.md §3.4 loud
/// errors, §4 components). Carries the element path (e.g. "slides[0].children[2]").
/// </summary>
public sealed class ComponentException : Exception
{
    /// <summary>Element path the failure is attached to.</summary>
    public string Path { get; }

    /// <summary>Creates the exception with a path-prefixed message.</summary>
    public ComponentException(string path, string message)
        : base($"{path}: {message}")
    {
        Path = path;
    }
}
