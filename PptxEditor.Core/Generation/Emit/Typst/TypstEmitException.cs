namespace PptxEditor.Core.Generation.Emit.Typst;

/// <summary>
/// Loud failure for draw-tree content the Typst emitter cannot render faithfully
/// (loud errors; Typst never approximates silently). Carries the element
/// path (e.g. "slides[0].root.children[2]").
/// </summary>
public sealed class TypstEmitException : Exception
{
    /// <summary>Element path the failure is attached to.</summary>
    public string Path { get; }

    /// <summary>Creates the exception with a path-prefixed message.</summary>
    public TypstEmitException(string path, string message)
        : base($"{path}: {message}")
    {
        Path = path;
    }
}
