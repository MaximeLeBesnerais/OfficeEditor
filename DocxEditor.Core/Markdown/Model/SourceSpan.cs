namespace DocxEditor.Core.Markdown.Model;

/// <summary>
/// A source location within the original markdown source. Character offsets are
/// absolute (0-based) positions into the source string; <see cref="Line"/> is the
/// 0-based line number. Where Markdig does not expose precise locations, both
/// <see cref="Start"/> and <see cref="End"/> are -1.
/// </summary>
public readonly record struct SourceSpan(int Start, int End, int Line)
{
    public static SourceSpan None { get; } = new(-1, -1, -1);

    public int Length => Start >= 0 && End >= Start ? End - Start + 1 : 0;

    public bool IsEmpty => Start < 0 || End < Start;

    public static implicit operator SourceSpan((int Start, int End, int Line) value) =>
        new(value.Start, value.End, value.Line);
}
