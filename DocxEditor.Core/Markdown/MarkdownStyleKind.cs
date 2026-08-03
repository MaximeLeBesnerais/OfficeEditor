namespace DocxEditor.Core.Markdown;

/// <summary>
/// The OOXML style kind a markdown element requires. Mirrors the Wordprocessing
/// <c>w:type</c> values so the emitter knows whether a resolved (or generated)
/// style is a paragraph, character or table style.
/// </summary>
public enum MarkdownStyleKind
{
    Paragraph,
    Character,
    Table
}

/// <summary>
/// Maps markdown semantic keys (the <c>StyleMapping</c> dictionary keys) to the
/// <see cref="MarkdownStyleKind"/> the element must use. Inline code and hyperlinks
/// are character styles (they style a run), tables are table styles, everything else
/// is a paragraph style.
/// </summary>
public static class MarkdownStyleKinds
{
    public static MarkdownStyleKind ForElement(string element) => element switch
    {
        "codeInline" or "hyperlink" => MarkdownStyleKind.Character,
        "table" => MarkdownStyleKind.Table,
        _ => MarkdownStyleKind.Paragraph
    };
}
