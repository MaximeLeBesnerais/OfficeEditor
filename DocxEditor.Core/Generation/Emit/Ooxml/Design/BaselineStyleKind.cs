namespace DocxEditor.Core.Generation.Emit.Ooxml.Design;

/// <summary>
/// The set of baseline styles the generator can produce (see <c>DocxEditor.Core/Generation/README.md</c>
/// "design tokens"). Each kind maps to a conventional Word style ID/name; the style manager
/// references the template's existing definition when present and only generates a fresh, ID-safe
/// style when it is missing ("generate baseline styles … where represented").
/// </summary>
public enum BaselineStyleKind
{
    /// <summary>Body text (the document's Normal/default paragraph style).</summary>
    Body,

    /// <summary>Document title.</summary>
    Title,

    /// <summary>Heading level 1.</summary>
    Heading1,

    /// <summary>Heading level 2.</summary>
    Heading2,

    /// <summary>Heading level 3.</summary>
    Heading3,

    /// <summary>Heading level 4.</summary>
    Heading4,

    /// <summary>Heading level 5.</summary>
    Heading5,

    /// <summary>Heading level 6.</summary>
    Heading6,

    /// <summary>Callout/note paragraph (shaded, indented, bordered).</summary>
    Callout,

    /// <summary>Table base style (single-line grid borders).</summary>
    Table,

    /// <summary>Table header cell paragraph style (emphasized header text).</summary>
    TableHeader,

    /// <summary>Code block paragraph style (monospace, shaded).</summary>
    Code
}

/// <summary>Helpers for working with <see cref="BaselineStyleKind"/>.</summary>
public static class BaselineStyleKindExtensions
{
    /// <summary>Maps a heading level (1..6) to its baseline style kind.</summary>
    public static BaselineStyleKind ForHeading(int level) => level switch
    {
        1 => BaselineStyleKind.Heading1,
        2 => BaselineStyleKind.Heading2,
        3 => BaselineStyleKind.Heading3,
        4 => BaselineStyleKind.Heading4,
        5 => BaselineStyleKind.Heading5,
        _ => BaselineStyleKind.Heading6
    };

    /// <summary>True when the kind is one of the six heading levels.</summary>
    public static bool IsHeading(this BaselineStyleKind kind) =>
        kind is >= BaselineStyleKind.Heading1 and <= BaselineStyleKind.Heading6;

    /// <summary>The 1-based outline level of a heading kind (throws for non-headings).</summary>
    public static int HeadingLevel(this BaselineStyleKind kind) =>
        kind.IsHeading()
            ? (kind - BaselineStyleKind.Heading1) + 1
            : throw new ArgumentOutOfRangeException(nameof(kind), kind, "only heading kinds have a level.");
}
