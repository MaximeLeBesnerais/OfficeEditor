using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Design;

/// <summary>
/// The set of baseline styles the generator can produce (see <c>DocxEditor.Core/Generation/README.md</c>
/// "design tokens"). Each kind maps to a conventional Word style ID/name; the style manager
/// references the template's existing definition when present and only generates a fresh, ID-safe
/// style when it is missing ("generate baseline styles … where represented"). Paragraph-style
/// kinds correspond one-to-one with <see cref="DocxEditor.Core.Generation.Model.TextRole"/> (see
/// <see cref="BaselineStyleKindExtensions.ToTextRole"/>).
///
/// Numeric values are explicitly assigned and stable: inserting a new kind in the middle of the
/// list must never renumber the trailing members, which would silently reinterpret any persisted
/// or compiled-constant value.
/// </summary>
public enum BaselineStyleKind
{
    /// <summary>Body text (the document's Normal/default paragraph style).</summary>
    Body = 0,

    /// <summary>Document title.</summary>
    Title = 1,

    /// <summary>Subtitle line under a title.</summary>
    Subtitle = 2,

    /// <summary>Small uppercase kicker above a title.</summary>
    Eyebrow = 3,

    /// <summary>Heading level 1.</summary>
    Heading1 = 4,

    /// <summary>Heading level 2.</summary>
    Heading2 = 5,

    /// <summary>Heading level 3.</summary>
    Heading3 = 6,

    /// <summary>Heading level 4.</summary>
    Heading4 = 7,

    /// <summary>Heading level 5.</summary>
    Heading5 = 8,

    /// <summary>Heading level 6.</summary>
    Heading6 = 9,

    /// <summary>Muted/secondary body text.</summary>
    MutedBody = 10,

    /// <summary>Small field label.</summary>
    Label = 11,

    /// <summary>Large standalone number or figure.</summary>
    Metric = 12,

    /// <summary>Caption under a metric.</summary>
    MetricLabel = 13,

    /// <summary>Callout/note paragraph (shaded, indented, bordered).</summary>
    Callout = 14,

    /// <summary>Page footer text.</summary>
    Footer = 15,

    /// <summary>Table base style (single-line grid borders).</summary>
    Table = 16,

    /// <summary>Table header cell paragraph style (emphasized header text).</summary>
    TableHeader = 17,

    /// <summary>Table body cell paragraph style.</summary>
    TableBody = 18,

    /// <summary>Code block paragraph style (monospace, shaded).</summary>
    Code = 19
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

    /// <summary>Maps a paragraph-style baseline kind to its semantic text role.</summary>
    public static TextRole ToTextRole(this BaselineStyleKind kind) => kind switch
    {
        BaselineStyleKind.Body => TextRole.Body,
        BaselineStyleKind.Title => TextRole.Title,
        BaselineStyleKind.Subtitle => TextRole.Subtitle,
        BaselineStyleKind.Eyebrow => TextRole.Eyebrow,
        BaselineStyleKind.Heading1 => TextRole.Heading1,
        BaselineStyleKind.Heading2 => TextRole.Heading2,
        BaselineStyleKind.Heading3 => TextRole.Heading3,
        BaselineStyleKind.Heading4 => TextRole.Heading4,
        BaselineStyleKind.Heading5 => TextRole.Heading5,
        BaselineStyleKind.Heading6 => TextRole.Heading6,
        BaselineStyleKind.MutedBody => TextRole.Muted,
        BaselineStyleKind.Label => TextRole.Label,
        BaselineStyleKind.Metric => TextRole.Metric,
        BaselineStyleKind.MetricLabel => TextRole.MetricLabel,
        BaselineStyleKind.Callout => TextRole.Callout,
        BaselineStyleKind.Footer => TextRole.Footer,
        BaselineStyleKind.TableHeader => TextRole.TableHeader,
        BaselineStyleKind.TableBody => TextRole.TableBody,
        _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, "only paragraph-style kinds map to a text role.")
    };

    /// <summary>Maps a semantic text role to its baseline paragraph-style kind.</summary>
    public static BaselineStyleKind ToBaselineStyleKind(this TextRole role) => role switch
    {
        TextRole.Title => BaselineStyleKind.Title,
        TextRole.Subtitle => BaselineStyleKind.Subtitle,
        TextRole.Eyebrow => BaselineStyleKind.Eyebrow,
        TextRole.Heading1 => BaselineStyleKind.Heading1,
        TextRole.Heading2 => BaselineStyleKind.Heading2,
        TextRole.Heading3 => BaselineStyleKind.Heading3,
        TextRole.Heading4 => BaselineStyleKind.Heading4,
        TextRole.Heading5 => BaselineStyleKind.Heading5,
        TextRole.Heading6 => BaselineStyleKind.Heading6,
        TextRole.Body => BaselineStyleKind.Body,
        TextRole.Muted => BaselineStyleKind.MutedBody,
        TextRole.Label => BaselineStyleKind.Label,
        TextRole.Metric => BaselineStyleKind.Metric,
        TextRole.MetricLabel => BaselineStyleKind.MetricLabel,
        TextRole.TableHeader => BaselineStyleKind.TableHeader,
        TextRole.TableBody => BaselineStyleKind.TableBody,
        TextRole.Callout => BaselineStyleKind.Callout,
        TextRole.Footer => BaselineStyleKind.Footer,
        _ => BaselineStyleKind.Body
    };
}
