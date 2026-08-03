namespace DocxEditor.Core.Generation.Model;

/// <summary>
/// Page orientation of a section. Custom and named page sizes are always expressed in
/// portrait (upright) dimensions; a landscape orientation swaps width and height at emit
/// time. Never pre-swap the dimensions.
/// </summary>
public enum PageOrientation
{
    /// <summary>Upright page (default).</summary>
    Portrait,

    /// <summary>Rotated page; width and height are swapped.</summary>
    Landscape
}

/// <summary>
/// Named ISO/ANSI page sizes. Each resolves to a concrete portrait dimension in points via
/// <see cref="PageSize.Catalog"/>.
/// </summary>
public enum PageSizeName
{
    /// <summary>ISO A3 (297 × 420 mm).</summary>
    A3,

    /// <summary>ISO A4 (210 × 297 mm) — the default.</summary>
    A4,

    /// <summary>ISO A5 (148 × 210 mm).</summary>
    A5,

    /// <summary>ISO B4 (250 × 353 mm).</summary>
    B4,

    /// <summary>ISO B5 (176 × 250 mm).</summary>
    B5,

    /// <summary>US Letter (8.5 × 11 in).</summary>
    Letter,

    /// <summary>US Legal (8.5 × 14 in).</summary>
    Legal,

    /// <summary>Executive (7.25 × 10.5 in).</summary>
    Executive,

    /// <summary>Statement (5.5 × 8.5 in).</summary>
    Statement,

    /// <summary>Tabloid (11 × 17 in).</summary>
    Tabloid
}

/// <summary>Kind of section break applied before a section (ignored for the first section).</summary>
public enum SectionBreakType
{
    /// <summary>Starts the section on a new page (default for later sections).</summary>
    NextPage,

    /// <summary>Starts the section on the same page (multi-column regions).</summary>
    Continuous,

    /// <summary>Starts the section on the next odd-numbered page.</summary>
    OddPage,

    /// <summary>Starts the section on the next even-numbered page.</summary>
    EvenPage
}

/// <summary>Horizontal alignment of paragraph, cell and table content.</summary>
public enum TextAlignment
{
    /// <summary>Left-aligned (default).</summary>
    Left,

    /// <summary>Centered.</summary>
    Center,

    /// <summary>Right-aligned.</summary>
    Right,

    /// <summary>Justified (both edges flush).</summary>
    Justify
}

/// <summary>List flavor.</summary>
public enum ListKind
{
    /// <summary>Bulleted list (v1: single level, default bullet glyph).</summary>
    Bullet,

    /// <summary>Numbered list (v1: single level, decimal).</summary>
    Ordered
}

/// <summary>How an image is fitted inside its declared box.</summary>
public enum ImageFitMode
{
    /// <summary>Scale to fill the box, cropping overflow (default).</summary>
    Fill,

    /// <summary>Scale to fit inside the box, preserving aspect ratio.</summary>
    Contain,

    /// <summary>Scale and honor the author crop rectangle.</summary>
    Crop,

    /// <summary>Stretch to the exact box dimensions, distorting aspect ratio.</summary>
    Stretch
}

/// <summary>
/// Floating-object text wrap behavior (the DOCX <c>wp:wrap*</c> family). <see
/// cref="BehindText"/> and <see cref="InFrontOfText"/> are z-order treatments, not wraps.
/// </summary>
public enum WrapMode
{
    /// <summary>No wrap; the object sits on top of the text.</summary>
    None,

    /// <summary>Text wraps around the bounding rectangle (default).</summary>
    Square,

    /// <summary>Text wraps around the visible contour.</summary>
    Tight,

    /// <summary>Text wraps around the contour, through open space.</summary>
    Through,

    /// <summary>Text wraps above and below the object, not beside it.</summary>
    TopAndBottom,

    /// <summary>The object sits behind the text.</summary>
    BehindText,

    /// <summary>The object sits in front of the text.</summary>
    InFrontOfText
}

/// <summary>
/// Reference point a positioned object's X/Y offsets are relative to (the DOCX
/// <c>relativeFrom</c> for both axes).
/// </summary>
public enum AnchorReference
{
    /// <summary>Relative to the page edge.</summary>
    Page,

    /// <summary>Relative to the page margin (default).</summary>
    Margin,

    /// <summary>Relative to the column edge.</summary>
    Column,

    /// <summary>Relative to the paragraph origin.</summary>
    Paragraph,

    /// <summary>Relative to the character origin.</summary>
    Character
}

/// <summary>
/// Semantic tone of a report archetype item (KPI value, roadmap phase). Positive/negative
/// tint the item against the theme palette (teal/coral); neutral keeps the theme default.
/// </summary>
public enum ReportTone
{
    /// <summary>On-track / favorable; tints toward the theme teal.</summary>
    Positive,

    /// <summary>Default; keeps the theme's default role color.</summary>
    Neutral,

    /// <summary>At-risk / unfavorable; tints toward the theme coral.</summary>
    Negative
}

/// <summary>Semantic tone of a callout (drives the default fill/emblem treatment).</summary>
public enum CalloutTone
{
    /// <summary>Neutral note (default).</summary>
    Note,

    /// <summary>Positive hint.</summary>
    Tip,

    /// <summary>Cautionary warning.</summary>
    Warning,

    /// <summary>Critical error.</summary>
    Error
}

/// <summary>Axis a straight line runs along inside its bounding box.</summary>
public enum LineOrientation
{
    /// <summary>Left to right.</summary>
    Horizontal,

    /// <summary>Top to bottom.</summary>
    Vertical
}

/// <summary>
/// Semantic text role. A role names what a piece of text <em>is</em> (a title, an eyebrow, a
/// metric) rather than how it should look; the active design theme supplies the default
/// formatting for each role, which content may override with a typography <c>token</c> and
/// direct run formatting. Roles back the generated semantic paragraph styles.
/// </summary>
public enum TextRole
{
    /// <summary>Document/section title.</summary>
    Title,

    /// <summary>Subtitle line under a title.</summary>
    Subtitle,

    /// <summary>Small uppercase kicker above a title.</summary>
    Eyebrow,

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

    /// <summary>Body copy (the default reading role).</summary>
    Body,

    /// <summary>Secondary/supporting body text.</summary>
    Muted,

    /// <summary>Small field label.</summary>
    Label,

    /// <summary>Large standalone number or figure.</summary>
    Metric,

    /// <summary>Caption under a metric.</summary>
    MetricLabel,

    /// <summary>Table header cell text.</summary>
    TableHeader,

    /// <summary>Table body cell text.</summary>
    TableBody,

    /// <summary>Callout/note body text.</summary>
    Callout,

    /// <summary>Page footer text.</summary>
    Footer
}

/// <summary>Helpers for working with <see cref="TextRole"/>.</summary>
public static class TextRoleExtensions
{
    /// <summary>Maps a heading level (1..6) to its heading role.</summary>
    public static TextRole ForHeading(int level) => level switch
    {
        1 => TextRole.Heading1,
        2 => TextRole.Heading2,
        3 => TextRole.Heading3,
        4 => TextRole.Heading4,
        5 => TextRole.Heading5,
        _ => TextRole.Heading6
    };

    /// <summary>True when the role is one of the six heading levels.</summary>
    public static bool IsHeading(this TextRole role) =>
        role is >= TextRole.Heading1 and <= TextRole.Heading6;

    /// <summary>The 1-based outline level of a heading role (throws for non-headings).</summary>
    public static int HeadingLevel(this TextRole role) =>
        role.IsHeading()
            ? (role - TextRole.Heading1) + 1
            : throw new ArgumentOutOfRangeException(nameof(role), role, "only heading roles have a level.");

    /// <summary>True when the role reads as running/reading text that benefits from a body-size
    /// readability guardrail (body, muted, labels, table cells, footer, callout).</summary>
    public static bool IsReadingRole(this TextRole role) => role switch
    {
        TextRole.Body or TextRole.Muted or TextRole.Label or TextRole.MetricLabel or
            TextRole.TableHeader or TextRole.TableBody or TextRole.Callout or TextRole.Footer => true,
        _ => false
    };
}

/// <summary>
/// Typographic density for a document: scales the theme's paragraph spacing roles. Defaults to
/// <see cref="Comfortable"/> (the editorial baseline); <see cref="Compact"/> tightens and
/// <see cref="Spacious"/> loosens.
/// </summary>
public enum Density
{
    /// <summary>Tight spacing (~0.75× the theme baseline).</summary>
    Compact,

    /// <summary>Editorial baseline spacing (1×).</summary>
    Comfortable,

    /// <summary>Looser, airier spacing (~1.4× the theme baseline).</summary>
    Spacious
}

/// <summary>Helpers for working with <see cref="Density"/>.</summary>
public static class DensityExtensions
{
    /// <summary>Deterministic multiplier applied to theme paragraph spacing roles.</summary>
    public static double Scale(this Density density) => density switch
    {
        Density.Compact => 0.75,
        Density.Comfortable => 1.0,
        Density.Spacious => 1.4,
        _ => 1.0
    };
}
