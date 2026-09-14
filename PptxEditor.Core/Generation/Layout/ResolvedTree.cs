using PptxEditor.Core.Generation.Model;
using PptxEditor.Core.Models;

namespace PptxEditor.Core.Generation.Layout;

/// <summary>
/// One resolved slide: the absolute draw tree both emitters consume —
/// layout is resolved exactly once, in C#. All geometry in points; all fill/stroke/shadow
/// colors resolved to #RRGGBB; all font slots resolved to family names.
/// </summary>
public sealed record ResolvedSlide
{
    /// <summary>Slide canvas width in points.</summary>
    public required double WidthPt { get; init; }

    /// <summary>Slide canvas height in points.</summary>
    public required double HeightPt { get; init; }

    /// <summary>Resolved root container (covers the whole slide).</summary>
    public required ResolvedContainer Root { get; init; }

    /// <summary>
    /// Speaker notes for the slide; null = none. Carried through layout as pure metadata:
    /// the OOXML emitter writes a NotesSlidePart, the Typst preview intentionally omits it.
    /// </summary>
    public string? Notes { get; init; }

    /// <summary>Stable element id of the slide root; null when the deck authors none.</summary>
    public string? Id { get; init; }

    /// <summary>
    /// Flat, paint-order view of every resolved element on the slide (depth-first; group
    /// children are already in absolute coordinates and appear between their group and the
    /// next sibling). Built from the single layout pass — no extra geometry computation —
    /// for hit-testing, selection boxes and id-based addressing in Studio.
    /// </summary>
    public required IReadOnlyList<ResolvedElementInfo> Elements { get; init; }

    /// <summary>
    /// Resolves the absolute rect (x, y, w, h in points) of the element addressed by
    /// <paramref name="id"/> on this slide. False when the id is unknown here. Linear scan
    /// over <see cref="Elements"/> — callers re-querying the same id should cache the result.
    /// </summary>
    public bool TryGetRect(string? id, out ElementRect rect)
    {
        if (id is not null)
        {
            foreach (var element in Elements)
            {
                if (string.Equals(element.Id, id, StringComparison.Ordinal))
                {
                    rect = new ElementRect(element.X, element.Y, element.Width, element.Height);
                    return true;
                }
            }
        }
        rect = default;
        return false;
    }
}

/// <summary>Result of resolving a whole document: one draw tree per slide plus non-fatal warnings.</summary>
public sealed record LayoutResult
{
    /// <summary>Resolved slides in document order.</summary>
    public required IReadOnlyList<ResolvedSlide> Slides { get; init; }

    /// <summary>Non-fatal diagnostics (e.g. text shrunk below MinScale).</summary>
    public required IReadOnlyList<string> Warnings { get; init; }

    /// <summary>
    /// Resolves the absolute rect of the element addressed by <paramref name="id"/> on the
    /// given slide; false when the slide index or id is unknown. Convenience over
    /// <see cref="ResolvedSlide.TryGetRect"/> — both ride the existing layout pass.
    /// </summary>
    public bool TryGetRect(int slideIndex, string? id, out ElementRect rect)
    {
        if ((uint)slideIndex < (uint)Slides.Count)
        {
            return Slides[slideIndex].TryGetRect(id, out rect);
        }
        rect = default;
        return false;
    }
}

/// <summary>
/// Kind of a resolved element, so introspection callers can switch on the shape without
/// pattern-matching every record type.
/// </summary>
public enum ResolvedElementType
{
    /// <summary>Layout container (root, card surface, row …).</summary>
    Container,

    /// <summary>Text primitive.</summary>
    Text,

    /// <summary>Rectangle primitive.</summary>
    Rect,

    /// <summary>Ellipse primitive.</summary>
    Ellipse,

    /// <summary>Straight line or connector.</summary>
    Line,

    /// <summary>Image primitive.</summary>
    Image,

    /// <summary>Group: paint-order children, already placed in absolute coordinates.</summary>
    Group
}

/// <summary>
/// One entry in a slide's flat paint-order element list (<see cref="ResolvedSlide.Elements"/>):
/// the element's stable id (null when the deck authors none), its kind and its final
/// absolute rect in points. Value projection of the resolved tree — no reference to the
/// underlying node, so callers can cache the list freely.
/// </summary>
public sealed record ResolvedElementInfo
{
    /// <summary>Stable element id from the source JSON; null when none was authored.</summary>
    public string? Id { get; init; }

    /// <summary>Element kind, for hit-testing switches.</summary>
    public ResolvedElementType Type { get; init; }

    /// <summary>Absolute X in points.</summary>
    public double X { get; init; }

    /// <summary>Absolute Y in points.</summary>
    public double Y { get; init; }

    /// <summary>Width in points.</summary>
    public double Width { get; init; }

    /// <summary>Height in points.</summary>
    public double Height { get; init; }
}

/// <summary>An axis-aligned rectangle in points (final layout geometry).</summary>
public readonly record struct ElementRect(double X, double Y, double Width, double Height);

/// <summary>
/// Base of every node in the absolute draw tree. Geometry is the final position in points,
/// rounded to 3 decimals; children of containers are positioned inside this rect.
/// </summary>
public abstract record ResolvedElement
{
    /// <summary>Absolute X in points.</summary>
    public double X { get; init; }

    /// <summary>Absolute Y in points.</summary>
    public double Y { get; init; }

    /// <summary>Width in points.</summary>
    public double Width { get; init; }

    /// <summary>Height in points.</summary>
    public double Height { get; init; }

    /// <summary>
    /// Stable element id from the source JSON, carried through expansion and layout;
    /// null when none was authored. Metadata only — never rendered.
    /// </summary>
    public string? Id { get; init; }
}

/// <summary>Resolved container: background/border plus its resolved children in paint order.</summary>
public sealed record ResolvedContainer : ResolvedElement
{
    /// <summary>Background fill with colors resolved to hex; null = none.</summary>
    public FillSpec? Fill { get; init; }

    /// <summary>Border stroke with color resolved to hex.</summary>
    public StrokeSpec? Stroke { get; init; }

    /// <summary>Per-corner radius in points.</summary>
    public CornerRadii? Radius { get; init; }

    /// <summary>Drop shadow with color resolved to hex.</summary>
    public ShadowSpec? Shadow { get; init; }

    /// <summary>The container's overflow policy (plumbing for emitters, e.g. clip).</summary>
    public OverflowPolicy Overflow { get; init; }

    /// <summary>Resolved children in document order — which is also the paint order.</summary>
    public required IReadOnlyList<ResolvedElement> Children { get; init; }
}

/// <summary>One text run with every default resolved (font family, size, hex color, weight, style).</summary>
public sealed record ResolvedTextRun
{
    /// <summary>Run text.</summary>
    public required string Text { get; init; }

    /// <summary>Resolved font family name; null = emitter default.</summary>
    public required string? FontFamily { get; init; }

    /// <summary>Resolved font size in points (before <see cref="ResolvedText.FontScale"/>).</summary>
    public required double FontSizePt { get; init; }

    /// <summary>Resolved #RRGGBB color; null = emitter default.</summary>
    public required string? ColorHex { get; init; }

    /// <summary>Effective bold weight.</summary>
    public required bool Bold { get; init; }

    /// <summary>Effective italic style.</summary>
    public required bool Italic { get; init; }
}

/// <summary>
/// Resolved text primitive: box + anchor + align + insets. Content is normalized to runs
/// (a plain "text" value becomes a single run with the element defaults applied).
/// </summary>
public sealed record ResolvedText : ResolvedElement
{
    /// <summary>Resolved runs with all defaults applied.</summary>
    public required IReadOnlyList<ResolvedTextRun> Runs { get; init; }

    /// <summary>Horizontal text alignment inside the box.</summary>
    public TextAlign TextAlign { get; init; }

    /// <summary>Vertical anchor inside the box.</summary>
    public TextAnchor Anchor { get; init; }

    /// <summary>Text box insets in points.</summary>
    public EdgeInsets Insets { get; init; }

    /// <summary>The element's overflow policy.</summary>
    public OverflowPolicy Overflow { get; init; }

    /// <summary>
    /// Font scale from the shrink pass (1 = no shrink). Filled by the
    /// <see cref="ITextMeasurer"/> seam (P3); stays 1 when no measurer is wired.
    /// </summary>
    public double FontScale { get; init; } = 1;

    /// <summary>Drop shadow with color resolved to hex.</summary>
    public ShadowSpec? Shadow { get; init; }
}

/// <summary>Resolved rectangle primitive.</summary>
public sealed record ResolvedRect : ResolvedElement
{
    /// <summary>Fill with colors resolved to hex; null = no fill.</summary>
    public FillSpec? Fill { get; init; }

    /// <summary>Border stroke with color resolved to hex.</summary>
    public StrokeSpec? Stroke { get; init; }

    /// <summary>Per-corner radius in points.</summary>
    public CornerRadii? Radius { get; init; }

    /// <summary>Drop shadow with color resolved to hex.</summary>
    public ShadowSpec? Shadow { get; init; }
}

/// <summary>Resolved ellipse primitive.</summary>
public sealed record ResolvedEllipse : ResolvedElement
{
    /// <summary>Fill with colors resolved to hex; null = no fill.</summary>
    public FillSpec? Fill { get; init; }

    /// <summary>Border stroke with color resolved to hex.</summary>
    public StrokeSpec? Stroke { get; init; }

    /// <summary>Drop shadow with color resolved to hex.</summary>
    public ShadowSpec? Shadow { get; init; }
}

/// <summary>Resolved straight line/connector inside its box.</summary>
public sealed record ResolvedLine : ResolvedElement
{
    /// <summary>True when authored as "connector" (OOXML cxnSp), false for "line".</summary>
    public bool IsConnector { get; init; }

    /// <summary>Line direction inside the box.</summary>
    public LineOrientation Orientation { get; init; }

    /// <summary>Stroke with color resolved to hex; null = emitter default.</summary>
    public StrokeSpec? Stroke { get; init; }
}

/// <summary>Resolved image primitive.</summary>
public sealed record ResolvedImage : ResolvedElement
{
    /// <summary>Image source (path, URL or base64 payload — interpreted by emitters).</summary>
    public required string Source { get; init; }

    /// <summary>Fit mode: fill / crop / contain.</summary>
    public ImageFitMode Fit { get; init; }

    /// <summary>Caller-supplied source crop (only meaningful with <see cref="ImageFitMode.Crop"/>).</summary>
    public SourceRect? Crop { get; init; }

    /// <summary>Accessibility alt text.</summary>
    public string? Alt { get; init; }
}

/// <summary>Resolved group: children in paint order, already placed in absolute coordinates.</summary>
public sealed record ResolvedGroup : ResolvedElement
{
    /// <summary>Resolved children in document order — which is also the paint order.</summary>
    public required IReadOnlyList<ResolvedElement> Children { get; init; }
}
