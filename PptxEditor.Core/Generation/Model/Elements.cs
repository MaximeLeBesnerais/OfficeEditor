using System.Text.Json;
using PptxEditor.Core.Models;

namespace PptxEditor.Core.Generation.Model;

/// <summary>
/// Base of every node in the slide tree (size and overflow semantics). <see cref="Size"/> constrains
/// the node inside a layout container; <see cref="At"/> is the absolute-placement escape
/// hatch, allowed only on children of layout-less parents.
/// </summary>
public abstract record GenElement
{
    /// <summary>Optional size constraints. Invalid on the root slide container.</summary>
    public SizeSpec? Size { get; init; }

    /// <summary>Absolute placement in points. Only valid on children of layout-less parents.</summary>
    public PointSpec? At { get; init; }
}

/// <summary>
/// Layout container. With <see cref="Layout"/> it resolves children via
/// row/column/grid; without it, it is a free canvas whose children must use
/// <see cref="GenElement.At"/>. The root of every slide is a container.
/// </summary>
public sealed record ContainerElement : GenElement
{
    /// <summary>Layout declaration; null = layout-less free canvas.</summary>
    public LayoutSpec? Layout { get; init; }

    /// <summary>Inner padding in points.</summary>
    public EdgeInsets? Padding { get; init; }

    /// <summary>Overflow policy. Defaults to <see cref="OverflowPolicy.Error"/>.</summary>
    public OverflowPolicy Overflow { get; init; } = OverflowPolicy.Error;

    /// <summary>Children in document order — which is also the paint order (no z-index).</summary>
    public required IReadOnlyList<GenElement> Children { get; init; }

    /// <summary>Background fill (e.g. slide paper color, styled card surface).</summary>
    public FillSpec? Fill { get; init; }

    /// <summary>Border stroke.</summary>
    public StrokeSpec? Stroke { get; init; }

    /// <summary>Per-corner radius in points.</summary>
    public CornerRadii? Radius { get; init; }

    /// <summary>Drop shadow.</summary>
    public ShadowSpec? Shadow { get; init; }

    /// <summary>
    /// Speaker notes for the slide; metadata only, never rendered. Valid only on the
    /// slide root (the parser rejects it on nested containers); the OOXML emitter writes
    /// it into a NotesSlidePart, the Typst preview deliberately ignores it.
    /// </summary>
    public string? Notes { get; init; }
}

/// <summary>
/// Text primitive: box + anchor + align + insets. Exactly one of
/// <see cref="Value"/> / <see cref="Runs"/>. <see cref="TextAlign"/> lives here — never
/// confuse it with container align.
/// </summary>
public sealed record TextElement : GenElement
{
    /// <summary>Plain text content ("text" in JSON).</summary>
    public string? Value { get; init; }

    /// <summary>Styled runs ("runs" in JSON); mutually exclusive with <see cref="Value"/>.</summary>
    public IReadOnlyList<TextRun>? Runs { get; init; }

    /// <summary>Font slot token ("display" | "body") or a raw family name.</summary>
    public string? Font { get; init; }

    /// <summary>Font size in points (&gt; 0).</summary>
    public double? FontSize { get; init; }

    /// <summary>Palette token name or #RRGGBB literal.</summary>
    public string? Color { get; init; }

    /// <summary>Bold weight.</summary>
    public bool Bold { get; init; }

    /// <summary>Italic style.</summary>
    public bool Italic { get; init; }

    /// <summary>Horizontal text alignment inside the box. Defaults to <see cref="TextAlign.Left"/>.</summary>
    public TextAlign TextAlign { get; init; } = TextAlign.Left;

    /// <summary>Vertical anchor inside the box. Defaults to <see cref="TextAnchor.Top"/>.</summary>
    public TextAnchor Anchor { get; init; } = TextAnchor.Top;

    /// <summary>Text box insets in points.</summary>
    public EdgeInsets? Insets { get; init; }

    /// <summary>Overflow policy. Defaults to <see cref="OverflowPolicy.Shrink"/>.</summary>
    public OverflowPolicy Overflow { get; init; } = OverflowPolicy.Shrink;

    /// <summary>Drop shadow.</summary>
    public ShadowSpec? Shadow { get; init; }
}

/// <summary>One styled run inside a <see cref="TextElement"/>.</summary>
public sealed record TextRun
{
    /// <summary>Run text. Required.</summary>
    public required string Text { get; init; }

    /// <summary>Font slot token ("display" | "body") or a raw family name.</summary>
    public string? Font { get; init; }

    /// <summary>Font size in points (&gt; 0).</summary>
    public double? FontSize { get; init; }

    /// <summary>Palette token name or #RRGGBB literal.</summary>
    public string? Color { get; init; }

    /// <summary>Bold weight.</summary>
    public bool Bold { get; init; }

    /// <summary>Italic style.</summary>
    public bool Italic { get; init; }
}

/// <summary>Horizontal text alignment inside a text box.</summary>
public enum TextAlign
{
    /// <summary>Left-aligned.</summary>
    Left,

    /// <summary>Centered.</summary>
    Center,

    /// <summary>Right-aligned.</summary>
    Right
}

/// <summary>Vertical anchor of text inside its box.</summary>
public enum TextAnchor
{
    /// <summary>Anchored to the top.</summary>
    Top,

    /// <summary>Anchored to the middle.</summary>
    Middle,

    /// <summary>Anchored to the bottom.</summary>
    Bottom
}

/// <summary>Rectangle primitive with per-corner radius.</summary>
public sealed record RectElement : GenElement
{
    /// <summary>Fill; null = no fill.</summary>
    public FillSpec? Fill { get; init; }

    /// <summary>Border stroke.</summary>
    public StrokeSpec? Stroke { get; init; }

    /// <summary>Per-corner radius in points (OOXML round1Rect/round2SameRect ↔ Typst rect radius).</summary>
    public CornerRadii? Radius { get; init; }

    /// <summary>Drop shadow.</summary>
    public ShadowSpec? Shadow { get; init; }
}

/// <summary>Ellipse primitive.</summary>
public sealed record EllipseElement : GenElement
{
    /// <summary>Fill; null = no fill.</summary>
    public FillSpec? Fill { get; init; }

    /// <summary>Border stroke.</summary>
    public StrokeSpec? Stroke { get; init; }

    /// <summary>Drop shadow.</summary>
    public ShadowSpec? Shadow { get; init; }
}

/// <summary>Straight line or connector (straight only in v1).</summary>
public sealed record LineElement : GenElement
{
    /// <summary>True when authored as "connector" (OOXML cxnSp), false for "line".</summary>
    public bool IsConnector { get; init; }

    /// <summary>Line direction inside its resolved box. Defaults to <see cref="LineOrientation.Horizontal"/>.</summary>
    public LineOrientation Orientation { get; init; } = LineOrientation.Horizontal;

    /// <summary>Stroke; null = emitter default.</summary>
    public StrokeSpec? Stroke { get; init; }
}

/// <summary>Direction of a straight line inside its resolved box.</summary>
public enum LineOrientation
{
    /// <summary>Left to right.</summary>
    Horizontal,

    /// <summary>Top to bottom.</summary>
    Vertical
}

/// <summary>Image primitive with fit modes from F7.</summary>
public sealed record ImageElement : GenElement
{
    /// <summary>Image source (path, URL or base64 payload — interpreted by emitters). Required.</summary>
    public required string Source { get; init; }

    /// <summary>Fit mode: fill / crop / contain (no stretch in the generation vocabulary). Defaults to fill.</summary>
    public ImageFitMode Fit { get; init; } = ImageFitMode.Fill;

    /// <summary>Caller-supplied source crop (only meaningful with <see cref="ImageFitMode.Crop"/>).</summary>
    public SourceRect? Crop { get; init; }

    /// <summary>Accessibility alt text.</summary>
    public string? Alt { get; init; }
}

/// <summary>
/// Group primitive: layout-less container whose children paint in document
/// order. Children are placed with <see cref="GenElement.At"/> in group coordinates.
/// </summary>
public sealed record GroupElement : GenElement
{
    /// <summary>Children in document order — which is also the paint order.</summary>
    public required IReadOnlyList<GenElement> Children { get; init; }
}

/// <summary>
/// Prebuilt component node (card, kpi, title_block, bullet_list, divider,
/// badge, image_card, table_block). P1 carries it as a name plus raw content bag; the
/// component implementations (P6) are C# functions over primitives, not a second layout
/// system.
/// </summary>
public sealed record ComponentElement : GenElement
{
    /// <summary>Known v1 component names (max 8).</summary>
    public static readonly IReadOnlySet<string> KnownNames = new HashSet<string>(StringComparer.Ordinal)
    {
        "card", "kpi", "title_block", "bullet_list", "divider", "badge", "image_card", "table_block"
    };

    /// <summary>Component name, one of <see cref="KnownNames"/>.</summary>
    public required string Name { get; init; }

    /// <summary>Raw component content (cloned JSON); strongly typed by the component layer.</summary>
    public JsonElement? Content { get; init; }
}
