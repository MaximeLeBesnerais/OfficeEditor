namespace DocxEditor.Core.Generation.Model;

/// <summary>
/// Base of every positioned primitive in the positioned tier: absolutely anchored content
/// (free text boxes, shapes, lines, floating pictures, callouts) that floats over a
/// section's flow content. <see cref="Position"/> carries geometry, rotation, z-order,
/// anchor reference, wrap and alt text. z-order paints low-to-high: larger
/// <see cref="PositionSpec.ZOrder"/> values render on top.
/// </summary>
public abstract record PositionedElement
{
    /// <summary>Position and sizing metadata. Required.</summary>
    public required PositionSpec Position { get; init; }
}

/// <summary>
/// Geometry of a positioned primitive. <see cref="X"/>/<see cref="Y"/> are the offsets
/// from the <see cref="Anchor"/> reference; box primitives (text box, image, rectangle,
/// callout) require both <see cref="WidthPt"/> and <see cref="HeightPt"/>, a line requires
/// only <see cref="WidthPt"/> (its length along <see cref="LineOrientation"/>).
/// </summary>
public sealed record PositionSpec
{
    /// <summary>Horizontal offset from the anchor reference in points (≥ 0).</summary>
    public required double X { get; init; }

    /// <summary>Vertical offset from the anchor reference in points (≥ 0).</summary>
    public required double Y { get; init; }

    /// <summary>Box width in points (&gt; 0).</summary>
    public double? WidthPt { get; init; }

    /// <summary>Box height in points (&gt; 0).</summary>
    public double? HeightPt { get; init; }

    /// <summary>Clockwise rotation in degrees (default 0).</summary>
    public double Rotation { get; init; }

    /// <summary>Paint order within the section: higher renders on top (default 0).</summary>
    public int ZOrder { get; init; }

    /// <summary>Reference point the offsets are relative to. Defaults to <see cref="AnchorReference.Margin"/>.</summary>
    public AnchorReference Anchor { get; init; } = AnchorReference.Margin;

    /// <summary>Text wrap behavior. Defaults to <see cref="WrapMode.Square"/>.</summary>
    public WrapMode Wrap { get; init; } = WrapMode.Square;

    /// <summary>Optional wrap distances (only meaningful with wrap modes that push text).</summary>
    public WrapDistances? WrapDistances { get; init; }

    /// <summary>Accessibility alt text for the object.</summary>
    public string? Alt { get; init; }
}

/// <summary>Distance between a floating object and the text that wraps around it, in points.</summary>
public sealed record WrapDistances
{
    /// <summary>Top distance (≥ 0).</summary>
    public double TopPt { get; init; }

    /// <summary>Left distance (≥ 0).</summary>
    public double LeftPt { get; init; }

    /// <summary>Bottom distance (≥ 0).</summary>
    public double BottomPt { get; init; }

    /// <summary>Right distance (≥ 0).</summary>
    public double RightPt { get; init; }
}

/// <summary>A stroke (line/border) definition: a color plus a width in points.</summary>
public sealed record StrokeSpec
{
    /// <summary>Stroke color: palette token name or #RRGGBB literal. Required.</summary>
    public required string Color { get; init; }

    /// <summary>Stroke width in points (&gt; 0). Defaults to 1.</summary>
    public double WidthPt { get; init; } = 1;
}

/// <summary>A free text box: anchored text with optional fill, stroke and corner radius.</summary>
public sealed record PositionedTextBox : PositionedElement
{
    /// <summary>Text box content. Required.</summary>
    public required TextModel Content { get; init; }

    /// <summary>Box fill: palette token or hex. Null = no fill.</summary>
    public string? Fill { get; init; }

    /// <summary>Box border. Null = no border.</summary>
    public StrokeSpec? Stroke { get; init; }

    /// <summary>Corner radius in points. Null = design <c>shapes.cornerRadius</c> default.</summary>
    public double? CornerRadiusPt { get; init; }
}

/// <summary>A floating picture (anchored image) in the positioned tier.</summary>
public sealed record PositionedImage : PositionedElement
{
    /// <summary>Image source (path, URL or base64 payload). Required.</summary>
    public required string Source { get; init; }

    /// <summary>Fit mode. Defaults to <see cref="ImageFitMode.Fill"/>.</summary>
    public ImageFitMode Fit { get; init; } = ImageFitMode.Fill;

    /// <summary>Optional source crop (0..1 fractions).</summary>
    public ImageCrop? Crop { get; init; }
}

/// <summary>An anchored rectangle with optional fill, stroke and corner radius.</summary>
public sealed record PositionedRectangle : PositionedElement
{
    /// <summary>Fill: palette token or hex. Null = no fill.</summary>
    public string? Fill { get; init; }

    /// <summary>Border. Null = no border.</summary>
    public StrokeSpec? Stroke { get; init; }

    /// <summary>Corner radius in points. Null = design default.</summary>
    public double? CornerRadiusPt { get; init; }
}

/// <summary>A straight line along <see cref="Orientation"/>; its length is
/// <see cref="PositionSpec.WidthPt"/>.</summary>
public sealed record PositionedLine : PositionedElement
{
    /// <summary>Line axis. Defaults to <see cref="LineOrientation.Horizontal"/>.</summary>
    public LineOrientation Orientation { get; init; } = LineOrientation.Horizontal;

    /// <summary>Stroke. Null = emitter default.</summary>
    public StrokeSpec? Stroke { get; init; }
}

/// <summary>A free-floating callout box: anchored text with tone, optional fill and stroke.</summary>
public sealed record PositionedCallout : PositionedElement
{
    /// <summary>Callout tone (drives the default visual treatment).</summary>
    public CalloutTone Tone { get; init; } = CalloutTone.Note;

    /// <summary>Callout text content. Required.</summary>
    public required TextModel Content { get; init; }

    /// <summary>Box fill: palette token or hex. Null = no fill.</summary>
    public string? Fill { get; init; }

    /// <summary>Box border. Null = no border.</summary>
    public StrokeSpec? Stroke { get; init; }

    /// <summary>Corner radius in points. Null = design default.</summary>
    public double? CornerRadiusPt { get; init; }
}
