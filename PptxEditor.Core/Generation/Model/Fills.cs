namespace PptxEditor.Core.Generation.Model;

/// <summary>
/// Fill of a shape (plan.md §3.3): a solid color reference or a linear gradient.
/// JSON: a color string (palette token or #RRGGBB) for <see cref="SolidFill"/>, or an
/// object with "angle" + "stops" for <see cref="LinearGradientFill"/>.
/// </summary>
public abstract record FillSpec
{
    protected FillSpec() { }
}

/// <summary>Solid fill. <see cref="Color"/> is a palette token name or a validated #RRGGBB literal.</summary>
public sealed record SolidFill(string Color) : FillSpec;

/// <summary>Linear gradient fill — linear only in v1 (Tier-3 gradients are a non-goal, plan.md §1).</summary>
public sealed record LinearGradientFill : FillSpec
{
    /// <summary>Gradient axis angle in degrees (OOXML a:lin ang ↔ Typst gradient.linear angle).</summary>
    public required double Angle { get; init; }

    /// <summary>Color stops, at least 2, offsets in [0, 1].</summary>
    public required IReadOnlyList<GradientStop> Stops { get; init; }
}

/// <summary>One gradient stop: color reference, offset in [0, 1], optional alpha in [0, 1].</summary>
public sealed record GradientStop
{
    /// <summary>Palette token name or #RRGGBB literal.</summary>
    public required string Color { get; init; }

    /// <summary>Stop position along the gradient axis, 0–1.</summary>
    public required double Offset { get; init; }

    /// <summary>Optional opacity multiplier, 0–1.</summary>
    public double? Alpha { get; init; }
}

/// <summary>Outline of a shape or line. JSON: an object {"color", "width"} or a bare color string (width 1pt).</summary>
public sealed record StrokeSpec
{
    /// <summary>Palette token name or #RRGGBB literal. Required.</summary>
    public required string Color { get; init; }

    /// <summary>Stroke width in points (&gt; 0). Defaults to 1.</summary>
    public double WidthPt { get; init; } = 1;
}

/// <summary>
/// Drop shadow (plan.md §3.3): native a:effectLst/outerShdw in OOXML, faked offset rect
/// in the Typst preview. Never a preview trick inside the PPTX (plan.md §2 rule 4).
/// </summary>
public sealed record ShadowSpec
{
    /// <summary>Palette token name or #RRGGBB literal. Required.</summary>
    public required string Color { get; init; }

    /// <summary>Horizontal offset in points.</summary>
    public double Dx { get; init; }

    /// <summary>Vertical offset in points.</summary>
    public double Dy { get; init; }

    /// <summary>Blur radius in points (≥ 0).</summary>
    public double Blur { get; init; }

    /// <summary>Optional opacity multiplier, 0–1.</summary>
    public double? Alpha { get; init; }
}

/// <summary>
/// Per-corner radius in points (plan.md §3.3 rect; §8 Q3 object form). JSON: a single
/// number (all corners) or {"tl", "tr", "br", "bl"} — missing corners default to 0.
/// Maps to OOXML round1Rect/round2SameRect adj values ↔ Typst rect radius corners.
/// </summary>
public readonly record struct CornerRadii(double TopLeft, double TopRight, double BottomRight, double BottomLeft)
{
    /// <summary>Uniform radius on all four corners.</summary>
    public static CornerRadii All(double value) => new(value, value, value, value);
}
