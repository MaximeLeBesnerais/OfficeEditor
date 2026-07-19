namespace PptxEditor.Core.Generation.Model;

/// <summary>
/// Child size constraints (plan.md §3.2): fixed dimensions, grow share, aspect ratio and
/// cross-axis self-alignment. All dimensions in points.
/// </summary>
public sealed record SizeSpec
{
    /// <summary>Fixed width in points (&gt; 0).</summary>
    public double? Width { get; init; }

    /// <summary>Fixed height in points (&gt; 0).</summary>
    public double? Height { get; init; }

    /// <summary>
    /// Share of the remaining space on the parent's layout axis after fixed children are
    /// subtracted (≥ 0). Only meaningful inside a layout container.
    /// </summary>
    public double? Grow { get; init; }

    /// <summary>Aspect constraint ("16:9"); resolves against the dimension the parent constrains first.</summary>
    public AspectRatio? Aspect { get; init; }

    /// <summary>Per-child cross-axis alignment override.</summary>
    public AlignItems? AlignSelf { get; init; }
}

/// <summary>Width:height ratio parsed from "W:H" (e.g. "16:9" → 16/9).</summary>
public readonly record struct AspectRatio(double Value)
{
    /// <summary>Builds a ratio from its two components; both must be positive.</summary>
    public static AspectRatio Of(double width, double height)
    {
        if (width <= 0 || height <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(width), "Aspect ratio components must be positive.");
        }
        return new AspectRatio(width / height);
    }
}

/// <summary>
/// Per-container overflow policy (plan.md §3.2). Defaults: <see cref="Shrink"/> for
/// text-bearing leaves, <see cref="Error"/> for layout containers.
/// </summary>
public enum OverflowPolicy
{
    /// <summary>Loud failure on overflow (default for layout containers).</summary>
    Error,

    /// <summary>Shrink text via TextFitService (fontScale ≥ MinScale, else warning).</summary>
    Shrink,

    /// <summary>Clip at the box edge.</summary>
    Clip
}

/// <summary>Absolute placement escape hatch (plan.md §3.2): only on children of layout-less parents.</summary>
public readonly record struct PointSpec(double X, double Y);
