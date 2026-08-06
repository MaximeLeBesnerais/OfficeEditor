namespace PptxEditor.Core.Generation.Model;

/// <summary>
/// Root of the generation vocabulary: a from-scratch deck described by
/// a version, design tokens and an ordered list of slides. Each slide's root is a
/// <see cref="ContainerElement"/> sized by <see cref="SlideSize"/>.
/// </summary>
public sealed record GenerationDocument
{
    /// <summary>Vocabulary version. Only "2.0" is supported.</summary>
    public required string Version { get; init; }

    /// <summary>Design tokens (palette, fonts, shape, metrics) referenced by all content.</summary>
    public required DesignTokens Design { get; init; }

    /// <summary>Slides in document order; each root is a container.</summary>
    public required IReadOnlyList<ContainerElement> Slides { get; init; }

    /// <summary>Slide canvas size in points. Defaults to 16:9 (960×540pt).</summary>
    public SlideSize SlideSize { get; init; } = SlideSize.Widescreen16x9;
}

/// <summary>
/// Slide canvas dimensions in points. Covers the two Office presets (16:9 = 960×540,
/// 4:3 = 720×540) plus arbitrary custom sizes, matching the pt units used by element
/// 'at'/'size'. Custom dimensions must lie within [<see cref="MinDimPt"/>, <see cref="MaxDimPt"/>]
/// (PowerPoint's 56-inch ceiling at 72 dpi).
/// </summary>
public readonly record struct SlideSize(double WidthPt, double HeightPt)
{
    /// <summary>Smallest allowed slide dimension in points (1 pt).</summary>
    public const double MinDimPt = 1;

    /// <summary>Largest allowed slide dimension in points (56 in × 72 dpi = 4032 pt, PowerPoint's ceiling).</summary>
    public const double MaxDimPt = 4032;

    /// <summary>13.333in × 7.5in — the v1 default.</summary>
    public static SlideSize Widescreen16x9 => new(960, 540);

    /// <summary>10in × 7.5in.</summary>
    public static SlideSize Standard4x3 => new(720, 540);

    /// <summary>
    /// Creates a slide size from point dimensions, or null when either dimension is not a
    /// finite number within [<see cref="MinDimPt"/>, <see cref="MaxDimPt"/>].
    /// </summary>
    public static SlideSize? TryCreate(double widthPt, double heightPt)
    {
        if (!double.IsFinite(widthPt) || !double.IsFinite(heightPt)
            || widthPt < MinDimPt || widthPt > MaxDimPt
            || heightPt < MinDimPt || heightPt > MaxDimPt)
        {
            return null;
        }
        return new SlideSize(widthPt, heightPt);
    }
}
