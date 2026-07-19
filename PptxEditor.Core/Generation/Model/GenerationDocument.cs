namespace PptxEditor.Core.Generation.Model;

/// <summary>
/// Root of the generation vocabulary (plan.md §3.4): a from-scratch deck described by
/// a version, design tokens and an ordered list of slides. Each slide's root is a
/// <see cref="ContainerElement"/> sized by <see cref="SlideSize"/>.
/// </summary>
public sealed record GenerationDocument
{
    /// <summary>Vocabulary version. Only "2.0" is supported.</summary>
    public required string Version { get; init; }

    /// <summary>Design tokens (palette, fonts, shape, metrics) referenced by all content.</summary>
    public required DesignTokens Design { get; init; }

    /// <summary>Slides in document order; each root is a container (plan.md §3.2).</summary>
    public required IReadOnlyList<ContainerElement> Slides { get; init; }

    /// <summary>Slide canvas size in points. Defaults to 16:9 (960×540pt).</summary>
    public SlideSize SlideSize { get; init; } = SlideSize.Widescreen16x9;
}

/// <summary>Slide canvas dimensions in points (v1: the two Office presets, plan.md §8 Q4).</summary>
public readonly record struct SlideSize(double WidthPt, double HeightPt)
{
    /// <summary>13.333in × 7.5in — the v1 default.</summary>
    public static SlideSize Widescreen16x9 => new(960, 540);

    /// <summary>10in × 7.5in.</summary>
    public static SlideSize Standard4x3 => new(720, 540);
}
