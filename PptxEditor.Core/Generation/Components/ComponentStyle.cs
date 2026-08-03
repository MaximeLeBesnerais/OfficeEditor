using PptxEditor.Core.Generation.Model;

namespace PptxEditor.Core.Generation.Components;

/// <summary>
/// Shared styling helpers for the v1 components (design tokens, components).
/// Everything here is expressed in palette token names and font slot names — resolution
/// to hex/families happens exactly once, in the layout resolver.
/// </summary>
internal static class ComponentStyle
{
    /// <summary>
    /// Line-height estimate used to give single-line texts a deterministic box height
    /// (the layout engine sizes boxes, not text). Generous vs. the typical
    /// 1.2–1.35 font pitch so texts do not shrink unnecessarily.
    /// </summary>
    public const double LineHeightFactor = 1.35;

    /// <summary>Estimated height of one text line at <paramref name="fontSizePt"/>, rounded to 3 decimals.</summary>
    public static double LineHeight(double fontSizePt) => Math.Round(fontSizePt * LineHeightFactor, 3);

    /// <summary>Card-surface fill for the token card style: always the "paper" token.</summary>
    public static FillSpec SurfaceFill(DesignTokens design) => new SolidFill("paper");

    /// <summary>Card-surface stroke: "muted" 1 pt for the outline style, none otherwise.</summary>
    public static StrokeSpec? SurfaceStroke(DesignTokens design)
        => design.Shape.CardStyle == CardStyle.Outline ? new StrokeSpec { Color = "muted", WidthPt = 1 } : null;

    /// <summary>Card-surface shadow: a soft "ink" drop shadow for the shadow style, none otherwise.</summary>
    public static ShadowSpec? SurfaceShadow(DesignTokens design)
        => design.Shape.CardStyle == CardStyle.Shadow
            ? new ShadowSpec { Color = "ink", Dx = 0, Dy = 2, Blur = 8, Alpha = 0.18 }
            : null;

    /// <summary>Card-surface corner radius from <c>shape.cornerRadius</c>; null when square.</summary>
    public static CornerRadii? SurfaceRadius(DesignTokens design)
        => design.Shape.CornerRadius > 0 ? CornerRadii.All(design.Shape.CornerRadius) : null;

    /// <summary>Applies the token card-surface treatment (fill/stroke/shadow/radius) to a container.</summary>
    public static ContainerElement ApplySurface(ContainerElement container, DesignTokens design)
        => container with
        {
            Fill = SurfaceFill(design),
            Stroke = SurfaceStroke(design),
            Shadow = SurfaceShadow(design),
            Radius = SurfaceRadius(design)
        };

    /// <summary>
    /// Verifies every palette token a component references exists, listing all missing
    /// ones in one loud error. Component defaults reference tokens by name; without this
    /// check the failure would surface later as a less actionable layout error.
    /// </summary>
    public static void RequireTokens(DesignTokens design, string path, string componentName, params string?[] tokens)
    {
        var missing = tokens
            .Where(t => t is not null)
            .Distinct()
            .Where(t => !design.Palette.ContainsKey(t!))
            .ToList();
        if (missing.Count > 0)
        {
            throw new ComponentException(path,
                $"component '{componentName}' requires palette token(s) {string.Join(", ", missing.Select(t => $"'{t}'"))} " +
                $"(declared tokens: {string.Join(", ", design.Palette.Keys)}).");
        }
    }

    /// <summary>
    /// One single-line text element with a deterministic box height (the line estimate).
    /// Overflow stays at the text default — shrink.
    /// </summary>
    public static TextElement Line(string value, string font, double fontSizePt, string color, bool bold = false, TextAlign align = TextAlign.Left)
        => new()
        {
            Value = value,
            Font = font,
            FontSize = fontSizePt,
            Color = color,
            Bold = bold,
            TextAlign = align,
            Size = new SizeSpec { Height = LineHeight(fontSizePt) }
        };

    /// <summary>One text element that grows into the remaining space of its layout parent.</summary>
    public static TextElement Growing(string value, string font, double fontSizePt, string color, bool bold = false, TextAlign align = TextAlign.Left, TextAnchor anchor = TextAnchor.Top)
        => new()
        {
            Value = value,
            Font = font,
            FontSize = fontSizePt,
            Color = color,
            Bold = bold,
            TextAlign = align,
            Anchor = anchor,
            Size = new SizeSpec { Grow = 1 }
        };
}
