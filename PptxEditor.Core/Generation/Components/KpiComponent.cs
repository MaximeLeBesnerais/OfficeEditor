using PptxEditor.Core.Generation.Model;

namespace PptxEditor.Core.Generation.Components;

/// <summary>
/// <c>kpi</c> (plan.md §4): a big number, its label and an optional accent delta,
/// centered on a card surface. Expands to a centered column container; the value renders
/// in the display font at <c>metrics.titleSizePt</c> in the "primary" token color.
/// <para>Overflow contract: all texts shrink (the value especially — plan.md §7.5);
/// structural overflow of the surface errors.</para>
/// </summary>
public static class KpiComponent
{
    private static readonly IReadOnlySet<string> ContentProps =
        new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "value", "label", "delta" };

    /// <summary>Expands <paramref name="element"/> into its primitive subtree.</summary>
    public static ContainerElement Expand(ComponentElement element, DesignTokens design, string path)
    {
        ArgumentNullException.ThrowIfNull(element);
        ArgumentNullException.ThrowIfNull(design);

        var reader = ComponentContentReader.For(path, "kpi", element.Content, ContentProps, design.Palette);
        var value = reader.String("value", required: true);
        var label = reader.String("label", required: true);
        var delta = reader.String("delta");
        reader.ThrowIfInvalid();

        var metrics = design.Metrics;
        var children = new List<GenElement>
        {
            ComponentStyle.Line(value!, "display", metrics.TitleSizePt, "primary", bold: true, align: TextAlign.Center),
            ComponentStyle.Line(label!, "body", metrics.BodySizePt, "muted", align: TextAlign.Center)
        };
        if (delta is not null)
        {
            children.Add(ComponentStyle.Line(delta, "body", metrics.BodySizePt - 2, "accent", align: TextAlign.Center));
        }

        ComponentStyle.RequireTokens(design, path, "kpi", "paper", "primary", "muted", delta is not null ? "accent" : null);

        return ComponentStyle.ApplySurface(new ContainerElement
        {
            Size = element.Size,
            At = element.At,
            Padding = EdgeInsets.All(metrics.GutterPt),
            Layout = new LayoutSpec { Mode = LayoutMode.Column, Gap = 2, Justify = Justify.Center, Align = AlignItems.Center },
            Children = children
        }, design);
    }
}
