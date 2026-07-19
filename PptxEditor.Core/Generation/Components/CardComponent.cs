using PptxEditor.Core.Generation.Model;

namespace PptxEditor.Core.Generation.Components;

/// <summary>
/// <c>card</c> (plan.md §4): a titled content surface styled by the token card style
/// (<c>shape.cardStyle</c>) and corner radius. Expands to a padded column container with
/// title / subtitle / body texts; the last present field grows to fill the card.
/// <para>Overflow contract: texts shrink (default text policy); structural overflow of
/// the card box errors (default container policy, plan.md §3.2).</para>
/// </summary>
public static class CardComponent
{
    private static readonly IReadOnlySet<string> ContentProps =
        new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "title", "subtitle", "body" };

    /// <summary>Expands <paramref name="element"/> into its primitive subtree.</summary>
    public static ContainerElement Expand(ComponentElement element, DesignTokens design, string path)
    {
        ArgumentNullException.ThrowIfNull(element);
        ArgumentNullException.ThrowIfNull(design);

        var reader = ComponentContentReader.For(path, "card", element.Content, ContentProps, design.Palette);
        var title = reader.String("title", required: true);
        var subtitle = reader.String("subtitle");
        var body = reader.String("body");
        reader.ThrowIfInvalid();

        var metrics = design.Metrics;
        var fields = new List<TextElement>
        {
            ComponentStyle.Line(title!, "display", metrics.BodySizePt + 2, "ink", bold: true)
        };
        if (subtitle is not null)
        {
            fields.Add(ComponentStyle.Line(subtitle, "body", metrics.BodySizePt, "muted"));
        }
        if (body is not null)
        {
            fields.Add(ComponentStyle.Line(body, "body", metrics.BodySizePt, "ink"));
        }
        // The last present field takes the remaining vertical space (single-line fields
        // before it keep their deterministic line height).
        fields[^1] = fields[^1] with { Size = new SizeSpec { Grow = 1 } };

        ComponentStyle.RequireTokens(design, path, "card", "paper", "ink", subtitle is not null ? "muted" : null);

        return ComponentStyle.ApplySurface(new ContainerElement
        {
            Size = element.Size,
            At = element.At,
            Padding = EdgeInsets.All(metrics.GutterPt),
            Layout = new LayoutSpec { Mode = LayoutMode.Column, Gap = 4 },
            Children = fields
        }, design);
    }
}
