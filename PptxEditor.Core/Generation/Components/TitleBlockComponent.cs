using PptxEditor.Core.Generation.Model;

namespace PptxEditor.Core.Generation.Components;

/// <summary>
/// <c>title_block</c> (component layer): a slide heading group — optional accent kicker over
/// the display-font title over an optional muted subtitle. Expands to a plain column
/// container (no surface); slack stays at the bottom (justify start).
/// <para>Overflow contract: texts shrink; structural overflow errors.</para>
/// </summary>
public static class TitleBlockComponent
{
    private static readonly IReadOnlySet<string> ContentProps =
        new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "kicker", "title", "subtitle" };

    /// <summary>Expands <paramref name="element"/> into its primitive subtree.</summary>
    public static ContainerElement Expand(ComponentElement element, DesignTokens design, string path)
    {
        ArgumentNullException.ThrowIfNull(element);
        ArgumentNullException.ThrowIfNull(design);

        var reader = ComponentContentReader.For(path, "title_block", element.Content, ContentProps, design.Palette);
        var kicker = reader.String("kicker");
        var title = reader.String("title", required: true);
        var subtitle = reader.String("subtitle");
        reader.ThrowIfInvalid();

        var metrics = design.Metrics;
        var children = new List<GenElement>();
        if (kicker is not null)
        {
            children.Add(ComponentStyle.Line(kicker, "body", metrics.BodySizePt - 2, "accent", bold: true));
        }
        children.Add(ComponentStyle.Line(title!, "display", metrics.TitleSizePt, "ink"));
        if (subtitle is not null)
        {
            children.Add(ComponentStyle.Line(subtitle, "body", metrics.BodySizePt, "muted"));
        }

        ComponentStyle.RequireTokens(design, path, "title_block", "ink", kicker is not null ? "accent" : null, subtitle is not null ? "muted" : null);

        return new ContainerElement
        {
            Size = element.Size,
            At = element.At,
            Layout = new LayoutSpec { Mode = LayoutMode.Column, Gap = 4 },
            Children = children
        };
    }
}
