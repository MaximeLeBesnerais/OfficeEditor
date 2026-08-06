using PptxEditor.Core.Generation.Model;

namespace PptxEditor.Core.Generation.Components;

/// <summary>
/// <c>badge</c> (component layer): a small pill label — an accent-filled container with a
/// radius of exactly half its height and a centered bold label. Height is computed from
/// the body-size line estimate plus padding unless the author fixes it; width (or grow)
/// is the author's to give on the parent's layout axis (no hug-content in v1).
/// <para>Overflow contract: the label shrinks; structural overflow errors.</para>
/// </summary>
public static class BadgeComponent
{
    private const double PadV = 3;
    private const double PadH = 10;

    private static readonly IReadOnlySet<string> ContentProps =
        new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "text", "color", "textColor" };

    /// <summary>Expands <paramref name="element"/> into its primitive subtree.</summary>
    internal static ContainerElement Expand(ComponentElement element, DesignTokens design, string path, ElementIdAllocator allocator)
    {
        ArgumentNullException.ThrowIfNull(element);
        ArgumentNullException.ThrowIfNull(design);

        var reader = ComponentContentReader.For(path, "badge", element.Content, ContentProps, design.Palette);
        var text = reader.String("text", required: true);
        var color = reader.Color("color") ?? "accent";
        var textColor = reader.Color("textColor") ?? "paper";
        reader.ThrowIfInvalid();

        var fontSize = design.Metrics.BodySizePt - 2;
        var height = element.Size?.Height ?? Math.Round(ComponentStyle.LineHeight(fontSize) + 2 * PadV, 3);
        var size = (element.Size ?? new SizeSpec()) with { Height = height };

        ComponentStyle.RequireTokens(design, path, "badge", color, textColor);

        return new ContainerElement
        {
            Id = element.Id,
            Size = size,
            At = element.At,
            Fill = new SolidFill(color),
            Radius = CornerRadii.All(height / 2),
            Padding = EdgeInsets.Symmetric(PadV, PadH),
            Layout = new LayoutSpec { Mode = LayoutMode.Row, Justify = Justify.Center, Align = AlignItems.Center },
            Children =
            [
                ComponentStyle.Growing(text!, "body", fontSize, textColor, bold: true, align: TextAlign.Center, anchor: TextAnchor.Middle) with { Id = allocator.ChildId(element.Id, "label") }
            ]
        };
    }
}
