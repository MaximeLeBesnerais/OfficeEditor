using PptxEditor.Core.Generation.Model;

namespace PptxEditor.Core.Generation.Components;

/// <summary>
/// <c>divider</c> (component layer): a straight rule. Expands to a single line primitive.
/// Defaults: horizontal, "muted", 1 pt; when the author does not fix the cross-axis
/// dimension it defaults to the stroke width (a horizontal divider is 1 pt tall).
/// <para>Overflow contract: no text — geometry overflow is governed by the parent
/// container's policy.</para>
/// </summary>
public static class DividerComponent
{
    private static readonly IReadOnlySet<string> ContentProps =
        new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "color", "width", "orientation" };

    private static readonly IReadOnlyDictionary<string, LineOrientation> Orientations =
        new Dictionary<string, LineOrientation>(StringComparer.OrdinalIgnoreCase)
        {
            ["horizontal"] = LineOrientation.Horizontal, ["vertical"] = LineOrientation.Vertical
        };

    /// <summary>Expands <paramref name="element"/> into its primitive subtree.</summary>
    public static LineElement Expand(ComponentElement element, DesignTokens design, string path)
    {
        ArgumentNullException.ThrowIfNull(element);
        ArgumentNullException.ThrowIfNull(design);

        var reader = ComponentContentReader.For(path, "divider", element.Content, ContentProps, design.Palette, required: false);
        var color = reader.Color("color") ?? "muted";
        var width = reader.Number("width", min: 0.01) ?? 1;
        var orientation = reader.Enum("orientation", Orientations, "line orientation") ?? LineOrientation.Horizontal;
        reader.ThrowIfInvalid();

        ComponentStyle.RequireTokens(design, path, "divider", color);

        var size = element.Size is null
            ? new SizeSpec()
            : element.Size with { };
        size = orientation == LineOrientation.Horizontal
            ? size with { Height = size.Height ?? width }
            : size with { Width = size.Width ?? width };

        return new LineElement
        {
            Size = size,
            At = element.At,
            Orientation = orientation,
            Stroke = new StrokeSpec { Color = color, WidthPt = width }
        };
    }
}
