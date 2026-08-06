using PptxEditor.Core.Generation.Model;

namespace PptxEditor.Core.Generation.Components;

/// <summary>
/// <c>bullet_list</c> (component layer): an optional bold heading over single-line items,
/// each a row with a small accent ellipse marker and a growing text. Item rows take a
/// deterministic line height; the list's own height comes from the author (grow or fixed).
/// <para>Overflow contract: item texts shrink (long items get smaller, never silently
/// clip); more items than the box holds is a structural error (default
/// container policy).</para>
/// </summary>
public static class BulletListComponent
{
    private static readonly IReadOnlySet<string> ContentProps =
        new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "title", "items", "markerColor" };

    /// <summary>Expands <paramref name="element"/> into its primitive subtree.</summary>
    internal static ContainerElement Expand(ComponentElement element, DesignTokens design, string path, ElementIdAllocator allocator)
    {
        ArgumentNullException.ThrowIfNull(element);
        ArgumentNullException.ThrowIfNull(design);

        var reader = ComponentContentReader.For(path, "bullet_list", element.Content, ContentProps, design.Palette);
        var title = reader.String("title");
        var items = reader.StringArray("items", required: true, minCount: 1);
        var markerColor = reader.Color("markerColor") ?? "accent";
        reader.ThrowIfInvalid();

        var metrics = design.Metrics;
        var children = new List<GenElement>();
        if (title is not null)
        {
            children.Add(ComponentStyle.Line(title, "display", metrics.BodySizePt + 2, "ink", bold: true) with { Id = allocator.ChildId(element.Id, "title") });
        }
        for (var i = 0; i < items!.Count; i++)
        {
            var rowId = allocator.ChildId(element.Id, $"item-{i}");
            children.Add(new ContainerElement
            {
                Id = rowId,
                Size = new SizeSpec { Height = ComponentStyle.LineHeight(metrics.BodySizePt) },
                Layout = new LayoutSpec { Mode = LayoutMode.Row, Gap = 8, Align = AlignItems.Center },
                Children =
                [
                    new EllipseElement
                    {
                        Id = allocator.ChildId(rowId, "marker"),
                        Size = new SizeSpec { Width = 6, Height = 6 },
                        Fill = new SolidFill(markerColor)
                    },
                    ComponentStyle.Growing(items[i], "body", metrics.BodySizePt, "ink", anchor: TextAnchor.Middle) with { Id = allocator.ChildId(rowId, "text") }
                ]
            });
        }

        ComponentStyle.RequireTokens(design, path, "bullet_list", "ink", markerColor);

        return new ContainerElement
        {
            Id = element.Id,
            Size = element.Size,
            At = element.At,
            Layout = new LayoutSpec { Mode = LayoutMode.Column, Gap = Math.Round(metrics.GutterPt / 2, 3) },
            Children = children
        };
    }
}
