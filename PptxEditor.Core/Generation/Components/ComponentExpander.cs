using PptxEditor.Core.Generation.Model;

namespace PptxEditor.Core.Generation.Components;

/// <summary>
/// The component layer entry point (component layer): replaces every
/// <see cref="ComponentElement"/> in a generation document with its primitive subtree
/// (containers, texts, shapes — never direct OOXML/Typst calls). Expansion happens
/// exactly once, between parsing and layout; the layout resolver rejects unexpanded
/// components, so a pipeline that forgets this step fails loudly.
/// <para>
/// Each component expands to a subtree rooted at an element that inherits the
/// component's <c>size</c>/<c>at</c>, so components compose with row/column/grid layout
/// exactly like primitives (absolute placement is an escape hatch only). Component content is validated against
/// the strongly typed payloads in <c>ComponentContents.cs</c> with loud, actionable
/// errors (<see cref="ComponentException"/>).
/// </para>
/// <para>
/// Ids: the expansion root keeps the component's id (an explicit user id always wins);
/// every generated node gets a deterministic role-based id derived from it —
/// "<c>{id}-title</c>", "<c>{id}-item-0</c>", … Collisions with authored ids are
/// disambiguated deterministically by <see cref="ElementIdAllocator"/>.
/// </para>
/// </summary>
public static class ComponentExpander
{
    /// <summary>Expands every component in <paramref name="document"/>; returns the rewritten document.</summary>
    public static GenerationDocument Expand(GenerationDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);
        var slides = new List<ContainerElement>(document.Slides.Count);
        // Seed with every id already in the document so synthesized child ids never collide
        // with authored ones (deterministic first-fit; see ElementIdAllocator).
        var allocator = ElementIdAllocator.SeedWith(document.Slides);
        for (var i = 0; i < document.Slides.Count; i++)
        {
            slides.Add((ContainerElement)ExpandElement(document.Slides[i], document.Design, $"slides[{i}]", allocator));
        }
        return document with { Slides = slides };
    }

    /// <summary>
    /// Expands one element subtree: components are replaced by their expansion,
    /// containers and groups are rewritten with expanded children, primitives pass
    /// through untouched.
    /// </summary>
    public static GenElement ExpandElement(GenElement element, DesignTokens design, string path)
        => ExpandElement(element, design, path, ElementIdAllocator.SeedWith(element));

    private static GenElement ExpandElement(GenElement element, DesignTokens design, string path, ElementIdAllocator allocator)
    {
        ArgumentNullException.ThrowIfNull(element);
        ArgumentNullException.ThrowIfNull(design);
        return element switch
        {
            ComponentElement c => ExpandComponent(c, design, path, allocator),
            ContainerElement c => c with { Children = ExpandChildren(c.Children, design, path, allocator) },
            GroupElement g => g with { Children = ExpandChildren(g.Children, design, path, allocator) },
            _ => element
        };
    }

    private static IReadOnlyList<GenElement> ExpandChildren(IReadOnlyList<GenElement> children, DesignTokens design, string path, ElementIdAllocator allocator)
    {
        var expanded = new List<GenElement>(children.Count);
        for (var i = 0; i < children.Count; i++)
        {
            expanded.Add(ExpandElement(children[i], design, $"{path}.children[{i}]", allocator));
        }
        return expanded;
    }

    private static GenElement ExpandComponent(ComponentElement component, DesignTokens design, string path, ElementIdAllocator allocator)
        => component.Name switch
        {
            "card" => CardComponent.Expand(component, design, path, allocator),
            "kpi" => KpiComponent.Expand(component, design, path, allocator),
            "title_block" => TitleBlockComponent.Expand(component, design, path, allocator),
            "bullet_list" => BulletListComponent.Expand(component, design, path, allocator),
            "divider" => DividerComponent.Expand(component, design, path, allocator),
            "badge" => BadgeComponent.Expand(component, design, path, allocator),
            "image_card" => ImageCardComponent.Expand(component, design, path, allocator),
            "table_block" => TableBlockComponent.Expand(component, design, path, allocator),
            _ => throw new ComponentException(path,
                $"unknown component '{component.Name}'. Known components: {string.Join(", ", ComponentElement.KnownNames.Order(StringComparer.Ordinal))}.")
        };
}
