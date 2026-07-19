using PptxEditor.Core.Generation.Model;

namespace PptxEditor.Core.Generation.Components;

/// <summary>
/// The component layer entry point (plan.md §4): replaces every
/// <see cref="ComponentElement"/> in a generation document with its primitive subtree
/// (containers, texts, shapes — never direct OOXML/Typst calls). Expansion happens
/// exactly once, between parsing and layout; the layout resolver rejects unexpanded
/// components, so a pipeline that forgets this step fails loudly.
/// <para>
/// Each component expands to a subtree rooted at an element that inherits the
/// component's <c>size</c>/<c>at</c>, so components compose with row/column/grid layout
/// exactly like primitives (plan.md §2 rule 3). Component content is validated against
/// the strongly typed payloads in <c>ComponentContents.cs</c> with loud, actionable
/// errors (<see cref="ComponentException"/>).
/// </para>
/// </summary>
public static class ComponentExpander
{
    /// <summary>Expands every component in <paramref name="document"/>; returns the rewritten document.</summary>
    public static GenerationDocument Expand(GenerationDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);
        var slides = new List<ContainerElement>(document.Slides.Count);
        for (var i = 0; i < document.Slides.Count; i++)
        {
            slides.Add((ContainerElement)ExpandElement(document.Slides[i], document.Design, $"slides[{i}]"));
        }
        return document with { Slides = slides };
    }

    /// <summary>
    /// Expands one element subtree: components are replaced by their expansion,
    /// containers and groups are rewritten with expanded children, primitives pass
    /// through untouched.
    /// </summary>
    public static GenElement ExpandElement(GenElement element, DesignTokens design, string path)
    {
        ArgumentNullException.ThrowIfNull(element);
        ArgumentNullException.ThrowIfNull(design);
        return element switch
        {
            ComponentElement c => ExpandComponent(c, design, path),
            ContainerElement c => c with { Children = ExpandChildren(c.Children, design, path) },
            GroupElement g => g with { Children = ExpandChildren(g.Children, design, path) },
            _ => element
        };
    }

    private static IReadOnlyList<GenElement> ExpandChildren(IReadOnlyList<GenElement> children, DesignTokens design, string path)
    {
        var expanded = new List<GenElement>(children.Count);
        for (var i = 0; i < children.Count; i++)
        {
            expanded.Add(ExpandElement(children[i], design, $"{path}.children[{i}]"));
        }
        return expanded;
    }

    private static GenElement ExpandComponent(ComponentElement component, DesignTokens design, string path)
        => component.Name switch
        {
            "card" => CardComponent.Expand(component, design, path),
            "kpi" => KpiComponent.Expand(component, design, path),
            "title_block" => TitleBlockComponent.Expand(component, design, path),
            "bullet_list" => BulletListComponent.Expand(component, design, path),
            "divider" => DividerComponent.Expand(component, design, path),
            "badge" => BadgeComponent.Expand(component, design, path),
            "image_card" => ImageCardComponent.Expand(component, design, path),
            "table_block" => TableBlockComponent.Expand(component, design, path),
            _ => throw new ComponentException(path,
                $"unknown component '{component.Name}'. Known components: {string.Join(", ", ComponentElement.KnownNames.Order(StringComparer.Ordinal))}.")
        };
}
