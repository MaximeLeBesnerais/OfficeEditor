using PptxEditor.Core.Generation.Model;

namespace PptxEditor.Core.Generation;

/// <summary>
/// Deterministic id reservation for expansion-synthesized element ids. Seeded with every id
/// already present in the document (user-authored or parser-synthesized), so a synthesized
/// "<c>{prefix}-{role}</c>" that collides with an authored id is disambiguated by a numeric
/// suffix ("-2", "-3", …) instead of silently creating a duplicate. Pure first-fit allocation
/// in document order: the same input JSON always yields the same ids across re-parses.
/// </summary>
internal sealed class ElementIdAllocator
{
    private readonly HashSet<string> _taken = new(StringComparer.Ordinal);

    /// <summary>Creates an allocator seeded with every id found in the given slides.</summary>
    public static ElementIdAllocator SeedWith(IEnumerable<ContainerElement> slides)
    {
        var allocator = new ElementIdAllocator();
        foreach (var slide in slides)
        {
            SeedElement(allocator, slide);
        }
        return allocator;
    }

    /// <summary>Creates an allocator seeded with every id found under the given element.</summary>
    public static ElementIdAllocator SeedWith(GenElement element)
    {
        var allocator = new ElementIdAllocator();
        SeedElement(allocator, element);
        return allocator;
    }

    /// <summary>True when <paramref name="id"/> is already claimed in the document.</summary>
    public bool Contains(string id) => _taken.Contains(id);

    /// <summary>Claims <paramref name="id"/>; false when it was already taken.</summary>
    public bool TryReserve(string id) => _taken.Add(id);

    /// <summary>
    /// Reserves and returns "<c>{prefix}-{role}</c>", suffixing "-2", "-3", … on collision.
    /// <paramref name="prefix"/> must be non-null and unique; roles are unique per expansion,
    /// so this stays collision-free against both authored ids and earlier synthesized ids.
    /// </summary>
    public string Take(string prefix, string role)
    {
        var candidate = $"{prefix}-{role}";
        if (_taken.Add(candidate))
        {
            return candidate;
        }
        for (var n = 2; ; n++)
        {
            var suffixed = $"{candidate}-{n}";
            if (_taken.Add(suffixed))
            {
                return suffixed;
            }
        }
    }

    /// <summary>
    /// Like <see cref="Take(string, string)"/> but returns null when <paramref name="prefix"/>
    /// is null — i.e. the parent element has no id (programmatic construction outside the
    /// parser), so there is nothing meaningful to derive a child id from.
    /// </summary>
    public string? ChildId(string? prefix, string role)
        => prefix is null ? null : Take(prefix, role);

    private static void SeedElement(ElementIdAllocator allocator, GenElement element)
    {
        if (element.Id is { } id)
        {
            allocator._taken.Add(id);
        }
        switch (element)
        {
            case ContainerElement container:
                foreach (var child in container.Children)
                {
                    SeedElement(allocator, child);
                }
                break;
            case GroupElement group:
                foreach (var child in group.Children)
                {
                    SeedElement(allocator, child);
                }
                break;
        }
    }
}
