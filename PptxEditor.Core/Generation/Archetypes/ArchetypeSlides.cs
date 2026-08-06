using System.Text.Json;
using PptxEditor.Core.Generation.Components;
using PptxEditor.Core.Generation.Model;

namespace PptxEditor.Core.Generation.Archetypes;

/// <summary>
/// Archetype slide functions (component pipeline, P10): the prompt-friendly authoring surface.
/// Each function composes a whole slide (<see cref="ContainerElement"/> root) exclusively
/// out of v1 components — container glue plus <see cref="ComponentElement"/> nodes; no
/// direct primitive construction, no emitter calls (enforced by test inspection). The
/// components expand in the regular <see cref="ComponentExpander"/> pass, so there is
/// exactly one component expansion code path (one layout,
/// one expansion).
/// <para>
/// The composition vocabulary is: a paper-filled slide root padded by
/// <c>metrics.marginPt</c>, a column flow gapped by <c>metrics.gutterPt</c>, and the
/// components themselves. Sizing uses grow weights only — coordinates stay the escape
/// hatch, never the API.
/// </para>
/// <para>
/// When an archetype expands an id-bearing slide (the parser always provides one, either
/// user-authored or synthesized "<c>{name}-{slideIndex}</c>"), every composed node gets a
/// deterministic role-based id derived from it: the slide root keeps the id, the title
/// block becomes "<c>{id}-title</c>", the accent rule "<c>{id}-accent</c>", the kpi row
/// "<c>{id}-row</c>", each kpi "<c>{id}-kpi-{n}</c>", and each two_col slot
/// "<c>{id}-left</c>"/"<c>{id}-right</c>" unless the slot declares its own explicit id.
/// Collisions with authored ids are disambiguated deterministically by the allocator.
/// </para>
/// </summary>
public static class ArchetypeSlides
{
    /// <summary>The five v1 archetype slide type names (JSON slide types).</summary>
    public static readonly IReadOnlySet<string> Names = new HashSet<string>(StringComparer.Ordinal)
    {
        "cover", "section", "kpi_row", "two_col", "table_slide"
    };

    /// <summary>True when <paramref name="name"/> is one of the v1 archetype slide types.</summary>
    public static bool IsArchetype(string name) => Names.Contains(name);

    private static readonly JsonSerializerOptions ContentJsonOptions = new()
    {
        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
    };

    /// <summary>
    /// <c>cover</c>: the opening slide — vertically centered title block (kicker, display
    /// title, muted subtitle) with an accent rule beneath, on generous editorial margins.
    /// </summary>
    public static ContainerElement Cover(CoverContent content, DesignTokens design)
        => Cover(content, design, prefix: null, allocator: null);

    internal static ContainerElement Cover(CoverContent content, DesignTokens design, string? prefix, ElementIdAllocator? allocator)
    {
        ArgumentNullException.ThrowIfNull(content);
        ArgumentNullException.ThrowIfNull(design);
        RequireText(content.Title, "cover", "title");

        var metrics = design.Metrics;
        return Slide(design, EdgeInsets.Symmetric(0, 2 * metrics.MarginPt), Justify.Center, metrics.GutterPt,
        [
            TitleBlock(content.Kicker, content.Title, content.Subtitle, size: HeadingSize(content.Kicker, content.Subtitle, metrics), prefix, allocator),
            AccentDivider(width: 2, prefix, allocator)
        ], prefix, allocator);
    }

    /// <summary>
    /// <c>section</c>: a bottom-anchored section divider behind an accent rule. With an
    /// index the slide composes a big-figure kpi cell beside the title block (1:2);
    /// without one, the title block stands alone.
    /// </summary>
    public static ContainerElement Section(SectionContent content, DesignTokens design)
        => Section(content, design, prefix: null, allocator: null);

    internal static ContainerElement Section(SectionContent content, DesignTokens design, string? prefix, ElementIdAllocator? allocator)
    {
        ArgumentNullException.ThrowIfNull(content);
        ArgumentNullException.ThrowIfNull(design);
        RequireText(content.Title, "section", "title");

        var metrics = design.Metrics;
        var children = new List<GenElement> { AccentDivider(width: 2, prefix, allocator) };
        if (content.Index is not null)
        {
            // The row is as tall as the kpi card's natural height: padding + value line + label line + inner gap.
            var figureHeight = Math.Round(
                2 * metrics.GutterPt + ComponentStyle.LineHeight(metrics.TitleSizePt) + ComponentStyle.LineHeight(metrics.BodySizePt) + 2, 3);
            children.Add(new ContainerElement
            {
                Id = allocator?.ChildId(prefix, "row"),
                Size = new SizeSpec { Height = figureHeight },
                Layout = new LayoutSpec { Mode = LayoutMode.Row, Gap = metrics.GutterPt, Align = AlignItems.Center },
                Children =
                [
                    Kpi(content.Index, content.Kicker ?? "SECTION", delta: null, size: new SizeSpec { Grow = 1 }, prefix, allocator, role: "kpi"),
                    TitleBlock(kicker: null, content.Title, content.Subtitle, size: new SizeSpec { Grow = 2 }, prefix, allocator)
                ]
            });
        }
        else
        {
            children.Add(TitleBlock(content.Kicker, content.Title, content.Subtitle, size: HeadingSize(content.Kicker, content.Subtitle, metrics), prefix, allocator));
        }
        return Slide(design, EdgeInsets.All(metrics.MarginPt), Justify.End, metrics.GutterPt, children, prefix, allocator);
    }

    /// <summary>
    /// <c>kpi_row</c>: an optional title block over a row of equally wide kpi cards
    /// (grow 1 each) filling the remaining slide height.
    /// </summary>
    public static ContainerElement KpiRow(KpiRowContent content, DesignTokens design)
        => KpiRow(content, design, prefix: null, allocator: null);

    internal static ContainerElement KpiRow(KpiRowContent content, DesignTokens design, string? prefix, ElementIdAllocator? allocator)
    {
        ArgumentNullException.ThrowIfNull(content);
        ArgumentNullException.ThrowIfNull(design);
        ArgumentNullException.ThrowIfNull(content.Kpis);
        if (content.Kpis.Count is < 1 or > 6)
        {
            throw new ComponentException("kpi_row", $"content.kpis: must contain between 1 and 6 KPIs (got {content.Kpis.Count}).");
        }
        for (var i = 0; i < content.Kpis.Count; i++)
        {
            RequireText(content.Kpis[i].Value, "kpi_row", $"kpis[{i}].value");
            RequireText(content.Kpis[i].Label, "kpi_row", $"kpis[{i}].label");
        }

        var metrics = design.Metrics;
        var children = new List<GenElement>();
        if (content.Title is not null)
        {
            children.Add(TitleBlock(kicker: null, content.Title, content.Subtitle, size: HeadingSize(null, content.Subtitle, metrics), prefix, allocator));
        }
        children.Add(new ContainerElement
        {
            Id = allocator?.ChildId(prefix, "row"),
            Size = new SizeSpec { Grow = 1 },
            Layout = new LayoutSpec { Mode = LayoutMode.Row, Gap = metrics.GutterPt },
            Children = content.Kpis
                .Select((kpi, index) => Kpi(kpi.Value, kpi.Label, kpi.Delta, size: new SizeSpec { Grow = 1 }, prefix, allocator, role: $"kpi-{index}"))
                .ToList<GenElement>()
        });
        return Slide(design, EdgeInsets.All(metrics.MarginPt), Justify.Start, metrics.GutterPt, children, prefix, allocator);
    }

    /// <summary>
    /// <c>two_col</c>: an optional title block over two component cells sharing the row
    /// by grow weights (default 1:1). Slots must be v1 components — archetypes never nest.
    /// </summary>
    public static ContainerElement TwoCol(TwoColContent content, DesignTokens design)
        => TwoCol(content, design, prefix: null, allocator: null);

    internal static ContainerElement TwoCol(TwoColContent content, DesignTokens design, string? prefix, ElementIdAllocator? allocator)
    {
        ArgumentNullException.ThrowIfNull(content);
        ArgumentNullException.ThrowIfNull(design);
        ArgumentNullException.ThrowIfNull(content.Left);
        ArgumentNullException.ThrowIfNull(content.Right);

        var weights = content.Weights;
        if (weights is not null)
        {
            if (weights.Count != 2)
            {
                throw new ComponentException("two_col", $"content.weights: must contain exactly 2 weights, left then right (got {weights.Count}).");
            }
            if (weights.Any(w => w <= 0))
            {
                throw new ComponentException("two_col", "content.weights: each weight must be > 0.");
            }
        }

        var metrics = design.Metrics;
        var children = new List<GenElement>();
        if (content.Title is not null)
        {
            children.Add(TitleBlock(kicker: null, content.Title, content.Subtitle, size: HeadingSize(null, content.Subtitle, metrics), prefix, allocator));
        }
        children.Add(new ContainerElement
        {
            Id = allocator?.ChildId(prefix, "row"),
            Size = new SizeSpec { Grow = 1 },
            Layout = new LayoutSpec { Mode = LayoutMode.Row, Gap = metrics.GutterPt },
            Children =
            [
                Slot(content.Left, weights?[0] ?? 1, "two_col", "left", prefix, allocator),
                Slot(content.Right, weights?[1] ?? 1, "two_col", "right", prefix, allocator)
            ]
        });
        return Slide(design, EdgeInsets.All(metrics.MarginPt), Justify.Start, metrics.GutterPt, children, prefix, allocator);
    }

    /// <summary>
    /// <c>table_slide</c>: an optional title block over a table_block that fills the
    /// remaining slide height.
    /// </summary>
    public static ContainerElement TableSlide(TableSlideContent content, DesignTokens design)
        => TableSlide(content, design, prefix: null, allocator: null);

    internal static ContainerElement TableSlide(TableSlideContent content, DesignTokens design, string? prefix, ElementIdAllocator? allocator)
    {
        ArgumentNullException.ThrowIfNull(content);
        ArgumentNullException.ThrowIfNull(design);
        ArgumentNullException.ThrowIfNull(content.Columns);
        ArgumentNullException.ThrowIfNull(content.Rows);
        if (content.Columns.Count < 1)
        {
            throw new ComponentException("table_slide", "content.columns: must contain at least one column header.");
        }

        var metrics = design.Metrics;
        var children = new List<GenElement>();
        if (content.Title is not null)
        {
            children.Add(TitleBlock(kicker: null, content.Title, content.Subtitle, size: HeadingSize(null, content.Subtitle, metrics), prefix, allocator));
        }
        children.Add(new ComponentElement
        {
            Id = allocator?.ChildId(prefix, "table"),
            Name = "table_block",
            Size = new SizeSpec { Grow = 1 },
            Content = ContentOf(new
            {
                columns = content.Columns,
                rows = content.Rows,
                header = content.Header,
                columnWeights = content.ColumnWeights,
                rowHeight = content.RowHeight
            })
        });
        return Slide(design, EdgeInsets.All(metrics.MarginPt), Justify.Start, metrics.GutterPt, children, prefix, allocator);
    }

    // ------------------------------------------------------------------ composition helpers

    /// <summary>The slide root every archetype shares: paper fill, margins, column flow.</summary>
    private static ContainerElement Slide(
        DesignTokens design, EdgeInsets padding, Justify justify, double gap, IReadOnlyList<GenElement> children,
        string? prefix, ElementIdAllocator? allocator)
        => new()
        {
            Id = prefix,
            Fill = new SolidFill("paper"),
            Padding = padding,
            Layout = new LayoutSpec { Mode = LayoutMode.Column, Gap = gap, Justify = justify },
            Children = children
        };

    private static ComponentElement TitleBlock(string? kicker, string title, string? subtitle, SizeSpec? size, string? prefix, ElementIdAllocator? allocator)
        => new()
        {
            Id = allocator?.ChildId(prefix, "title"),
            Name = "title_block",
            Size = size,
            Content = ContentOf(new { kicker, title, subtitle })
        };

    /// <summary>
    /// Deterministic heading height for a title_block placed on a column's main axis:
    /// the component's line estimates plus its 4 pt inter-line gaps (the layout engine
    /// rejects undetermined main-axis sizes — components sized by the archetype never
    /// hit that path).
    /// </summary>
    private static SizeSpec HeadingSize(string? kicker, string? subtitle, MetricTokens metrics)
    {
        var height = ComponentStyle.LineHeight(metrics.TitleSizePt);
        if (kicker is not null)
        {
            height += ComponentStyle.LineHeight(metrics.BodySizePt - 2) + 4;
        }
        if (subtitle is not null)
        {
            height += ComponentStyle.LineHeight(metrics.BodySizePt) + 4;
        }
        return new SizeSpec { Height = Math.Round(height, 3) };
    }

    private static ComponentElement Kpi(string value, string label, string? delta, SizeSpec? size, string? prefix, ElementIdAllocator? allocator, string role)
        => new()
        {
            Id = allocator?.ChildId(prefix, role),
            Name = "kpi",
            Size = size,
            Content = ContentOf(new { value, label, delta })
        };

    private static ComponentElement AccentDivider(double width, string? prefix, ElementIdAllocator? allocator)
        => new()
        {
            Id = allocator?.ChildId(prefix, "accent"),
            Name = "divider",
            Content = ContentOf(new { color = "accent", width })
        };

    /// <summary>Assigns the slot's grow weight; the archetype owns slot geometry, never the author.</summary>
    private static ComponentElement Slot(ComponentElement slot, double grow, string archetype, string slotName, string? prefix, ElementIdAllocator? allocator)
    {
        if (!ComponentElement.KnownNames.Contains(slot.Name))
        {
            throw new ComponentException(archetype,
                $"content.{slotName}.type: unknown component '{slot.Name}'. Known components: {string.Join(", ", ComponentElement.KnownNames.Order(StringComparer.Ordinal))}.");
        }
        if (slot.Size is not null || slot.At is not null)
        {
            throw new ComponentException(archetype,
                $"content.{slotName}: the archetype assigns slot geometry — 'size' and 'at' are not allowed on a two_col slot.");
        }
        // An explicit user id on the slot wins; otherwise the slot gets its role id.
        var id = slot.Id ?? allocator?.ChildId(prefix, slotName);
        return slot with { Id = id, Size = new SizeSpec { Grow = grow } };
    }

    private static JsonElement ContentOf(object payload) => JsonSerializer.SerializeToElement(payload, ContentJsonOptions);

    private static void RequireText(string? value, string archetype, string property)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            throw new ComponentException(archetype, $"content.{property}: is required and must not be empty.");
        }
    }
}
