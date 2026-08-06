using System.Text.Json;
using PptxEditor.Core.Generation.Components;
using PptxEditor.Core.Generation.Model;

namespace PptxEditor.Core.Generation.Archetypes;

/// <summary>
/// The archetype layer entry point (component pipeline, P10): replaces every archetype marker
/// slide in a generation document with the archetype's composed slide. A marker is the
/// shape the schema layer produces for an archetype slide type (<c>{"type": "kpi_row",
/// "content": {…}}</c> at slide root): a bare root container whose only
/// child is a <see cref="ComponentElement"/> named after the archetype.
/// <para>
/// Runs exactly once, after parsing and BEFORE <see cref="ComponentExpander"/>: the
/// archetype composes <see cref="ComponentElement"/> nodes, which the component pass
/// then expands — one expansion order, no second layout system. Raw content is validated
/// with the component layer's reader, so malformed archetype payloads fail loudly with
/// JSON paths and suggestions, exactly like component payloads.
/// </para>
/// </summary>
public static class ArchetypeExpander
{
    private static readonly IReadOnlySet<string> CoverProps = Set("kicker", "title", "subtitle");
    private static readonly IReadOnlySet<string> SectionProps = Set("index", "kicker", "title", "subtitle");
    private static readonly IReadOnlySet<string> KpiRowProps = Set("title", "subtitle", "kpis");
    private static readonly IReadOnlySet<string> KpiProps = Set("value", "label", "delta");
    private static readonly IReadOnlySet<string> TwoColProps = Set("title", "subtitle", "left", "right", "weights");
    private static readonly IReadOnlySet<string> SlotProps = Set("type", "content");
    private static readonly IReadOnlySet<string> TableSlideProps = Set("title", "subtitle", "columns", "rows", "header", "columnWeights", "rowHeight");

    /// <summary>Expands every archetype marker slide; other slides pass through untouched.</summary>
    public static GenerationDocument Expand(GenerationDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);
        var changed = false;
        var slides = new List<ContainerElement>(document.Slides.Count);
        for (var i = 0; i < document.Slides.Count; i++)
        {
            var path = $"slides[{i}]";
            var slideRoot = document.Slides[i];
            if (IsArchetypeSlide(slideRoot, out var marker))
            {
                if (marker.Size is not null || marker.At is not null)
                {
                    throw new ComponentException(path,
                        $"archetype slide '{marker.Name}' fills the whole slide; 'size' and 'at' are not allowed on it.");
                }
                // Notes are per-slide metadata on the root container: carry them onto the
                // composed slide (ArchetypeSlides builds a fresh container without them).
                slides.Add(ExpandSlide(marker, document.Design, path) with { Notes = slideRoot.Notes });
                changed = true;
            }
            else
            {
                slides.Add(slideRoot);
            }
        }
        return changed ? document with { Slides = slides } : document;
    }

    /// <summary>
    /// True when <paramref name="slide"/> is an archetype marker: a bare root container
    /// (no layout, padding, surface or overflow override) holding exactly one
    /// <see cref="ComponentElement"/> whose name is an archetype slide type. Only the
    /// schema layer produces this shape — archetype names are rejected at child level.
    /// </summary>
    public static bool IsArchetypeSlide(ContainerElement slide, out ComponentElement marker)
    {
        ArgumentNullException.ThrowIfNull(slide);
        if (slide.Layout is null
            && slide.Padding is null
            && slide.Fill is null
            && slide.Stroke is null
            && slide.Radius is null
            && slide.Shadow is null
            && slide.Overflow == OverflowPolicy.Error
            && slide.Children.Count == 1
            && slide.Children[0] is ComponentElement component
            && ArchetypeSlides.IsArchetype(component.Name))
        {
            marker = component;
            return true;
        }
        marker = null!;
        return false;
    }

    private static ContainerElement ExpandSlide(ComponentElement marker, DesignTokens design, string path)
        => marker.Name switch
        {
            "cover" => ArchetypeSlides.Cover(ReadCover(marker.Content, design, path), design),
            "section" => ArchetypeSlides.Section(ReadSection(marker.Content, design, path), design),
            "kpi_row" => ArchetypeSlides.KpiRow(ReadKpiRow(marker.Content, design, path), design),
            "two_col" => ArchetypeSlides.TwoCol(ReadTwoCol(marker.Content, design, path), design),
            "table_slide" => ArchetypeSlides.TableSlide(ReadTableSlide(marker.Content, design, path), design),
            _ => throw new ComponentException(path,
                $"unknown archetype slide '{marker.Name}'. Known archetypes: {string.Join(", ", ArchetypeSlides.Names.Order(StringComparer.Ordinal))}.")
        };

    // ------------------------------------------------------------------ content readers (JSON route → typed payloads)

    private static CoverContent ReadCover(JsonElement? content, DesignTokens design, string path)
    {
        var reader = ComponentContentReader.For(path, "cover", content, CoverProps, design.Palette);
        var title = reader.String("title", required: true);
        var subtitle = reader.String("subtitle");
        var kicker = reader.String("kicker");
        reader.ThrowIfInvalid();
        return new CoverContent { Title = title!, Subtitle = subtitle, Kicker = kicker };
    }

    private static SectionContent ReadSection(JsonElement? content, DesignTokens design, string path)
    {
        var reader = ComponentContentReader.For(path, "section", content, SectionProps, design.Palette);
        var title = reader.String("title", required: true);
        var subtitle = reader.String("subtitle");
        var kicker = reader.String("kicker");
        var index = reader.String("index");
        reader.ThrowIfInvalid();
        return new SectionContent { Title = title!, Subtitle = subtitle, Kicker = kicker, Index = index };
    }

    private static KpiRowContent ReadKpiRow(JsonElement? content, DesignTokens design, string path)
    {
        var reader = ComponentContentReader.For(path, "kpi_row", content, KpiRowProps, design.Palette);
        var title = reader.String("title");
        var subtitle = reader.String("subtitle");
        reader.ThrowIfInvalid();

        if (!TryGet(content, "kpis", out var kpisEl))
        {
            throw new ComponentException(path, "content.kpis: is required (an array of {\"value\":…,\"label\":…,\"delta\":…}).");
        }
        if (kpisEl.ValueKind != JsonValueKind.Array)
        {
            throw new ComponentException(path, "content.kpis: must be an array of {\"value\":…,\"label\":…,\"delta\":…}.");
        }
        var kpis = new List<KpiItem>();
        var index = 0;
        foreach (var item in kpisEl.EnumerateArray())
        {
            var itemPath = $"{path}.content.kpis[{index}]";
            var itemReader = ComponentContentReader.For(itemPath, "kpi", item, KpiProps, design.Palette);
            var value = itemReader.String("value", required: true);
            var label = itemReader.String("label", required: true);
            var delta = itemReader.String("delta");
            itemReader.ThrowIfInvalid();
            kpis.Add(new KpiItem { Value = value!, Label = label!, Delta = delta });
            index++;
        }
        if (kpis.Count is < 1 or > 6)
        {
            throw new ComponentException(path, $"content.kpis: must contain between 1 and 6 KPIs (got {kpis.Count}).");
        }
        return new KpiRowContent { Title = title, Subtitle = subtitle, Kpis = kpis };
    }

    private static TwoColContent ReadTwoCol(JsonElement? content, DesignTokens design, string path)
    {
        var reader = ComponentContentReader.For(path, "two_col", content, TwoColProps, design.Palette);
        var title = reader.String("title");
        var subtitle = reader.String("subtitle");
        var weights = reader.NumberArray("weights");
        reader.ThrowIfInvalid();
        if (weights is not null && weights.Count != 2)
        {
            throw new ComponentException(path, $"content.weights: must contain exactly 2 weights, left then right (got {weights.Count}).");
        }
        return new TwoColContent
        {
            Title = title,
            Subtitle = subtitle,
            Left = ReadSlot(content, "left", design, path),
            Right = ReadSlot(content, "right", design, path),
            Weights = weights
        };
    }

    private static ComponentElement ReadSlot(JsonElement? content, string slotName, DesignTokens design, string path)
    {
        var slotPath = $"{path}.content.{slotName}";
        if (!TryGet(content, slotName, out var slotEl))
        {
            throw new ComponentException(slotPath, $"'{slotName}' is required (a component slot: {{\"type\": \"bullet_list\", \"content\": {{…}}}}).");
        }
        var reader = ComponentContentReader.For(slotPath, $"two_col.{slotName}", slotEl, SlotProps, design.Palette);
        var type = reader.String("type", required: true);
        reader.ThrowIfInvalid();

        if (ArchetypeSlides.IsArchetype(type!))
        {
            throw new ComponentException(slotPath,
                $"content.{slotName}.type: '{type}' is an archetype slide type — archetypes never nest; use a component ({string.Join(", ", ComponentElement.KnownNames.Order(StringComparer.Ordinal))}).");
        }
        if (!ComponentElement.KnownNames.Contains(type!))
        {
            throw new ComponentException(slotPath,
                $"content.{slotName}.type: unknown component '{type}'. Known components: {string.Join(", ", ComponentElement.KnownNames.Order(StringComparer.Ordinal))}.");
        }

        JsonElement? slotContent = null;
        if (TryGet(slotEl, "content", out var contentEl))
        {
            if (contentEl.ValueKind != JsonValueKind.Object)
            {
                throw new ComponentException(slotPath, $"content.{slotName}.content: must be an object (component payload).");
            }
            slotContent = contentEl.Clone();
        }
        return new ComponentElement { Name = type!, Content = slotContent };
    }

    private static TableSlideContent ReadTableSlide(JsonElement? content, DesignTokens design, string path)
    {
        var reader = ComponentContentReader.For(path, "table_slide", content, TableSlideProps, design.Palette);
        var title = reader.String("title");
        var subtitle = reader.String("subtitle");
        var columns = reader.StringArray("columns", required: true, minCount: 1);
        var rows = reader.StringMatrix("rows", required: true);
        var header = reader.Bool("header", defaultValue: true);
        var weights = reader.NumberArray("columnWeights");
        var rowHeight = reader.Number("rowHeight", min: 1);
        reader.ThrowIfInvalid();
        return new TableSlideContent
        {
            Title = title,
            Subtitle = subtitle,
            Columns = columns!,
            Rows = rows!,
            Header = header,
            ColumnWeights = weights,
            RowHeight = rowHeight
        };
    }

    private static IReadOnlySet<string> Set(params string[] names) => new HashSet<string>(names, StringComparer.OrdinalIgnoreCase);

    private static bool TryGet(JsonElement? obj, string name, out JsonElement value)
    {
        if (obj is { ValueKind: JsonValueKind.Object } content)
        {
            foreach (var property in content.EnumerateObject())
            {
                if (string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    value = property.Value;
                    return true;
                }
            }
        }
        value = default;
        return false;
    }
}
