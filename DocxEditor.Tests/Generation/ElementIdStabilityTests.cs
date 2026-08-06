using PptxEditor.Core.Generation.Archetypes;
using PptxEditor.Core.Generation.Components;
using PptxEditor.Core.Generation.Model;
using PptxEditor.Core.Generation.Schema;

namespace DocxEditor.Tests.Generation;

/// <summary>
/// SPEC §9.2 (stable element ids): expansion is deterministic — the same deck parsed and
/// expanded twice yields identical ids, explicit user ids always win over synthesized ones,
/// and expansion-synthesized ids never collide with authored ids (suffix "-2", "-3", …).
/// </summary>
public sealed class ElementIdStabilityTests
{
    private readonly GenerationDocumentParser _parser = new();

    private const string DeckJson = """
        {
          "version": "2.0",
          "design": {
            "palette": { "primary": "#0B3D91", "accent": "#FF6B00", "ink": "#1A1A1A",
                         "paper": "#FFFFFF", "muted": "#8A94A6" }
          },
          "slides": [
            {
              "type": "cover",
              "id": "opening",
              "content": { "title": "Northwind Labs", "kicker": "Q3 2026", "subtitle": "Quarterly results" }
            },
            {
              "type": "container",
              "layout": { "mode": "column" },
              "children": [
                { "type": "card", "id": "revenue", "content": { "title": "+34%", "subtitle": "Revenue" } }
              ]
            }
          ]
        }
        """;

    private static IReadOnlyList<string> CollectIds(GenerationDocument document)
    {
        var ids = new List<string>();
        Walk(document.Slides);
        return ids;

        void Walk(IReadOnlyList<ContainerElement> slides)
        {
            foreach (var slide in slides)
            {
                WalkElement(slide);
            }
        }

        void WalkElement(GenElement element)
        {
            if (element.Id is { } id)
            {
                ids.Add(id);
            }
            switch (element)
            {
                case ContainerElement container:
                    foreach (var child in container.Children)
                    {
                        WalkElement(child);
                    }
                    break;
                case GroupElement group:
                    foreach (var child in group.Children)
                    {
                        WalkElement(child);
                    }
                    break;
            }
        }
    }

    private static IReadOnlyList<string> RunFullPipeline(string json)
    {
        var document = new GenerationDocumentParser().Parse(json);
        var archetyped = ArchetypeExpander.Expand(document);
        var componentized = ComponentExpander.Expand(archetyped);
        return CollectIds(componentized);
    }

    [Fact]
    public void FullPipeline_ReParsedDeck_YieldsIdenticalIds()
    {
        var first = RunFullPipeline(DeckJson);
        var second = RunFullPipeline(DeckJson);

        Assert.Equal(first, second);
    }

    [Fact]
    public void ArchetypeExpansion_DerivesRoleBasedIds_AndUserIdWins()
    {
        var document = _parser.Parse(DeckJson);
        var archetyped = ArchetypeExpander.Expand(document);

        // The composed slide root keeps the marker's id.
        var cover = archetyped.Slides[0];
        Assert.Equal("opening", cover.Id);

        // Children derive role-based ids from it: title block + accent divider (still
        // unexpanded components at this stage — the component pass expands them next).
        var titleBlock = Assert.IsType<ComponentElement>(cover.Children[0]);
        Assert.Equal("opening-title", titleBlock.Id);
        Assert.Equal("title_block", titleBlock.Name);
        var accent = Assert.IsType<ComponentElement>(cover.Children[1]);
        Assert.Equal("opening-accent", accent.Id);
        Assert.Equal("divider", accent.Name);

        // The un-ided container slide got a deterministic fallback id.
        Assert.Equal("slide-0", archetyped.Slides[1].Id);
    }

    [Fact]
    public void ComponentExpansion_DerivesRoleBasedIds_AndExplicitIdWins()
    {
        var document = _parser.Parse(DeckJson);
        var archetyped = ArchetypeExpander.Expand(document);
        var componentized = ComponentExpander.Expand(archetyped);

        // The card keeps its explicit id and its fields derive from it.
        var card = Assert.IsType<ContainerElement>(componentized.Slides[1].Children[0]);
        Assert.Equal("revenue", card.Id);
        Assert.Equal("revenue-title", Assert.IsType<TextElement>(card.Children[0]).Id);
        Assert.Equal("revenue-subtitle", Assert.IsType<TextElement>(card.Children[1]).Id);

        // The cover's title block children derive from the block id.
        var titleBlock = Assert.IsType<ContainerElement>(componentized.Slides[0].Children[0]);
        Assert.Equal("opening-title", titleBlock.Id);
        Assert.Equal("opening-title-kicker", Assert.IsType<TextElement>(titleBlock.Children[0]).Id);
        Assert.Equal("opening-title-title", Assert.IsType<TextElement>(titleBlock.Children[1]).Id);
        Assert.Equal("opening-title-subtitle", Assert.IsType<TextElement>(titleBlock.Children[2]).Id);
    }

    [Fact]
    public void SynthesizedSlideId_IsDeterministicPerArchetypeAndIndex()
    {
        var document = _parser.Parse("""
            {
              "version": "2.0",
              "design": { "palette": { "accent": "#FF6B00", "ink": "#1A1A1A", "paper": "#FFFFFF" } },
              "slides": [
                { "type": "cover", "content": { "title": "One" } },
                { "type": "cover", "content": { "title": "Two" } }
              ]
            }
            """);

        var archetyped = ArchetypeExpander.Expand(document);

        Assert.Equal("cover-0", archetyped.Slides[0].Id);
        Assert.Equal("cover-1", archetyped.Slides[1].Id);
        Assert.Equal("cover-0-title", Assert.IsType<ComponentElement>(archetyped.Slides[0].Children[0]).Id);
        Assert.Equal("cover-1-title", Assert.IsType<ComponentElement>(archetyped.Slides[1].Children[0]).Id);
    }

    [Fact]
    public void TwoColSlots_ExplicitSlotIdWins_SynthesizedSlotGetsRoleId()
    {
        var document = _parser.Parse("""
            {
              "version": "2.0",
              "design": { "palette": { "accent": "#FF6B00", "ink": "#1A1A1A", "paper": "#FFFFFF" } },
              "slides": [
                {
                  "type": "two_col",
                  "id": "compare",
                  "content": {
                    "left":  { "type": "bullet_list", "id": "left-panel", "content": { "items": ["A"] } },
                    "right": { "type": "card", "content": { "title": "B" } }
                  }
                }
              ]
            }
            """);

        var archetyped = ArchetypeExpander.Expand(document);

        // The row holds the two slots; the explicit slot id wins, the other gets its role id.
        var row = Assert.IsType<ContainerElement>(archetyped.Slides[0].Children[0]);
        var left = Assert.IsType<ComponentElement>(row.Children[0]);
        var right = Assert.IsType<ComponentElement>(row.Children[1]);
        Assert.Equal("left-panel", left.Id);
        Assert.Equal("compare-right", right.Id);

        // The component pass derives children ids from the slot ids.
        var componentized = ComponentExpander.Expand(archetyped);
        var expandedRow = Assert.IsType<ContainerElement>(componentized.Slides[0].Children[0]);
        var leftExpanded = Assert.IsType<ContainerElement>(expandedRow.Children[0]);
        Assert.Equal("left-panel", leftExpanded.Id);
        Assert.Equal("left-panel-item-0", Assert.IsType<ContainerElement>(leftExpanded.Children[0]).Id);
        var rightExpanded = Assert.IsType<ContainerElement>(expandedRow.Children[1]);
        Assert.Equal("compare-right-title", Assert.IsType<TextElement>(rightExpanded.Children[0]).Id);
    }

    [Fact]
    public void SynthesizedId_CollidingWithAuthoredId_IsSuffixedDeterministically()
    {
        // The user authored "opening-title" on a text elsewhere, so the cover's title block
        // cannot take that name; the allocator disambiguates to "opening-title-2".
        var document = _parser.Parse("""
            {
              "version": "2.0",
              "design": { "palette": { "accent": "#FF6B00", "ink": "#1A1A1A", "paper": "#FFFFFF" } },
              "slides": [
                { "type": "cover", "id": "opening", "content": { "title": "Northwind" } },
                {
                  "type": "container",
                  "layout": { "mode": "column" },
                  "children": [
                    { "type": "text", "id": "opening-title", "text": "authored", "size": { "h": 20 } }
                  ]
                }
              ]
            }
            """);

        var archetyped = ArchetypeExpander.Expand(document);

        var titleBlock = Assert.IsType<ComponentElement>(archetyped.Slides[0].Children[0]);
        Assert.Equal("opening-title-2", titleBlock.Id);
        // The authored id is untouched.
        Assert.Equal("opening-title", Assert.IsType<TextElement>(archetyped.Slides[1].Children[0]).Id);
    }
}
