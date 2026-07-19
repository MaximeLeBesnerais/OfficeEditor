using PptxEditor.Core.Generation.Archetypes;
using PptxEditor.Core.Generation.Components;
using PptxEditor.Core.Generation.Emit.Ooxml;
using PptxEditor.Core.Generation.Layout;
using PptxEditor.Core.Generation.Model;
using PptxEditor.Core.Generation.Schema;

namespace DocxEditor.Tests.Generation.Archetypes;

public sealed class ArchetypeExpanderTests
{
    private static DesignTokens Tokens()
        => new()
        {
            Palette = new Dictionary<string, string>
            {
                ["primary"] = "#CEBA80",
                ["accent"] = "#169C9A",
                ["ink"] = "#2C3932",
                ["paper"] = "#FFFFFF",
                ["muted"] = "#7A7567"
            },
            Fonts = new FontTokens { Display = "Aptos Light", Body = "Aptos" },
            Shape = new ShapeTokens { CornerRadius = 0, CardStyle = CardStyle.Flat },
            Metrics = new MetricTokens { MarginPt = 43, GutterPt = 18, TitleSizePt = 40, BodySizePt = 14 }
        };

    // ------------------------------------------------------------------ marker recognition

    [Fact]
    public void IsArchetypeSlide_BareContainerWithSingleArchetypeChild_ReturnsTrue()
    {
        var marker = new ContainerElement
        {
            Children = [new ComponentElement { Name = "cover" }]
        };

        Assert.True(ArchetypeExpander.IsArchetypeSlide(marker, out var component));
        Assert.Equal("cover", component.Name);
    }

    [Fact]
    public void IsArchetypeSlide_ContainerWithLayout_ReturnsFalse()
    {
        var marker = new ContainerElement
        {
            Layout = new LayoutSpec { Mode = LayoutMode.Column },
            Children = [new ComponentElement { Name = "cover" }]
        };

        Assert.False(ArchetypeExpander.IsArchetypeSlide(marker, out _));
    }

    [Fact]
    public void IsArchetypeSlide_ContainerWithFill_ReturnsFalse()
    {
        var marker = new ContainerElement
        {
            Fill = new SolidFill("paper"),
            Children = [new ComponentElement { Name = "cover" }]
        };

        Assert.False(ArchetypeExpander.IsArchetypeSlide(marker, out _));
    }

    [Fact]
    public void IsArchetypeSlide_ContainerWithMultipleChildren_ReturnsFalse()
    {
        var marker = new ContainerElement
        {
            Children =
            [
                new ComponentElement { Name = "cover" },
                new ComponentElement { Name = "section" }
            ]
        };

        Assert.False(ArchetypeExpander.IsArchetypeSlide(marker, out _));
    }

    [Fact]
    public void IsArchetypeSlide_ChildIsNotArchetype_ReturnsFalse()
    {
        var marker = new ContainerElement
        {
            Children = [new ComponentElement { Name = "card" }]
        };

        Assert.False(ArchetypeExpander.IsArchetypeSlide(marker, out _));
    }

    [Fact]
    public void IsArchetypeSlide_ChildIsPrimitive_ReturnsFalse()
    {
        var marker = new ContainerElement
        {
            Children = [new TextElement { Value = "hello" }]
        };

        Assert.False(ArchetypeExpander.IsArchetypeSlide(marker, out _));
    }

    // ------------------------------------------------------------------ expansion of each archetype

    [Fact]
    public void Expand_CoverMarker_ReturnsSlideWithPaperFillAndCenteredLayout()
    {
        var document = DocumentWithMarkers([("cover", """{"title": "Q3 Review", "kicker": "NORTHWIND", "subtitle": "Exec Deck"}""")]);
        var expanded = ArchetypeExpander.Expand(document);

        var slide = GetExpandedSlide(expanded);
        Assert.Equal(new SolidFill("paper"), slide.Fill);
        Assert.Equal(Justify.Center, slide.Layout!.Justify);
        Assert.Equal(LayoutMode.Column, slide.Layout.Mode);
        Assert.True(slide.Children.Count > 0);
    }

    [Fact]
    public void Expand_SectionMarker_WithIndex_ProducesRowWithKpi()
    {
        var document = DocumentWithMarkers([("section", """{"index": "02", "title": "Outlook", "kicker": "OUTLOOK", "subtitle": "Next"}""")]);
        var expanded = ArchetypeExpander.Expand(document);

        var slide = GetExpandedSlide(expanded);
        Assert.Equal(Justify.End, slide.Layout!.Justify);
        Assert.True(slide.Children.Count >= 2);

        var row = Assert.IsType<ContainerElement>(slide.Children[1]);
        Assert.Equal(LayoutMode.Row, row.Layout!.Mode);
        Assert.IsType<ComponentElement>(row.Children[0]); // kpi
        Assert.IsType<ComponentElement>(row.Children[1]); // title_block
    }

    [Fact]
    public void Expand_KpiRowMarker_GrowsIntoRowOfKpis()
    {
        var document = DocumentWithMarkers([("kpi_row",
            """{"title": "At a glance", "subtitle": "vs LY", "kpis": [{"value":"+34%","label":"Revenue"},{"value":"12k","label":"Users"}]}""")]);
        var expanded = ArchetypeExpander.Expand(document);

        var slide = GetExpandedSlide(expanded);
        var row = Assert.IsType<ContainerElement>(slide.Children[1]);
        Assert.Equal(2, row.Children.Count);
        Assert.All(row.Children.Cast<ComponentElement>(), c => Assert.Equal("kpi", c.Name));
    }

    [Fact]
    public void Expand_TwoColMarker_GrowsIntoWeightedComponents()
    {
        var document = DocumentWithMarkers([("two_col",
            """{"title": "Highlights", "weights": [3, 2], "left": {"type": "bullet_list", "content": {"items": ["a","b"]}}, "right": {"type": "card", "content": {"title": "Watch"}}}""")]);
        var expanded = ArchetypeExpander.Expand(document);

        var slide = GetExpandedSlide(expanded);
        var row = Assert.IsType<ContainerElement>(slide.Children[1]);
        Assert.Equal(3, Assert.IsType<ComponentElement>(row.Children[0]).Size!.Grow);
        Assert.Equal(2, Assert.IsType<ComponentElement>(row.Children[1]).Size!.Grow);
    }

    [Fact]
    public void Expand_TableSlideMarker_GrowsIntoTableBlock()
    {
        var document = DocumentWithMarkers([("table_slide",
            """{"title": "Regions", "columns": ["Region","Rev"], "rows": [["NA","8"],["EU","4"]], "columnWeights": [2,1]}""")]);
        var expanded = ArchetypeExpander.Expand(document);

        var slide = GetExpandedSlide(expanded);
        var table = Assert.IsType<ComponentElement>(slide.Children[1]);
        Assert.Equal("table_block", table.Name);
        Assert.Equal(1, table.Size!.Grow);
    }

    // ------------------------------------------------------------------ error paths

    [Fact]
    public void Expand_CoverWithoutTitle_Throws()
    {
        var document = DocumentWithMarkers([("cover", """{"kicker": "KV"}""")]);
        var ex = Assert.Throws<ComponentException>(() => ArchetypeExpander.Expand(document));
        Assert.Contains("title", ex.Message);
    }

    [Fact]
    public void Expand_SectionWithoutTitle_Throws()
    {
        var document = DocumentWithMarkers([("section", """{"index": "01"}""")]);
        var ex = Assert.Throws<ComponentException>(() => ArchetypeExpander.Expand(document));
        Assert.Contains("title", ex.Message);
    }

    [Fact]
    public void Expand_KpiRowWithoutKpis_Throws()
    {
        var document = DocumentWithMarkers([("kpi_row", """{"title": "T", "kpis": []}""")]);
        var ex = Assert.Throws<ComponentException>(() => ArchetypeExpander.Expand(document));
        Assert.Contains("kpis", ex.Message);
    }

    [Fact]
    public void Expand_TwoColMissingSlots_Throws()
    {
        var document = DocumentWithMarkers([("two_col",
            """{"left": {"type": "card", "content": {"title": "A"}}}""")]);
        var ex = Assert.Throws<ComponentException>(() => ArchetypeExpander.Expand(document));
        Assert.Contains("right", ex.Message);
    }

    [Fact]
    public void Expand_TableSlideMissingColumns_Throws()
    {
        var document = DocumentWithMarkers([("table_slide", """{"rows": [["1","2"]]}""")]);
        var ex = Assert.Throws<ComponentException>(() => ArchetypeExpander.Expand(document));
        Assert.Contains("columns", ex.Message);
    }

    [Fact]
    public void Expand_UnknownMarkerName_PassesThroughUnchanged()
    {
        var component = new ComponentElement { Name = "conclusion" };
        var document = new GenerationDocument
        {
            Version = "2.0",
            Design = Tokens(),
            Slides = [new ContainerElement { Children = [component] }]
        };
        var expanded = ArchetypeExpander.Expand(document);
        Assert.Same(component, Assert.Single(expanded.Slides[0].Children));
    }

    [Fact]
    public void Expand_ArchetypeWithSize_Throws()
    {
        var document = new GenerationDocument
        {
            Version = "2.0",
            Design = Tokens(),
            Slides =
            [
                new ContainerElement
                {
                    Children = [new ComponentElement { Name = "cover", Size = new SizeSpec { Width = 100 } }]
                }
            ]
        };
        var ex = Assert.Throws<ComponentException>(() => ArchetypeExpander.Expand(document));
        Assert.Contains("size", ex.Message);
    }

    [Fact]
    public void Expand_NonArchetypeSlides_PassThroughUntouched()
    {
        var text = new TextElement { Value = "hello" };
        var regularSlide = new ContainerElement
        {
            Layout = new LayoutSpec { Mode = LayoutMode.Column },
            Children = [text]
        };
        var document = new GenerationDocument
        {
            Version = "2.0",
            Design = Tokens(),
            Slides = [regularSlide]
        };
        var expanded = ArchetypeExpander.Expand(document);
        Assert.Same(text, expanded.Slides[0].Children[0]);
    }

    // ------------------------------------------------------------------ helpers

    private static GenerationDocument DocumentWithMarkers(params (string name, string contentJson)[] markers)
    {
        var slides = new List<ContainerElement>(markers.Length);
        foreach (var (name, contentJson) in markers)
        {
            var content = string.IsNullOrWhiteSpace(contentJson)
                ? (System.Text.Json.JsonElement?)null
                : System.Text.Json.JsonDocument.Parse(contentJson).RootElement.Clone();
            slides.Add(new ContainerElement
            {
                Children = [new ComponentElement { Name = name, Content = content }]
            });
        }
        return new GenerationDocument { Version = "2.0", Design = Tokens(), Slides = slides };
    }

    private static ContainerElement GetExpandedSlide(GenerationDocument document)
        => Assert.IsType<ContainerElement>(Assert.Single(document.Slides));
}
