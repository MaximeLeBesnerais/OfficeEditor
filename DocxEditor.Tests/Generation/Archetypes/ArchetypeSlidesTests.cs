using PptxEditor.Core.Generation.Archetypes;
using PptxEditor.Core.Generation.Components;
using PptxEditor.Core.Generation.Model;

namespace DocxEditor.Tests.Generation.Archetypes;

public sealed class ArchetypeSlidesTests
{
    private static DesignTokens Tokens(CardStyle cardStyle = CardStyle.Flat)
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
            Shape = new ShapeTokens { CornerRadius = 0, CardStyle = cardStyle },
            Metrics = new MetricTokens { MarginPt = 43, GutterPt = 18, TitleSizePt = 40, BodySizePt = 14 }
        };

    // ------------------------------------------------------------------ missing property helpers

    private static CoverContent Cover(string title, string? kicker = null, string? subtitle = null)
        => new() { Title = title, Kicker = kicker, Subtitle = subtitle };

    private static SectionContent Section(string title, string? index = null, string? kicker = null, string? subtitle = null)
        => new() { Title = title, Index = index, Kicker = kicker, Subtitle = subtitle };

    private static KpiRowContent KpiRow(string? title = null, string? subtitle = null, params (string value, string label, string? delta)[] kpis)
        => new()
        {
            Title = title,
            Subtitle = subtitle,
            Kpis = [.. kpis.Select(k => new KpiItem { Value = k.value, Label = k.label, Delta = k.delta })]
        };

    private static TwoColContent TwoCol(
        ComponentElement left,
        ComponentElement right,
        string? title = null,
        string? subtitle = null,
        IReadOnlyList<double>? weights = null)
        => new() { Title = title, Subtitle = subtitle, Left = left, Right = right, Weights = weights };

    private static ComponentElement Component(string name, string contentJson)
    {
        var content = System.Text.Json.JsonDocument.Parse(contentJson).RootElement.Clone();
        return new ComponentElement { Name = name, Content = content };
    }

    // ------------------------------------------------------------------ names

    [Fact]
    public void Names_HasAllFiveArchetypes()
    {
        var expected = new[] { "cover", "section", "kpi_row", "two_col", "table_slide" };
        foreach (var name in expected)
        {
            Assert.True(ArchetypeSlides.Names.Contains(name), $"missing '{name}'");
        }
        Assert.Equal(5, ArchetypeSlides.Names.Count);
    }

    [Fact]
    public void IsArchetype_KnownNamesReturnTrue()
    {
        Assert.True(ArchetypeSlides.IsArchetype("cover"));
        Assert.True(ArchetypeSlides.IsArchetype("section"));
        Assert.True(ArchetypeSlides.IsArchetype("kpi_row"));
        Assert.True(ArchetypeSlides.IsArchetype("two_col"));
        Assert.True(ArchetypeSlides.IsArchetype("table_slide"));
        Assert.False(ArchetypeSlides.IsArchetype("card"));
        Assert.False(ArchetypeSlides.IsArchetype("unknown"));
    }

    // ------------------------------------------------------------------ cover

    [Fact]
    public void Cover_ProducesCenteredSlideWithTitleBlockAndDivider()
    {
        var content = Cover("Q3 Business Review", kicker: "NORTHWIND", subtitle: "Executive team");
        var slide = ArchetypeSlides.Cover(content, Tokens());

        var root = Assert.IsType<ContainerElement>(slide);
        Assert.Equal(new SolidFill("paper"), root.Fill);
        Assert.Equal(Justify.Center, root.Layout!.Justify);
        Assert.Equal(LayoutMode.Column, root.Layout.Mode);

        Assert.Equal(2, root.Children.Count);
        var titleBlock = Assert.IsType<ComponentElement>(root.Children[0]);
        Assert.Equal("title_block", titleBlock.Name);
        var divider = Assert.IsType<ComponentElement>(root.Children[1]);
        Assert.Equal("divider", divider.Name);
    }

    [Fact]
    public void Cover_MissingTitle_ThrowsComponentException()
    {
        var ex = Assert.Throws<ComponentException>(() => ArchetypeSlides.Cover(Cover(""), Tokens()));
        Assert.Contains("title", ex.Message);
    }

    [Fact]
    public void Cover_NullTitle_ThrowsComponentException()
    {
        var content = Cover("  ");
        var ex = Assert.Throws<ComponentException>(() => ArchetypeSlides.Cover(content, Tokens()));
        Assert.Contains("title", ex.Message);
    }

    [Fact]
    public void Cover_NullContent_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>(() => ArchetypeSlides.Cover(null!, Tokens()));
    }

    // ------------------------------------------------------------------ section

    [Fact]
    public void Section_WithIndex_ProducesRowWithKpiAndTitleBlock()
    {
        var content = Section("What is next", index: "02", kicker: "OUTLOOK", subtitle: "Three bets");
        var slide = ArchetypeSlides.Section(content, Tokens());

        var root = Assert.IsType<ContainerElement>(slide);
        Assert.Equal(Justify.End, root.Layout!.Justify);

        Assert.Equal(2, root.Children.Count);
        var divider = Assert.IsType<ComponentElement>(root.Children[0]);
        Assert.Equal("divider", divider.Name);

        var row = Assert.IsType<ContainerElement>(root.Children[1]);
        Assert.Equal(LayoutMode.Row, row.Layout!.Mode);
        Assert.Equal(2, row.Children.Count);
        var kpi = Assert.IsType<ComponentElement>(row.Children[0]);
        Assert.Equal("kpi", kpi.Name);
        var titleBlock = Assert.IsType<ComponentElement>(row.Children[1]);
        Assert.Equal("title_block", titleBlock.Name);
    }

    [Fact]
    public void Section_WithoutIndex_ProducesTitleBlockOnly()
    {
        var content = Section("Outlook", kicker: "KV", subtitle: "Next steps");
        var slide = ArchetypeSlides.Section(content, Tokens());

        Assert.Equal(2, slide.Children.Count);
        Assert.IsType<ComponentElement>(slide.Children[1]); // title block without row
    }

    [Fact]
    public void Section_MissingTitle_ThrowsComponentException()
    {
        var ex = Assert.Throws<ComponentException>(() => ArchetypeSlides.Section(Section(""), Tokens()));
        Assert.Contains("title", ex.Message);
    }

    // ------------------------------------------------------------------ kpi_row

    [Fact]
    public void KpiRow_WithFourKpis_ProducesRowOfKpiComponents()
    {
        var content = KpiRow(title: "At a glance", subtitle: "vs last year",
            ("+34%", "Revenue", "+12% vs LY"),
            ("12.4k", "Users", "+8% vs LY"),
            ("98.2%", "Uptime", "+0.4 pts"),
            ("4.1d", "Cycle", "-1.2 days"));

        var slide = ArchetypeSlides.KpiRow(content, Tokens());

        var root = Assert.IsType<ContainerElement>(slide);
        Assert.Equal(LayoutMode.Column, root.Layout!.Mode);
        Assert.Equal(Justify.Start, root.Layout.Justify);

        Assert.Equal(2, root.Children.Count);
        var titleBlock = Assert.IsType<ComponentElement>(root.Children[0]);
        Assert.Equal("title_block", titleBlock.Name);

        var row = Assert.IsType<ContainerElement>(root.Children[1]);
        Assert.Equal(LayoutMode.Row, row.Layout!.Mode);
        Assert.Equal(4, row.Children.Count);
        Assert.All(row.Children.Cast<ComponentElement>(), c => Assert.Equal("kpi", c.Name));
    }

    [Fact]
    public void KpiRow_WithoutTitle_OmitsTitleBlock()
    {
        var content = KpiRow(kpis: ("+34%", "Revenue", null));
        var slide = ArchetypeSlides.KpiRow(content, Tokens());
        Assert.Single(slide.Children);
    }

    [Fact]
    public void KpiRow_EmptyKpis_ThrowsComponentException()
    {
        var content = KpiRow(title: "T", kpis: []);
        var ex = Assert.Throws<ComponentException>(() => ArchetypeSlides.KpiRow(content, Tokens()));
        Assert.Contains("kpis", ex.Message);
    }

    [Fact]
    public void KpiRow_MoreThanSixKpis_ThrowsComponentException()
    {
        var content = KpiRow(kpis: Enumerable.Range(1, 7).Select(i => ($"{i}", $"L{i}", (string?)null)).ToArray());
        var ex = Assert.Throws<ComponentException>(() => ArchetypeSlides.KpiRow(content, Tokens()));
        Assert.Contains("kpis", ex.Message);
    }

    [Fact]
    public void KpiRow_MissingValue_ThrowsComponentException()
    {
        var content = new KpiRowContent { Kpis = [new KpiItem { Value = "", Label = "L" }] };
        var ex = Assert.Throws<ComponentException>(() => ArchetypeSlides.KpiRow(content, Tokens()));
        Assert.Contains("value", ex.Message);
    }

    [Fact]
    public void KpiRow_MissingLabel_ThrowsComponentException()
    {
        var content = new KpiRowContent { Kpis = [new KpiItem { Value = "V", Label = "  " }] };
        var ex = Assert.Throws<ComponentException>(() => ArchetypeSlides.KpiRow(content, Tokens()));
        Assert.Contains("label", ex.Message);
    }

    // ------------------------------------------------------------------ two_col

    [Fact]
    public void TwoCol_WithBulletListAndCard_ProducesWeightedRow()
    {
        var content = TwoCol(
            left: Component("bullet_list", """{"title": "Highlights", "items": ["One", "Two"]}"""),
            right: Component("card", """{"title": "Watch", "subtitle": "Costs", "body": "Details"}"""),
            title: "What moved",
            subtitle: "Q3 analysis",
            weights: [3, 2]);

        var slide = ArchetypeSlides.TwoCol(content, Tokens());

        var root = Assert.IsType<ContainerElement>(slide);
        Assert.Equal(LayoutMode.Column, root.Layout!.Mode);

        Assert.Equal(2, root.Children.Count);
        var heading = Assert.IsType<ComponentElement>(root.Children[0]);
        Assert.Equal("title_block", heading.Name);

        var row = Assert.IsType<ContainerElement>(root.Children[1]);
        Assert.Equal(LayoutMode.Row, row.Layout!.Mode);
        Assert.Equal(2, row.Children.Count);

        var left = Assert.IsType<ComponentElement>(row.Children[0]);
        Assert.Equal("bullet_list", left.Name);
        Assert.Equal(3, left.Size!.Grow);

        var right = Assert.IsType<ComponentElement>(row.Children[1]);
        Assert.Equal("card", right.Name);
        Assert.Equal(2, right.Size!.Grow);
    }

    [Fact]
    public void TwoCol_DefaultWeightsAreEqual()
    {
        var content = TwoCol(
            Component("card", """{"title": "A"}"""),
            Component("card", """{"title": "B"}"""));
        var slide = ArchetypeSlides.TwoCol(content, Tokens());
        var row = Assert.IsType<ContainerElement>(slide.Children[0]);
        Assert.Equal(1, Assert.IsType<ComponentElement>(row.Children[0]).Size!.Grow);
        Assert.Equal(1, Assert.IsType<ComponentElement>(row.Children[1]).Size!.Grow);
    }

    [Fact]
    public void TwoCol_NullLeft_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>(() => ArchetypeSlides.TwoCol(
            TwoCol(null!, Component("card", """{"title": "B"}""")), Tokens()));
    }

    [Fact]
    public void TwoCol_NullRight_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>(() => ArchetypeSlides.TwoCol(
            TwoCol(Component("card", """{"title": "A"}"""), null!), Tokens()));
    }

    [Fact]
    public void TwoCol_UnknownComponentInSlot_ThrowsComponentException()
    {
        var content = TwoCol(
            Component("unknown", "{}"),
            Component("card", """{"title": "B"}"""));
        var ex = Assert.Throws<ComponentException>(() => ArchetypeSlides.TwoCol(content, Tokens()));
        Assert.Contains("unknown component", ex.Message);
        Assert.Contains("unknown", ex.Message);
    }

    [Fact]
    public void TwoCol_ArchetypeInSlot_ThrowsComponentException()
    {
        var content = TwoCol(
            Component("cover", "{}"),
            Component("card", """{"title": "B"}"""));
        var ex = Assert.Throws<ComponentException>(() => ArchetypeSlides.TwoCol(content, Tokens()));
        Assert.Contains("unknown component", ex.Message);
    }

    [Fact]
    public void TwoCol_SizeOnSlot_ThrowsComponentException()
    {
        var slot = new ComponentElement
        {
            Name = "card",
            Content = System.Text.Json.JsonDocument.Parse("""{"title": "B"}""").RootElement.Clone(),
            Size = new SizeSpec { Grow = 2 }
        };
        var content = TwoCol(Component("card", """{"title": "A"}"""), slot);
        var ex = Assert.Throws<ComponentException>(() => ArchetypeSlides.TwoCol(content, Tokens()));
        Assert.Contains("size", ex.Message);
    }

    [Fact]
    public void TwoCol_WrongWeightsCount_ThrowsComponentException()
    {
        var content = TwoCol(
            Component("card", """{"title": "A"}"""),
            Component("card", """{"title": "B"}"""),
            weights: [1]);
        var ex = Assert.Throws<ComponentException>(() => ArchetypeSlides.TwoCol(content, Tokens()));
        Assert.Contains("weights", ex.Message);
    }

    [Fact]
    public void TwoCol_ZeroWeight_ThrowsComponentException()
    {
        var content = TwoCol(
            Component("card", """{"title": "A"}"""),
            Component("card", """{"title": "B"}"""),
            weights: [0, 1]);
        var ex = Assert.Throws<ComponentException>(() => ArchetypeSlides.TwoCol(content, Tokens()));
        Assert.Contains("weights", ex.Message);
    }

    // ------------------------------------------------------------------ table_slide

    [Fact]
    public void TableSlide_WithColumnsAndRows_ProducesTableBlock()
    {
        var content = new TableSlideContent
        {
            Title = "Regional performance",
            Columns = ["Region", "Revenue", "Growth", "Pipeline"],
            Rows = [["NA", "$8.2M", "+38%", "$5.1M"], ["EMEA", "$4.6M", "+41%", "$3.8M"]],
            ColumnWeights = [2, 1, 1, 1]
        };

        var slide = ArchetypeSlides.TableSlide(content, Tokens());

        var root = Assert.IsType<ContainerElement>(slide);
        Assert.Equal(LayoutMode.Column, root.Layout!.Mode);
        Assert.Equal(Justify.Start, root.Layout.Justify);

        Assert.Equal(2, root.Children.Count);
        var titleBlock = Assert.IsType<ComponentElement>(root.Children[0]);
        Assert.Equal("title_block", titleBlock.Name);

        var table = Assert.IsType<ComponentElement>(root.Children[1]);
        Assert.Equal("table_block", table.Name);
        Assert.Equal(1, table.Size!.Grow);
    }

    [Fact]
    public void TableSlide_WithoutTitle_OmitsTitleBlock()
    {
        var content = new TableSlideContent
        {
            Columns = ["A", "B"],
            Rows = [["1", "2"]]
        };
        var slide = ArchetypeSlides.TableSlide(content, Tokens());
        var table = Assert.IsType<ComponentElement>(Assert.Single(slide.Children));
        Assert.Equal("table_block", table.Name);
    }

    [Fact]
    public void TableSlide_EmptyColumns_ThrowsComponentException()
    {
        var content = new TableSlideContent { Columns = [], Rows = [["1"]] };
        var ex = Assert.Throws<ComponentException>(() => ArchetypeSlides.TableSlide(content, Tokens()));
        Assert.Contains("columns", ex.Message);
    }

    [Fact]
    public void TableSlide_NullColumns_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>(() => ArchetypeSlides.TableSlide(
            new TableSlideContent { Columns = null!, Rows = [] }, Tokens()));
    }

    // ------------------------------------------------------------------ overflow safety: no component left unexpanded

    [Fact]
    public void EveryArchetypeRoot_IsContainerWithOnlyComponentsAndContainers()
    {
        var design = Tokens();
        var cases = new Func<ContainerElement>[]
        {
            () => ArchetypeSlides.Cover(Cover("T"), design),
            () => ArchetypeSlides.Section(Section("T", index: "01"), design),
            () => ArchetypeSlides.KpiRow(KpiRow(kpis: ("V", "L", null)), design),
            () => ArchetypeSlides.TwoCol(TwoCol(Component("card", """{"title":"A"}"""), Component("card", """{"title":"B"}""")), design),
            () => ArchetypeSlides.TableSlide(new TableSlideContent { Columns = ["A"], Rows = [["1"]] }, design)
        };

        foreach (var build in cases)
        {
            var slide = build();
            Assert.NotNull(slide.Layout);
            foreach (var child in AllNodes(slide))
            {
                Assert.True(child is ContainerElement or ComponentElement,
                    $"archetype slide tree must only contain containers and component nodes; found {child.GetType().Name}");
            }
        }
    }

    private static IEnumerable<GenElement> AllNodes(GenElement root)
    {
        yield return root;
        var children = root switch
        {
            ContainerElement c => c.Children,
            ComponentElement => null,
            _ => null
        };
        if (children is null)
        {
            yield break;
        }
        foreach (var child in children)
        {
            foreach (var descendant in AllNodes(child))
            {
                yield return descendant;
            }
        }
    }
}
