using System.Text.Json;
using PptxEditor.Core.Generation.Components;
using PptxEditor.Core.Generation.Layout;
using PptxEditor.Core.Generation.Model;
using PptxEditor.Core.Generation.Schema;

namespace DocxEditor.Tests.Generation.Components;

/// <summary>
/// Validation and pipeline tests for the component layer (plan.md §3.4 loud errors, §4):
/// malformed content fails with actionable <see cref="ComponentException"/>s (path +
/// "Did you mean …?"), and a component-bearing document flows parse → expand → layout.
/// </summary>
public sealed class ComponentErrorTests
{
    private const string Path = "slides[0].children[2]";

    private static DesignTokens Tokens(bool withoutPrimary = false)
    {
        var palette = new Dictionary<string, string>
        {
            ["primary"] = "#0B3D91",
            ["accent"] = "#FF6B00",
            ["ink"] = "#1A1A1A",
            ["paper"] = "#FFFFFF",
            ["muted"] = "#8A94A6"
        };
        if (withoutPrimary)
        {
            palette.Remove("primary");
        }
        return new DesignTokens { Palette = palette };
    }

    private static ComponentElement Component(string name, string contentJson)
        => new() { Name = name, Content = JsonDocument.Parse(contentJson).RootElement.Clone() };

    private static ComponentException ExpandError(string name, string contentJson, DesignTokens? tokens = null)
        => Assert.Throws<ComponentException>(
            () => ComponentExpander.ExpandElement(Component(name, contentJson), tokens ?? Tokens(), Path));

    [Fact]
    public void UnknownProperty_SuggestsCorrection()
    {
        var ex = ExpandError("kpi", """{"value": "+34%", "lable": "Revenue"}""");
        Assert.Equal(Path, ex.Path);
        Assert.Contains("unknown property 'lable'", ex.Message);
        Assert.Contains("Did you mean 'label'?", ex.Message);
        Assert.Contains("slides[0].children[2].content.lable", ex.Message);
    }

    [Fact]
    public void MissingRequiredProperty_IsLoud()
    {
        var ex = ExpandError("card", """{"subtitle": "only"}""");
        Assert.Contains("'title' is required", ex.Message);
    }

    [Fact]
    public void MissingContent_IsLoud()
    {
        var ex = Assert.Throws<ComponentException>(() => ComponentExpander.ExpandElement(
            new ComponentElement { Name = "kpi", Content = null }, Tokens(), Path));
        Assert.Contains("'content' is required for component 'kpi'", ex.Message);
    }

    [Fact]
    public void WrongPropertyType_IsLoud()
    {
        var ex = ExpandError("card", """{"title": 42}""");
        Assert.Contains("content.title", ex.Message);
        Assert.Contains("must be a string", ex.Message);
    }

    [Fact]
    public void PercentageValue_IsRejectedAsV1NonGoal()
    {
        var ex = ExpandError("divider", """{"width": "50%"}""");
        Assert.Contains("percentages are not supported", ex.Message);
    }

    [Fact]
    public void UnknownEnumValue_ListsValidValues()
    {
        var ex = ExpandError("divider", """{"orientation": "diagonal"}""");
        Assert.Contains("horizontal", ex.Message);
        Assert.Contains("vertical", ex.Message);
    }

    [Fact]
    public void UnknownColorToken_SuggestsPaletteNames()
    {
        var ex = ExpandError("badge", """{"text": "X", "color": "accnet"}""");
        Assert.Contains("unknown color 'accnet'", ex.Message);
        Assert.Contains("Did you mean 'accent'?", ex.Message);
    }

    [Fact]
    public void MissingPaletteToken_FailsAtExpansionWithActionableMessage()
    {
        var ex = ExpandError("kpi", """{"value": "1", "label": "L"}""", Tokens(withoutPrimary: true));
        Assert.Contains("requires palette token(s) 'primary'", ex.Message);
        Assert.Contains("declared tokens", ex.Message);
    }

    [Fact]
    public void UnknownComponent_IsLoud()
    {
        var ex = Assert.Throws<ComponentException>(() => ComponentExpander.ExpandElement(
            new ComponentElement { Name = "carousel", Content = null }, Tokens(), Path));
        Assert.Contains("unknown component 'carousel'", ex.Message);
        Assert.Contains("card", ex.Message); // known names are listed
    }

    [Fact]
    public void TableBlock_RowCellCountMismatch_IsLoud()
    {
        var ex = ExpandError("table_block", """{"columns": ["A", "B"], "rows": [["1"]]}""");
        Assert.Contains("rows[0]", ex.Message);
        Assert.Contains("exactly 2 cell(s)", ex.Message);
    }

    [Fact]
    public void TableBlock_ColumnWeightsCountMismatch_IsLoud()
    {
        var ex = ExpandError("table_block", """{"columns": ["A", "B"], "rows": [], "columnWeights": [1]}""");
        Assert.Contains("exactly 2 weight(s)", ex.Message);
    }

    [Fact]
    public void BulletList_RequiresAtLeastOneItem()
    {
        var ex = ExpandError("bullet_list", """{"items": []}""");
        Assert.Contains("at least 1 item(s)", ex.Message);
    }

    [Fact]
    public void MultipleErrors_AreCollectedIntoOneException()
    {
        var ex = ExpandError("kpi", """{"value": 1, "lable": "x"}""");
        Assert.Contains("lable", ex.Message);
        Assert.Contains("label' is required", ex.Message);
    }

    [Fact]
    public void EndToEnd_ComponentDocumentParsesExpandsAndResolves()
    {
        const string json = """
        {
          "version": "2.0",
          "design": {
            "palette": { "primary": "#0B3D91", "accent": "#FF6B00", "ink": "#1A1A1A",
                         "paper": "#FFFFFF", "muted": "#8A94A6" },
            "fonts": { "display": "Aptos Display", "body": "Aptos" },
            "shape": { "cornerRadius": 8, "cardStyle": "flat" },
            "metrics": { "marginPt": 43, "gutterPt": 18, "titleSizePt": 30, "bodySizePt": 14 }
          },
          "slides": [
            {
              "type": "container",
              "padding": 43,
              "layout": { "mode": "column", "gap": 12 },
              "children": [
                { "type": "title_block", "size": { "h": 84 },
                  "content": { "kicker": "Q3 REPORT", "title": "Growth", "subtitle": "All regions" } },
                { "type": "divider", "content": {} },
                { "type": "container", "size": { "grow": 3 },
                  "layout": { "mode": "row", "gap": 18 },
                  "children": [
                    { "type": "kpi", "size": { "grow": 1 },
                      "content": { "value": "+34%", "label": "Revenue", "delta": "+12% vs LY" } },
                    { "type": "kpi", "size": { "grow": 1 },
                      "content": { "value": "12k", "label": "Users" } },
                    { "type": "image_card", "size": { "grow": 1 },
                      "content": { "src": "team.png", "title": "The team" } }
                  ]
                },
                { "type": "bullet_list", "size": { "h": 76 },
                  "content": { "items": ["Enterprise up 34%", "12k active users", "Churn below 2%"] } },
                { "type": "table_block", "size": { "grow": 2 },
                  "content": { "columns": ["Region", "Revenue"], "rows": [["EU", "12"]] } },
                { "type": "container", "size": { "h": 24 },
                  "layout": { "mode": "row" },
                  "children": [
                    { "type": "badge", "size": { "w": 80 }, "content": { "text": "CONFIDENTIAL" } }
                  ]
                }
              ]
            }
          ]
        }
        """;

        var document = new GenerationDocumentParser().Parse(json);
        var expanded = ComponentExpander.Expand(document);
        var result = new LayoutResolver().Resolve(expanded);

        Assert.Empty(result.Warnings);
        var slide = Assert.Single(result.Slides);
        // Every expanded subtree resolves to plain primitives — no component ever reaches layout.
        Assert.All(Flatten(slide.Root), e => Assert.IsNotType<ComponentElement>(e));
    }

    [Fact]
    public void UnexpandedComponent_IsRejectedByLayout() // the pipeline contract, plan.md §4
    {
        var document = new GenerationDocument
        {
            Version = "2.0",
            Design = Tokens(),
            Slides =
            [
                new ContainerElement
                {
                    Layout = new LayoutSpec { Mode = LayoutMode.Column },
                    Children = [new ComponentElement { Name = "card", Size = new SizeSpec { Height = 100 }, Content = null }]
                }
            ]
        };

        Assert.Throws<LayoutException>(() => new LayoutResolver().Resolve(document));
    }

    private static IEnumerable<object> Flatten(ResolvedElement element)
    {
        yield return element;
        if (element is ResolvedContainer container)
        {
            foreach (var child in container.Children)
            {
                foreach (var descendant in Flatten(child))
                {
                    yield return descendant;
                }
            }
        }
    }
}
