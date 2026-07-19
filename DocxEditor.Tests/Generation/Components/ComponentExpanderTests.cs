using System.Text.Json;
using PptxEditor.Core.Generation.Components;
using PptxEditor.Core.Generation.Model;
using PptxEditor.Core.Models;

namespace DocxEditor.Tests.Generation.Components;

/// <summary>
/// Structural tests for the v1 component set (plan.md §4): each component expands to the
/// expected primitive subtree with token-driven styling and the author's size/at
/// transferred to the expansion root.
/// </summary>
public sealed class ComponentExpanderTests
{
    private const string Path = "slides[0].children[0]";

    private static DesignTokens Tokens(CardStyle cardStyle = CardStyle.Flat, double cornerRadius = 8)
        => new()
        {
            Palette = new Dictionary<string, string>
            {
                ["primary"] = "#0B3D91",
                ["accent"] = "#FF6B00",
                ["ink"] = "#1A1A1A",
                ["paper"] = "#FFFFFF",
                ["muted"] = "#8A94A6"
            },
            Fonts = new FontTokens { Display = "Aptos Display", Body = "Aptos" },
            Shape = new ShapeTokens { CornerRadius = cornerRadius, CardStyle = cardStyle },
            Metrics = new MetricTokens() // 43 / 18 / 30 / 14 (plan.md §3.1 defaults)
        };

    private static JsonElement Content(string json) => JsonDocument.Parse(json).RootElement.Clone();

    private static ComponentElement Component(string name, string contentJson, SizeSpec? size = null, PointSpec? at = null)
        => new() { Name = name, Content = Content(contentJson), Size = size, At = at };

    private static GenElement Expand(string name, string contentJson, DesignTokens? tokens = null, SizeSpec? size = null, PointSpec? at = null)
        => ComponentExpander.ExpandElement(Component(name, contentJson, size, at), tokens ?? Tokens(), Path);

    // ------------------------------------------------------------------ card

    [Fact]
    public void Card_ExpandsToSurfaceColumnWithFields()
    {
        var size = new SizeSpec { Grow = 1, Aspect = AspectRatio.Of(4, 3) };
        var root = Assert.IsType<ContainerElement>(Expand("card",
            """{"title": "+34%", "subtitle": "Revenue", "body": "vs last year"}""", size: size));

        Assert.Same(size, root.Size);
        Assert.Equal(new SolidFill("paper"), root.Fill);
        Assert.Null(root.Stroke);
        Assert.Null(root.Shadow);
        Assert.Equal(CornerRadii.All(8), root.Radius);
        Assert.Equal(EdgeInsets.All(18), root.Padding);
        Assert.Equal(LayoutMode.Column, root.Layout!.Mode);
        Assert.Equal(4, root.Layout.Gap);
        Assert.Equal(OverflowPolicy.Error, root.Overflow);

        Assert.Equal(3, root.Children.Count);
        var title = Assert.IsType<TextElement>(root.Children[0]);
        Assert.Equal("+34%", title.Value);
        Assert.Equal("display", title.Font);
        Assert.Equal(16, title.FontSize);
        Assert.True(title.Bold);
        Assert.Equal("ink", title.Color);
        Assert.Equal(21.6, title.Size!.Height);

        var subtitle = Assert.IsType<TextElement>(root.Children[1]);
        Assert.Equal("muted", subtitle.Color);
        Assert.Equal(18.9, subtitle.Size!.Height);

        // The last present field grows to fill the card.
        var body = Assert.IsType<TextElement>(root.Children[2]);
        Assert.Equal(1, body.Size!.Grow);
    }

    [Fact]
    public void Card_TitleOnly_GrowsTitle()
    {
        var root = Assert.IsType<ContainerElement>(Expand("card", """{"title": "Solo"}"""));
        var title = Assert.IsType<TextElement>(Assert.Single(root.Children));
        Assert.Equal(1, title.Size!.Grow);
    }

    [Fact]
    public void Card_OutlineStyle_AddsMutedStroke()
    {
        var root = Assert.IsType<ContainerElement>(
            Expand("card", """{"title": "t"}""", tokens: Tokens(CardStyle.Outline)));
        Assert.Equal(new StrokeSpec { Color = "muted", WidthPt = 1 }, root.Stroke);
        Assert.Null(root.Shadow);
    }

    [Fact]
    public void Card_ShadowStyle_AddsSoftInkShadow()
    {
        var root = Assert.IsType<ContainerElement>(
            Expand("card", """{"title": "t"}""", tokens: Tokens(CardStyle.Shadow)));
        Assert.NotNull(root.Shadow);
        var shadow = root.Shadow;
        Assert.Equal("ink", shadow.Color);
        Assert.Equal(2, shadow.Dy);
        Assert.Equal(8, shadow.Blur);
        Assert.Equal(0.18, shadow.Alpha);
        Assert.Null(root.Stroke);
    }

    // ------------------------------------------------------------------ kpi

    [Fact]
    public void Kpi_ExpandsToCenteredValueLabelDelta()
    {
        var root = Assert.IsType<ContainerElement>(Expand("kpi",
            """{"value": "+34%", "label": "Revenue", "delta": "+12% vs LY"}"""));

        Assert.Equal(Justify.Center, root.Layout!.Justify);
        Assert.Equal(AlignItems.Center, root.Layout.Align);
        Assert.Equal(new SolidFill("paper"), root.Fill);

        var value = Assert.IsType<TextElement>(root.Children[0]);
        Assert.Equal("+34%", value.Value);
        Assert.Equal("display", value.Font);
        Assert.Equal(30, value.FontSize);
        Assert.True(value.Bold);
        Assert.Equal("primary", value.Color);
        Assert.Equal(TextAlign.Center, value.TextAlign);

        var label = Assert.IsType<TextElement>(root.Children[1]);
        Assert.Equal("Revenue", label.Value);
        Assert.Equal("muted", label.Color);

        var delta = Assert.IsType<TextElement>(root.Children[2]);
        Assert.Equal("accent", delta.Color);
        Assert.Equal(12, delta.FontSize);
    }

    [Fact]
    public void Kpi_WithoutDelta_HasTwoChildren()
    {
        var root = Assert.IsType<ContainerElement>(Expand("kpi", """{"value": "12k", "label": "Users"}"""));
        Assert.Equal(2, root.Children.Count);
    }

    // ------------------------------------------------------------------ title_block

    [Fact]
    public void TitleBlock_KickerTitleSubtitle_InOrder()
    {
        var root = Assert.IsType<ContainerElement>(Expand("title_block",
            """{"kicker": "Q3 REPORT", "title": "Growth", "subtitle": "All regions"}"""));

        Assert.Null(root.Fill);
        Assert.Equal(LayoutMode.Column, root.Layout!.Mode);

        var kicker = Assert.IsType<TextElement>(root.Children[0]);
        Assert.Equal("accent", kicker.Color);
        Assert.True(kicker.Bold);
        Assert.Equal(12, kicker.FontSize);

        var title = Assert.IsType<TextElement>(root.Children[1]);
        Assert.Equal("display", title.Font);
        Assert.Equal(30, title.FontSize);
        Assert.Equal("ink", title.Color);

        var subtitle = Assert.IsType<TextElement>(root.Children[2]);
        Assert.Equal("muted", subtitle.Color);
    }

    // ------------------------------------------------------------------ bullet_list

    [Fact]
    public void BulletList_ItemsBecomeMarkerRows()
    {
        var root = Assert.IsType<ContainerElement>(Expand("bullet_list",
            """{"title": "Highlights", "items": ["One", "Two"]}"""));

        Assert.Equal(LayoutMode.Column, root.Layout!.Mode);
        Assert.Equal(9, root.Layout.Gap); // gutter / 2
        Assert.Equal(3, root.Children.Count);

        Assert.IsType<TextElement>(root.Children[0]);

        var row = Assert.IsType<ContainerElement>(root.Children[1]);
        Assert.Equal(18.9, row.Size!.Height);
        Assert.Equal(LayoutMode.Row, row.Layout!.Mode);
        Assert.Equal(8, row.Layout.Gap);
        Assert.Equal(AlignItems.Center, row.Layout.Align);

        var marker = Assert.IsType<EllipseElement>(row.Children[0]);
        Assert.Equal(new SizeSpec { Width = 6, Height = 6 }, marker.Size);
        Assert.Equal(new SolidFill("accent"), marker.Fill);

        var text = Assert.IsType<TextElement>(row.Children[1]);
        Assert.Equal("One", text.Value);
        Assert.Equal(1, text.Size!.Grow);
        Assert.Equal(TextAnchor.Middle, text.Anchor);
    }

    [Fact]
    public void BulletList_MarkerColorOverride()
    {
        var root = Assert.IsType<ContainerElement>(Expand("bullet_list",
            """{"items": ["One"], "markerColor": "primary"}"""));
        var row = Assert.IsType<ContainerElement>(root.Children[0]);
        var marker = Assert.IsType<EllipseElement>(row.Children[0]);
        Assert.Equal(new SolidFill("primary"), marker.Fill);
    }

    // ------------------------------------------------------------------ divider

    [Fact]
    public void Divider_DefaultsToMutedOnePointHorizontalRule()
    {
        var line = Assert.IsType<LineElement>(Expand("divider", "{}"));

        Assert.Equal(LineOrientation.Horizontal, line.Orientation);
        Assert.Equal(new StrokeSpec { Color = "muted", WidthPt = 1 }, line.Stroke);
        Assert.Equal(1, line.Size!.Height); // cross-axis defaults to the stroke width
        Assert.Null(line.Size.Width);
    }

    [Fact]
    public void Divider_AuthorSizeWins()
    {
        var line = Assert.IsType<LineElement>(ComponentExpander.ExpandElement(
            Component("divider", """{"color": "primary", "width": 2.5, "orientation": "vertical"}""",
                size: new SizeSpec { Height = 100 }),
            Tokens(), Path));

        Assert.Equal(LineOrientation.Vertical, line.Orientation);
        Assert.Equal(2.5, line.Stroke!.WidthPt);
        Assert.Equal("primary", line.Stroke.Color);
        Assert.Equal(2.5, line.Size!.Width); // synthesized cross-axis
        Assert.Equal(100, line.Size.Height); // author-fixed main axis untouched
    }

    // ------------------------------------------------------------------ badge

    [Fact]
    public void Badge_IsAPillWithCenteredLabel()
    {
        var root = Assert.IsType<ContainerElement>(Expand("badge", """{"text": "NEW"}"""));

        var expectedHeight = 22.2; // fontSize 12 + 2×3 padding
        Assert.Equal(expectedHeight, root.Size!.Height);
        Assert.Equal(CornerRadii.All(expectedHeight / 2), root.Radius); // pill: radius = h/2
        Assert.Equal(new SolidFill("accent"), root.Fill);
        Assert.Equal(Justify.Center, root.Layout!.Justify);

        var text = Assert.IsType<TextElement>(Assert.Single(root.Children));
        Assert.Equal("NEW", text.Value);
        Assert.True(text.Bold);
        Assert.Equal("paper", text.Color);
        Assert.Equal(TextAlign.Center, text.TextAlign);
        Assert.Equal(1, text.Size!.Grow);
    }

    [Fact]
    public void Badge_AuthorHeightAndColorsWin()
    {
        var root = Assert.IsType<ContainerElement>(ComponentExpander.ExpandElement(
            Component("badge", """{"text": "B", "color": "primary", "textColor": "ink"}""",
                size: new SizeSpec { Width = 60, Height = 20 }),
            Tokens(), Path));

        Assert.Equal(20, root.Size!.Height);
        Assert.Equal(60, root.Size.Width);
        Assert.Equal(CornerRadii.All(10), root.Radius);
        Assert.Equal(new SolidFill("primary"), root.Fill);
        Assert.Equal("ink", Assert.IsType<TextElement>(root.Children[0]).Color);
    }

    // ------------------------------------------------------------------ image_card

    [Fact]
    public void ImageCard_ImageGrowsAboveFixedCaption()
    {
        var root = Assert.IsType<ContainerElement>(Expand("image_card",
            """{"src": "team.png", "title": "The team", "subtitle": "2026"}"""));

        Assert.Equal(new SolidFill("paper"), root.Fill);
        Assert.Null(root.Padding); // image bleeds to the card edge

        var image = Assert.IsType<ImageElement>(root.Children[0]);
        Assert.Equal("team.png", image.Source);
        Assert.Equal(ImageFitMode.Crop, image.Fit);
        Assert.Equal(1, image.Size!.Grow);

        var caption = Assert.IsType<ContainerElement>(root.Children[1]);
        const double expectedHeight = 2 * 18 + 21.6 + 18.9 + 4; // padding + title line + subtitle line + gap
        Assert.Equal(expectedHeight, caption.Size!.Height);
        Assert.Equal(EdgeInsets.All(18), caption.Padding);
        Assert.Equal(2, caption.Children.Count);
        Assert.Equal("The team", Assert.IsType<TextElement>(caption.Children[0]).Value);
    }

    [Fact]
    public void ImageCard_WithoutCaption_IsImageOnly()
    {
        var root = Assert.IsType<ContainerElement>(Expand("image_card", """{"src": "a.png", "fit": "contain", "alt": "A"}"""));
        var image = Assert.IsType<ImageElement>(Assert.Single(root.Children));
        Assert.Equal(ImageFitMode.Contain, image.Fit);
        Assert.Equal("A", image.Alt);
    }

    // ------------------------------------------------------------------ table_block

    [Fact]
    public void TableBlock_HeaderAndBodyRows()
    {
        var root = Assert.IsType<ContainerElement>(Expand("table_block",
            """{"columns": ["Region", "Rev"], "rows": [["EU", "12"], ["US", "9"]], "columnWeights": [2, 1]}"""));

        Assert.Equal(LayoutMode.Column, root.Layout!.Mode);
        Assert.Equal(3, root.Children.Count);

        var header = Assert.IsType<ContainerElement>(root.Children[0]);
        Assert.Equal(new SolidFill("primary"), header.Fill);
        Assert.Equal(TableBlockComponent.DefaultRowHeight(14), header.Size!.Height);

        var headerCell = Assert.IsType<ContainerElement>(header.Children[0]);
        Assert.Equal(2, headerCell.Size!.Grow); // column weight
        Assert.Equal(EdgeInsets.Symmetric(6, 8), headerCell.Padding);
        var headerText = Assert.IsType<TextElement>(Assert.Single(headerCell.Children));
        Assert.Equal("Region", headerText.Value);
        Assert.Equal("paper", headerText.Color);
        Assert.True(headerText.Bold);
        Assert.Equal(TextAnchor.Middle, headerText.Anchor);

        var bodyRow = Assert.IsType<ContainerElement>(root.Children[1]);
        Assert.Null(bodyRow.Fill);
        var bodyCell = Assert.IsType<ContainerElement>(bodyRow.Children[1]);
        Assert.Equal(1, bodyCell.Size!.Grow);
        var bodyText = Assert.IsType<TextElement>(Assert.Single(bodyCell.Children));
        Assert.Equal("12", bodyText.Value);
        Assert.Equal("ink", bodyText.Color);
        Assert.False(bodyText.Bold);
    }

    [Fact]
    public void TableBlock_HeaderFalse_OmitsHeaderRow()
    {
        var root = Assert.IsType<ContainerElement>(Expand("table_block",
            """{"columns": ["A"], "rows": [["1"], ["2"]], "header": false}"""));
        Assert.Equal(2, root.Children.Count);
        Assert.All(root.Children, row => Assert.Null(Assert.IsType<ContainerElement>(row).Fill));
    }

    [Fact]
    public void TableBlock_RowHeightOverride()
    {
        var root = Assert.IsType<ContainerElement>(Expand("table_block",
            """{"columns": ["A"], "rows": [], "rowHeight": 40}"""));
        var header = Assert.IsType<ContainerElement>(Assert.Single(root.Children));
        Assert.Equal(40, header.Size!.Height);
    }

    // ------------------------------------------------------------------ plumbing

    [Fact]
    public void Expansion_TransfersSizeAndAtToRoot()
    {
        var at = new PointSpec(10, 20);
        var size = new SizeSpec { Width = 100, Height = 80 };
        var root = Assert.IsType<ContainerElement>(ComponentExpander.ExpandElement(
            Component("kpi", """{"value": "1", "label": "L"}""", size: size, at: at), Tokens(), Path));
        Assert.Same(size, root.Size);
        Assert.Equal(at, root.At);
    }

    [Fact]
    public void Expander_RecursesIntoContainersAndLeavesPrimitivesUntouched()
    {
        var rect = new RectElement { Size = new SizeSpec { Width = 10, Height = 10 }, Fill = new SolidFill("accent") };
        var slide = new ContainerElement
        {
            Layout = new LayoutSpec { Mode = LayoutMode.Column },
            Children =
            [
                rect,
                new ContainerElement
                {
                    Layout = new LayoutSpec { Mode = LayoutMode.Row },
                    Children =
                    [
                        Component("badge", """{"text": "X"}""", size: new SizeSpec { Width = 40 }),
                        new GroupElement
                        {
                            Size = new SizeSpec { Width = 60, Height = 30 },
                            Children = [Component("divider", "{}", size: new SizeSpec { Width = 60 })]
                        }
                    ]
                }
            ]
        };

        var document = new GenerationDocument { Version = "2.0", Design = Tokens(), Slides = [slide] };
        var expanded = ComponentExpander.Expand(document);

        var root = expanded.Slides[0];
        Assert.Same(rect, root.Children[0]); // primitives untouched (same instance)
        var row = Assert.IsType<ContainerElement>(root.Children[1]);
        Assert.IsType<ContainerElement>(row.Children[0]); // badge → pill container
        var group = Assert.IsType<GroupElement>(row.Children[1]);
        Assert.IsType<LineElement>(group.Children[0]); // divider → line, even inside groups
    }

    [Fact]
    public void EveryComponent_TextsShrink_ContainersError()
    {
        var cases = new (string Name, string Content)[]
        {
            ("card", """{"title": "T", "subtitle": "S", "body": "B"}"""),
            ("kpi", """{"value": "V", "label": "L", "delta": "D"}"""),
            ("title_block", """{"kicker": "K", "title": "T", "subtitle": "S"}"""),
            ("bullet_list", """{"title": "T", "items": ["a", "b"]}"""),
            ("divider", "{}"),
            ("badge", """{"text": "X"}"""),
            ("image_card", """{"src": "a.png", "title": "T", "subtitle": "S"}"""),
            ("table_block", """{"columns": ["A"], "rows": [["1"]]}""")
        };

        foreach (var (name, content) in cases)
        {
            var root = Expand(name, content);
            foreach (var element in Flatten(root))
            {
                switch (element)
                {
                    case TextElement text:
                        Assert.True(text.Overflow == OverflowPolicy.Shrink,
                            $"{name}: text '{text.Value}' must shrink (plan.md §3.2).");
                        break;
                    case ContainerElement container:
                        Assert.True(container.Overflow == OverflowPolicy.Error,
                            $"{name}: a container must error on structural overflow (plan.md §3.2).");
                        break;
                }
            }
        }
    }

    private static IEnumerable<GenElement> Flatten(GenElement element)
    {
        yield return element;
        var children = element switch
        {
            ContainerElement c => c.Children,
            GroupElement g => g.Children,
            _ => null
        };
        if (children is null)
        {
            yield break;
        }
        foreach (var child in children)
        {
            foreach (var descendant in Flatten(child))
            {
                yield return descendant;
            }
        }
    }
}
