using PptxEditor.Core.Generation.Layout;
using PptxEditor.Core.Generation.Schema;

namespace DocxEditor.Tests.Generation.Layout;

/// <summary>
/// SPEC §9.1 (layout introspection): the resolved tree exposes a flat paint-order element
/// list per slide (ids, kinds, final rects) and an id → rect lookup — both riding the single
/// existing layout pass, no extra geometry. Group children are flattened depth-first and are
/// already in absolute coordinates.
/// </summary>
public sealed class LayoutIntrospectionTests
{
    private readonly GenerationDocumentParser _parser = new();

    [Fact]
    public void StructuredSlide_FlatList_HasExpectedRectsAndKinds()
    {
        var layout = Resolve("""
            {
              "version": "2.0",
              "design": { "palette": { "primary": "#0B3D91" } },
              "slides": [ {
                "type": "container",
                "id": "root",
                "layout": { "mode": "column", "gap": 10 },
                "padding": 20,
                "children": [
                  { "type": "text", "id": "title", "text": "T", "size": { "h": 40 } },
                  { "type": "rect", "id": "bar", "size": { "h": 30 }, "fill": "primary" }
                ]
              } ]
            }
            """);

        var slide = Assert.Single(layout.Slides);
        Assert.Equal("root", slide.Id);

        // Paint order = document order: root container, then children depth-first.
        var elements = slide.Elements;
        Assert.Equal(new[] { "root", "title", "bar" }, elements.Select(e => e.Id));
        Assert.Equal(new[] { ResolvedElementType.Container, ResolvedElementType.Text, ResolvedElementType.Rect },
            elements.Select(e => e.Type));

        // Column layout with padding 20 and gap 10: children stretch to the content width.
        AssertRect(elements[0], 0, 0, 960, 540);
        AssertRect(elements[1], 20, 20, 920, 40);
        AssertRect(elements[2], 20, 70, 920, 30);

        // The underlying resolved tree carries the same ids.
        Assert.Equal("root", slide.Root.Id);
        Assert.Equal("title", slide.Root.Children[0].Id);
        Assert.Equal("bar", slide.Root.Children[1].Id);
    }

    [Fact]
    public void TryGetRect_ById_ReturnsTheResolvedRect()
    {
        var layout = Resolve("""
            {
              "version": "2.0",
              "design": { "palette": { "primary": "#0B3D91" } },
              "slides": [ {
                "type": "container",
                "layout": { "mode": "column", "gap": 10 },
                "padding": 20,
                "children": [
                  { "type": "text", "id": "title", "text": "T", "size": { "h": 40 } },
                  { "type": "rect", "id": "bar", "size": { "h": 30 }, "fill": "primary" }
                ]
              } ]
            }
            """);

        Assert.True(layout.TryGetRect(0, "title", out var titleRect));
        Assert.Equal(new ElementRect(20, 20, 920, 40), titleRect);

        Assert.True(layout.TryGetRect(0, "bar", out var barRect));
        Assert.Equal(new ElementRect(20, 70, 920, 30), barRect);

        // Per-slide lookup mirrors the document-level one.
        Assert.True(layout.Slides[0].TryGetRect("title", out var perSlide));
        Assert.Equal(titleRect, perSlide);
    }

    [Fact]
    public void TryGetRect_UnknownIdOrSlideIndex_ReturnsFalse()
    {
        var layout = Resolve("""
            {
              "version": "2.0",
              "design": { "palette": { "primary": "#0B3D91" } },
              "slides": [ {
                "type": "container",
                "layout": { "mode": "column" },
                "children": [ { "type": "text", "id": "known", "text": "T", "size": { "h": 40 } } ]
              } ]
            }
            """);

        Assert.False(layout.TryGetRect(0, "nope", out _));
        Assert.False(layout.TryGetRect(0, null, out _));
        Assert.False(layout.TryGetRect(3, "known", out _));
        Assert.False(layout.TryGetRect(-1, "known", out _));
        Assert.False(layout.Slides[0].TryGetRect("nope", out _));
    }

    [Fact]
    public void GroupChildren_AreFlattenedWithAbsoluteCoordinates()
    {
        var layout = Resolve("""
            {
              "version": "2.0",
              "design": { "palette": { "primary": "#0B3D91", "ink": "#1A1A1A" } },
              "slides": [ {
                "type": "container",
                "id": "canvas",
                "children": [
                  {
                    "type": "group",
                    "id": "g1",
                    "at": { "x": 100, "y": 50 },
                    "size": { "w": 200, "h": 100 },
                    "children": [
                      { "type": "rect", "id": "inner", "at": { "x": 10, "y": 10 },
                        "size": { "w": 50, "h": 50 }, "fill": "ink" }
                    ]
                  }
                ]
              } ]
            }
            """);

        var slide = Assert.Single(layout.Slides);

        // Depth-first paint order: root, group, then the group's child.
        Assert.Equal(new[] { "canvas", "g1", "inner" }, slide.Elements.Select(e => e.Id));
        Assert.Equal(ResolvedElementType.Group, slide.Elements[1].Type);

        // The group child is already absolute in the resolved tree, so flattening needs no
        // extra math: 100 + 10 = 110, 50 + 10 = 60.
        AssertRect(slide.Elements[1], 100, 50, 200, 100);
        AssertRect(slide.Elements[2], 110, 60, 50, 50);

        Assert.True(layout.TryGetRect(0, "inner", out var inner));
        Assert.Equal(new ElementRect(110, 60, 50, 50), inner);
    }

    private static void AssertRect(ResolvedElementInfo element, double x, double y, double w, double h)
    {
        Assert.Equal(x, element.X, precision: 3);
        Assert.Equal(y, element.Y, precision: 3);
        Assert.Equal(w, element.Width, precision: 3);
        Assert.Equal(h, element.Height, precision: 3);
    }

    private LayoutResult Resolve(string json)
        => new LayoutResolver().Resolve(_parser.Parse(json));
}
