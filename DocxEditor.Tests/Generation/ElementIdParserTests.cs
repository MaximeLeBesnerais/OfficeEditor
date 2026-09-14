using System.Text.Json;
using PptxEditor.Core.Generation.Model;
using PptxEditor.Core.Generation.Schema;

namespace DocxEditor.Tests.Generation;

/// <summary>
/// SPEC §9.2 (stable element ids): the generation vocabulary accepts an optional 'id' on
/// every element type — loud JSON-path errors for non-string, malformed and duplicate ids,
/// and a deterministic per-type fallback so un-ided elements stay addressable.
/// </summary>
public sealed class ElementIdParserTests
{
    private readonly GenerationDocumentParser _parser = new();

    [Fact]
    public void Validate_EveryElementType_AcceptsExplicitIds()
    {
        var result = _parser.Validate("""
            {
              "version": "2.0",
              "design": {
                "palette": { "primary": "#0B3D91", "accent": "#FF6B00", "ink": "#1A1A1A",
                             "paper": "#FFFFFF", "muted": "#8A94A6" }
              },
              "slides": [
                {
                  "type": "container",
                  "id": "slide-hero",
                  "layout": { "mode": "row" },
                  "children": [
                    { "type": "container", "id": "col", "children": [] },
                    { "type": "text", "id": "headline", "text": "Hi", "size": { "w": 100, "h": 40 } },
                    { "type": "rect", "id": "chip", "size": { "w": 10, "h": 10 } },
                    { "type": "ellipse", "id": "dot", "size": { "w": 8, "h": 8 } },
                    { "type": "line", "id": "rule", "size": { "w": 50, "h": 1 } },
                    { "type": "connector", "id": "link", "size": { "w": 1, "h": 50 } },
                    { "type": "image", "id": "logo", "src": "logo.png", "size": { "w": 40, "h": 40 } },
                    { "type": "group", "id": "cluster", "size": { "w": 60, "h": 60 },
                      "children": [
                        { "type": "rect", "id": "cluster-inner", "at": { "x": 0, "y": 0 }, "size": { "w": 10, "h": 10 } }
                      ] },
                    { "type": "card", "id": "revenue", "content": { "title": "+34%" } },
                    { "type": "kpi", "id": "users", "content": { "value": "12k", "label": "Users" } }
                  ]
                },
                {
                  "type": "cover",
                  "id": "opening",
                  "content": { "title": "Northwind Labs", "kicker": "Q3 2026" }
                }
              ]
            }
            """);

        Assert.True(result.IsValid, string.Join(Environment.NewLine, result.Errors.Select(e => e.ToString())));
        var document = result.Document!;

        var slide = document.Slides[0];
        Assert.Equal("slide-hero", slide.Id);
        Assert.Equal("col", Assert.IsType<ContainerElement>(slide.Children[0]).Id);
        Assert.Equal("headline", Assert.IsType<TextElement>(slide.Children[1]).Id);
        Assert.Equal("chip", Assert.IsType<RectElement>(slide.Children[2]).Id);
        Assert.Equal("dot", Assert.IsType<EllipseElement>(slide.Children[3]).Id);
        Assert.Equal("rule", Assert.IsType<LineElement>(slide.Children[4]).Id);
        Assert.Equal("link", Assert.IsType<LineElement>(slide.Children[5]).Id);
        Assert.Equal("logo", Assert.IsType<ImageElement>(slide.Children[6]).Id);
        var group = Assert.IsType<GroupElement>(slide.Children[7]);
        Assert.Equal("cluster", group.Id);
        Assert.Equal("cluster-inner", Assert.IsType<RectElement>(Assert.Single(group.Children)).Id);
        Assert.Equal("revenue", Assert.IsType<ComponentElement>(slide.Children[8]).Id);
        Assert.Equal("users", Assert.IsType<ComponentElement>(slide.Children[9]).Id);

        // Archetype slides: the id lands on both the slide root and its marker component.
        var archetypeSlide = document.Slides[1];
        Assert.Equal("opening", archetypeSlide.Id);
        var marker = Assert.IsType<ComponentElement>(Assert.Single(archetypeSlide.Children));
        Assert.Equal("opening", marker.Id);
    }

    [Fact]
    public void Validate_NonStringId_RejectedAtJsonPath()
    {
        var result = _parser.Validate("""
            {
              "version": "2.0",
              "design": { "palette": {} },
              "slides": [ {
                "type": "container",
                "layout": { "mode": "row" },
                "children": [ { "type": "text", "id": 42, "text": "x", "size": { "w": 10, "h": 10 } } ]
              } ]
            }
            """);

        Assert.False(result.IsValid);
        Assert.Null(result.Document);
        Assert.Contains(result.Errors,
            e => e.Path == "$.slides[0].children[0].id" && e.Message.Contains("must be a string", StringComparison.Ordinal));
    }

    [Fact]
    public void Validate_DuplicateId_RejectedWithBothPaths()
    {
        var result = _parser.Validate("""
            {
              "version": "2.0",
              "design": { "palette": {} },
              "slides": [ {
                "type": "container",
                "layout": { "mode": "row" },
                "children": [
                  { "type": "text", "id": "dup", "text": "a", "size": { "w": 10, "h": 10 } },
                  { "type": "text", "id": "dup", "text": "b", "size": { "w": 10, "h": 10 } }
                ]
              } ]
            }
            """);

        Assert.False(result.IsValid);
        Assert.Null(result.Document);
        var duplicate = Assert.Single(result.Errors);
        Assert.Equal("$.slides[0].children[1].id", duplicate.Path);
        Assert.Contains("duplicate id 'dup'", duplicate.Message);
        Assert.Contains("first declared at $.slides[0].children[0]", duplicate.Message);
    }

    [Fact]
    public void Validate_DuplicateIdAcrossSlides_Rejected()
    {
        var result = _parser.Validate("""
            {
              "version": "2.0",
              "design": { "palette": {} },
              "slides": [
                { "type": "container", "layout": { "mode": "row" },
                  "children": [ { "type": "text", "id": "dup", "text": "a", "size": { "w": 10, "h": 10 } } ] },
                { "type": "container", "layout": { "mode": "row" },
                  "children": [ { "type": "text", "id": "dup", "text": "b", "size": { "w": 10, "h": 10 } } ] }
              ]
            }
            """);

        Assert.False(result.IsValid);
        Assert.Contains(result.Errors,
            e => e.Path == "$.slides[1].children[0].id" && e.Message.Contains("duplicate id 'dup'", StringComparison.Ordinal));
    }

    [Fact]
    public void Validate_InvalidIdFormat_RejectedAtJsonPath()
    {
        var result = _parser.Validate("""
            {
              "version": "2.0",
              "design": { "palette": {} },
              "slides": [ {
                "type": "container",
                "layout": { "mode": "row" },
                "children": [ { "type": "text", "id": "Hello World!", "text": "x", "size": { "w": 10, "h": 10 } } ]
              } ]
            }
            """);

        Assert.False(result.IsValid);
        Assert.Contains(result.Errors,
            e => e.Path == "$.slides[0].children[0].id"
                 && e.Message.Contains("not a valid id", StringComparison.Ordinal));
    }

    [Fact]
    public void Parse_WithoutIds_SynthesizesDeterministicFallbacks()
    {
        var document = _parser.Parse("""
            {
              "version": "2.0",
              "design": { "palette": {} },
              "slides": [ {
                "type": "container",
                "layout": { "mode": "row" },
                "children": [
                  { "type": "text", "text": "a", "size": { "w": 10, "h": 10 } },
                  { "type": "text", "text": "b", "size": { "w": 10, "h": 10 } },
                  { "type": "card", "content": { "title": "T" } }
                ]
              } ]
            }
            """);

        var slide = Assert.Single(document.Slides);
        Assert.Equal("slide-0", slide.Id);
        Assert.Equal("text-0", Assert.IsType<TextElement>(slide.Children[0]).Id);
        Assert.Equal("text-1", Assert.IsType<TextElement>(slide.Children[1]).Id);
        Assert.Equal("card-0", Assert.IsType<ComponentElement>(slide.Children[2]).Id);
    }

    [Fact]
    public void Parse_UnidedAndExplicitIds_DoNotCollide()
    {
        // The explicit "text-1" is registered, so the synthesized fallback for the un-ided
        // text skips used ids and lands deterministically on the first free one — collision-free.
        var document = _parser.Parse("""
            {
              "version": "2.0",
              "design": { "palette": {} },
              "slides": [ {
                "type": "container",
                "layout": { "mode": "row" },
                "children": [
                  { "type": "text", "id": "text-1", "text": "a", "size": { "w": 10, "h": 10 } },
                  { "type": "text", "text": "b", "size": { "w": 10, "h": 10 } }
                ]
              } ]
            }
            """);

        var slide = Assert.Single(document.Slides);
        Assert.Equal("text-1", Assert.IsType<TextElement>(slide.Children[0]).Id);
        Assert.Equal("text-0", Assert.IsType<TextElement>(slide.Children[1]).Id);
    }

    [Fact]
    public void Schema_DeclaresIdOnEveryElementTypeAndArchetypeMarker()
    {
        using var document = JsonDocument.Parse(GenerationSchema.SchemaJson);
        var defs = document.RootElement.GetProperty("$defs");
        foreach (var defName in new[] { "container", "text", "rect", "ellipse", "line", "image", "group", "component", "archetypeSlide" })
        {
            var props = defs.GetProperty(defName).GetProperty("properties");
            Assert.True(props.TryGetProperty("id", out var idProp), $"$defs.{defName} is missing 'id'.");
            Assert.Equal("#/$defs/id", idProp.GetProperty("$ref").GetString());
        }

        var slidesItems = document.RootElement.GetProperty("properties").GetProperty("slides").GetProperty("items");
        var refs = slidesItems.GetProperty("oneOf").EnumerateArray()
            .Select(item => item.GetProperty("$ref").GetString())
            .ToList();
        Assert.Contains("#/$defs/container", refs);
        Assert.Contains("#/$defs/archetypeSlide", refs);
    }
}
