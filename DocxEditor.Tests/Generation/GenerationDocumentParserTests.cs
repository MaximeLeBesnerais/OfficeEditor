using PptxEditor.Core.Generation.Model;
using PptxEditor.Core.Generation.Schema;
using PptxEditor.Core.Models;

namespace DocxEditor.Tests.Generation;

public class GenerationDocumentParserTests
{
    private readonly GenerationDocumentParser _parser = new();

    /// <summary>The sample document from plan.md §3.4, with the §3.1 token set filled in.</summary>
    private const string PlanSample = """
        {
          "version": "2.0",
          "design": {
            "palette":  { "primary": "#0B3D91", "accent": "#FF6B00", "ink": "#1A1A1A",
                          "paper": "#FFFFFF", "muted": "#8A94A6" },
            "fonts":    { "display": "Aptos Display", "body": "Aptos" },
            "shape":    { "cornerRadius": 0, "cardStyle": "flat" },
            "metrics":  { "marginPt": 43, "gutterPt": 18, "titleSizePt": 30, "bodySizePt": 14 }
          },
          "slides": [
            {
              "type": "container",
              "layout": { "mode": "row", "gap": 18, "justify": "space-evenly", "align": "center" },
              "children": [
                { "type": "card", "size": { "grow": 1, "aspect": "4:3" },
                  "content": { "title": "+34%", "subtitle": "Revenue" } },
                { "type": "card", "size": { "grow": 1, "aspect": "4:3" },
                  "content": { "title": "12k", "subtitle": "Users" } }
              ]
            }
          ]
        }
        """;

    [Fact]
    public void Validate_PlanSample_ValidatesClean()
    {
        var result = _parser.Validate(PlanSample);

        Assert.True(result.IsValid, string.Join(Environment.NewLine, result.Errors.Select(e => e.ToString())));
        Assert.Empty(result.Warnings);
        Assert.NotNull(result.Document);
    }

    [Fact]
    public void Parse_PlanSample_BuildsExpectedModel()
    {
        var doc = _parser.Parse(PlanSample);

        Assert.Equal("2.0", doc.Version);
        Assert.Equal(960, doc.SlideSize.WidthPt);
        Assert.Equal(540, doc.SlideSize.HeightPt);

        Assert.Equal(5, doc.Design.Palette.Count);
        Assert.Equal("#0B3D91", doc.Design.Palette["primary"]);
        Assert.Equal("Aptos Display", doc.Design.Fonts.Display);
        Assert.Equal("Aptos", doc.Design.Fonts.Body);
        Assert.Equal(0, doc.Design.Shape.CornerRadius);
        Assert.Equal(CardStyle.Flat, doc.Design.Shape.CardStyle);
        Assert.Equal(43, doc.Design.Metrics.MarginPt);
        Assert.Equal(18, doc.Design.Metrics.GutterPt);
        Assert.Equal(30, doc.Design.Metrics.TitleSizePt);
        Assert.Equal(14, doc.Design.Metrics.BodySizePt);

        var slide = Assert.Single(doc.Slides);
        Assert.Null(slide.Size);
        Assert.Null(slide.At);
        Assert.Equal(OverflowPolicy.Error, slide.Overflow);

        var layout = slide.Layout;
        Assert.NotNull(layout);
        Assert.Equal(LayoutMode.Row, layout.Mode);
        Assert.Equal(18, layout.Gap);
        Assert.Equal(Justify.SpaceEvenly, layout.Justify);
        Assert.Equal(AlignItems.Center, layout.Align);

        Assert.Equal(2, slide.Children.Count);
        var first = Assert.IsType<ComponentElement>(slide.Children[0]);
        Assert.Equal("card", first.Name);
        var size = first.Size;
        Assert.NotNull(size);
        Assert.Equal(1, size.Grow);
        var aspect = size.Aspect;
        Assert.NotNull(aspect);
        Assert.Equal(4.0 / 3.0, aspect.Value.Value, precision: 6);
        var content = first.Content;
        Assert.NotNull(content);
        Assert.Equal("+34%", content.Value.GetProperty("title").GetString());
        Assert.Equal("Revenue", content.Value.GetProperty("subtitle").GetString());
        var second = Assert.IsType<ComponentElement>(slide.Children[1]);
        Assert.Equal("12k", second.Content!.Value.GetProperty("title").GetString());
    }

    [Fact]
    public void Parse_MinimalDocument_AppliesDefaults()
    {
        var doc = _parser.Parse("""
            {
              "version": "2.0",
              "design": { "palette": { "primary": "#0B3D91" } },
              "slides": [ { "type": "container", "children": [] } ]
            }
            """);

        Assert.Equal(SlideSize.Widescreen16x9, doc.SlideSize);
        Assert.Null(doc.Design.Fonts.Display);
        Assert.Equal(43, doc.Design.Metrics.MarginPt);
        Assert.Equal(CardStyle.Flat, doc.Design.Shape.CardStyle);
        var slide = Assert.Single(doc.Slides);
        Assert.Null(slide.Layout);
        Assert.Empty(slide.Children);
        Assert.Equal(OverflowPolicy.Error, slide.Overflow);
    }

    [Fact]
    public void Parse_SlideSize4x3_UsesStandardCanvas()
    {
        var doc = _parser.Parse("""
            {
              "version": "2.0",
              "slideSize": "4:3",
              "design": { "palette": {} },
              "slides": [ { "type": "container", "children": [] } ]
            }
            """);

        Assert.Equal(720, doc.SlideSize.WidthPt);
        Assert.Equal(540, doc.SlideSize.HeightPt);
    }

    [Fact]
    public void Parse_TextWithRuns_ParsesStyleAndDefaults()
    {
        var doc = _parser.Parse("""
            {
              "version": "2.0",
              "design": { "palette": { "accent": "#FF6B00" } },
              "slides": [ {
                "type": "container",
                "layout": { "mode": "column" },
                "children": [ {
                  "type": "text",
                  "size": { "grow": 1 },
                  "runs": [
                    { "text": "Revenue ", "bold": true, "fontSize": 20 },
                    { "text": "+34%", "color": "accent", "italic": true }
                  ],
                  "font": "display",
                  "textAlign": "center",
                  "anchor": "middle",
                  "insets": [2, 4]
                } ]
              } ]
            }
            """);

        var text = Assert.IsType<TextElement>(Assert.Single(doc.Slides[0].Children));
        Assert.Null(text.Value);
        Assert.Equal(2, text.Runs!.Count);
        Assert.Equal("Revenue ", text.Runs[0].Text);
        Assert.True(text.Runs[0].Bold);
        Assert.Equal(20, text.Runs[0].FontSize);
        Assert.Equal("accent", text.Runs[1].Color);
        Assert.True(text.Runs[1].Italic);
        Assert.Equal("display", text.Font);
        Assert.Equal(TextAlign.Center, text.TextAlign);
        Assert.Equal(TextAnchor.Middle, text.Anchor);
        Assert.Equal(new EdgeInsets(2, 4, 2, 4), text.Insets);
        Assert.Equal(OverflowPolicy.Shrink, text.Overflow);
    }

    [Fact]
    public void Parse_LayoutlessCanvas_AbsolutePrimitivesValidateClean()
    {
        var result = _parser.Validate("""
            {
              "version": "2.0",
              "design": { "palette": { "primary": "#0B3D91", "paper": "#FFFFFF", "ink": "#1A1A1A" } },
              "slides": [ {
                "type": "container",
                "fill": "paper",
                "children": [
                  { "type": "rect", "at": { "x": 10, "y": 10 }, "size": { "w": 100, "h": 60 },
                    "fill": "primary", "stroke": { "color": "ink", "width": 1.5 },
                    "radius": { "tl": 8, "tr": 8 },
                    "shadow": { "color": "ink", "dx": 2, "dy": 3, "blur": 6, "alpha": 0.4 } },
                  { "type": "ellipse", "at": { "x": 120, "y": 10 }, "size": { "w": 40, "h": 40 },
                    "fill": { "angle": 90, "stops": [
                      { "color": "primary", "offset": 0 },
                      { "color": "ink", "offset": 1, "alpha": 0.5 } ] } },
                  { "type": "line", "at": { "x": 10, "y": 80 }, "size": { "w": 200, "h": 1 },
                    "orientation": "horizontal", "stroke": "ink" },
                  { "type": "connector", "at": { "x": 10, "y": 90 }, "size": { "w": 1, "h": 100 },
                    "orientation": "vertical" },
                  { "type": "image", "at": { "x": 220, "y": 10 }, "size": { "w": 160, "h": 90 },
                    "src": "images/logo.png", "fit": "crop",
                    "crop": { "left": 0, "top": 5000, "right": 0, "bottom": 5000 }, "alt": "Logo" },
                  { "type": "group", "at": { "x": 10, "y": 200 }, "size": { "w": 100, "h": 100 },
                    "children": [
                      { "type": "rect", "at": { "x": 0, "y": 0 }, "size": { "w": 50, "h": 50 }, "fill": "primary" }
                    ] },
                  { "type": "text", "at": { "x": 10, "y": 320 }, "size": { "w": 300, "h": 40 },
                    "text": "Absolute", "overflow": "clip" }
                ]
              } ]
            }
            """);

        Assert.True(result.IsValid, string.Join(Environment.NewLine, result.Errors.Select(e => e.ToString())));
        Assert.Empty(result.Warnings);

        var slide = Assert.Single(result.Document!.Slides);
        Assert.IsType<SolidFill>(slide.Fill);
        Assert.Equal(7, slide.Children.Count);

        var rect = Assert.IsType<RectElement>(slide.Children[0]);
        Assert.Equal(new PointSpec(10, 10), rect.At);
        Assert.Equal(new CornerRadii(8, 8, 0, 0), rect.Radius);
        Assert.Equal(1.5, rect.Stroke!.WidthPt);
        Assert.Equal(0.4, rect.Shadow!.Alpha);

        var ellipse = Assert.IsType<EllipseElement>(slide.Children[1]);
        var gradient = Assert.IsType<LinearGradientFill>(ellipse.Fill);
        Assert.Equal(90, gradient.Angle);
        Assert.Equal(2, gradient.Stops.Count);
        Assert.Equal(0.5, gradient.Stops[1].Alpha);

        var line = Assert.IsType<LineElement>(slide.Children[2]);
        Assert.False(line.IsConnector);
        Assert.Equal("ink", line.Stroke!.Color);
        Assert.Equal(1, line.Stroke.WidthPt);
        Assert.True(Assert.IsType<LineElement>(slide.Children[3]).IsConnector);

        var image = Assert.IsType<ImageElement>(slide.Children[4]);
        Assert.Equal(ImageFitMode.Crop, image.Fit);
        Assert.Equal(new SourceRect(0, 5000, 0, 5000), image.Crop);

        var group = Assert.IsType<GroupElement>(slide.Children[5]);
        Assert.Single(group.Children);

        var text = Assert.IsType<TextElement>(slide.Children[6]);
        Assert.Equal(OverflowPolicy.Clip, text.Overflow);
    }

    [Fact]
    public void Parse_LayoutlessChild_WithAspectAndOneDimension_ResolvesBoth()
    {
        var result = _parser.Validate("""
            {
              "version": "2.0",
              "design": { "palette": {} },
              "slides": [ {
                "type": "container",
                "children": [
                  { "type": "image", "at": { "x": 0, "y": 0 }, "size": { "w": 160, "aspect": "16:9" }, "src": "a.png" }
                ]
              } ]
            }
            """);

        Assert.True(result.IsValid, string.Join(Environment.NewLine, result.Errors.Select(e => e.ToString())));
    }

    [Fact]
    public void Parse_PaddingForms_AllAccepted()
    {
        var doc = _parser.Parse("""
            {
              "version": "2.0",
              "design": { "palette": {} },
              "slides": [ {
                "type": "container",
                "layout": { "mode": "column", "gap": 4 },
                "children": [
                  { "type": "container", "padding": 8, "children": [] },
                  { "type": "container", "padding": [4, 12], "children": [] },
                  { "type": "container", "padding": [1, 2, 3, 4], "children": [] }
                ]
              } ]
            }
            """);

        var children = doc.Slides[0].Children.Cast<ContainerElement>().ToList();
        Assert.Equal(new EdgeInsets(8, 8, 8, 8), children[0].Padding);
        Assert.Equal(new EdgeInsets(4, 12, 4, 12), children[1].Padding);
        Assert.Equal(new EdgeInsets(1, 2, 3, 4), children[2].Padding);
    }

    [Fact]
    public void Parse_GridLayout_ParsesColumnsAndGaps()
    {
        var doc = _parser.Parse("""
            {
              "version": "2.0",
              "design": { "palette": {} },
              "slides": [ {
                "type": "container",
                "layout": { "mode": "grid", "cols": 3, "gap": 8, "rowGap": 12, "columnGap": 6, "justify": "space-between", "align": "end" },
                "children": []
              } ]
            }
            """);

        var layout = doc.Slides[0].Layout!;
        Assert.Equal(LayoutMode.Grid, layout.Mode);
        Assert.Equal(3, layout.Columns);
        Assert.Equal(12, layout.RowGap);
        Assert.Equal(6, layout.ColumnGap);
        Assert.Equal(Justify.SpaceBetween, layout.Justify);
        Assert.Equal(AlignItems.End, layout.Align);
    }

    [Fact]
    public void Validate_RawHexColor_IsValidButWarns()
    {
        var result = _parser.Validate("""
            {
              "version": "2.0",
              "design": { "palette": { "primary": "#0B3D91" } },
              "slides": [ {
                "type": "container",
                "layout": { "mode": "row" },
                "children": [
                  { "type": "rect", "size": { "grow": 1 }, "fill": "#FF0000" }
                ]
              } ]
            }
            """);

        Assert.True(result.IsValid, string.Join(Environment.NewLine, result.Errors.Select(e => e.ToString())));
        var warning = Assert.Single(result.Warnings);
        Assert.Equal(GenerationIssueSeverity.Warning, warning.Severity);
        Assert.Contains("off-palette", warning.Message);
        Assert.EndsWith(".fill", warning.Path);
    }

    [Fact]
    public void Validate_MultipleErrors_AllCollectedInOnePass()
    {
        var result = _parser.Validate("""
            {
              "version": "1.0",
              "design": { "palette": { "primary": "blue" } },
              "slides": [ {
                "type": "container",
                "layout": { "mode": "grid" },
                "children": [ { "type": "squiggle" } ]
              } ]
            }
            """);

        Assert.False(result.IsValid);
        Assert.Null(result.Document);
        Assert.Equal(4, result.Errors.Count);
        Assert.Contains(result.Errors, e => e.Path == "$.version");
        Assert.Contains(result.Errors, e => e.Path == "$.design.palette.primary");
        Assert.Contains(result.Errors, e => e.Path == "$.slides[0].layout");
        Assert.Contains(result.Errors, e => e.Path == "$.slides[0].children[0]");
    }

    [Fact]
    public void Parse_InvalidDocument_ThrowsWithEveryIssue()
    {
        var ex = Assert.Throws<GenerationValidationException>(() => _parser.Parse("""
            { "version": "2.0", "design": { "palette": {} }, "slides": [ { "type": "container", "childern": [] } ] }
            """));

        Assert.Equal(2, ex.Issues.Count);
        Assert.Contains("childern", ex.Message);
        Assert.Contains("Did you mean 'children'?", ex.Message);
    }

    [Fact]
    public void Parse_PropertyNames_AreCaseInsensitive()
    {
        var doc = _parser.Parse("""
            {
              "Version": "2.0",
              "Design": { "Palette": {} },
              "Slides": [ { "Type": "container", "Layout": { "Mode": "Row" }, "Children": [] } ]
            }
            """);

        Assert.Equal(LayoutMode.Row, doc.Slides[0].Layout!.Mode);
    }
}
