using PptxEditor.Core.Generation.Layout;
using PptxEditor.Core.Generation.Model;

namespace DocxEditor.Tests.Generation.Layout;

/// <summary>
/// Loud-failure and seam tests for <see cref="LayoutResolver"/> (plan.md §3.2 overflow
/// policies, §3.4 loud errors). Models are built in C# to bypass the JSON validator, which
/// already rejects some of these shapes — the resolver must defend itself.
/// </summary>
public sealed class LayoutResolverErrorTests
{
    private static readonly DesignTokens Design = new()
    {
        Palette = new Dictionary<string, string> { ["primary"] = "#0B3D91" }
    };

    private static LayoutResult Resolve(ContainerElement slide, ITextMeasurer? measurer = null)
        => new LayoutResolver(measurer).Resolve(new GenerationDocument
        {
            Version = "2.0",
            Design = Design,
            Slides = [slide]
        });

    private static ContainerElement RootRow(params GenElement[] children)
        => new() { Layout = new LayoutSpec { Mode = LayoutMode.Row }, Children = children };

    private static RectElement Rect(double? w = null, double? h = null, double? grow = null, string? aspect = null)
        => new()
        {
            Size = new SizeSpec
            {
                Width = w,
                Height = h,
                Grow = grow,
                Aspect = aspect is null ? null : AspectRatio.Of(
                    double.Parse(aspect.Split(':')[0]),
                    double.Parse(aspect.Split(':')[1]))
            }
        };

    [Fact]
    public void Resolve_ChildrenExceedRow_ErrorPolicyThrows()
    {
        var root = RootRow(Rect(w: 600, h: 100), Rect(w: 600, h: 100));

        var ex = Assert.Throws<LayoutException>(() => Resolve(root));
        Assert.Contains("slides[0]", ex.Message);
        Assert.Contains("overflow", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Resolve_ChildrenExceedRow_ClipPolicyResolvesBeyondEdge()
    {
        var root = RootRow(Rect(w: 600, h: 100), Rect(w: 600, h: 100));
        var clipped = root with { Overflow = OverflowPolicy.Clip };

        var result = Resolve(clipped);

        var second = result.Slides[0].Root.Children[1];
        Assert.Equal(600, second.X);
        Assert.Equal(1200, second.X + second.Width); // extends past the 960 slide edge, unclipped by layout
    }

    [Fact]
    public void Resolve_ChildExceedsCrossAxis_ErrorPolicyThrows()
    {
        var root = RootRow(Rect(w: 100, h: 600));

        var ex = Assert.Throws<LayoutException>(() => Resolve(root));
        Assert.Contains("cross", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Resolve_GrowWithFixedMainSize_Throws()
    {
        var root = RootRow(Rect(w: 100, grow: 1));

        var ex = Assert.Throws<LayoutException>(() => Resolve(root));
        Assert.Contains("grow", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Resolve_AspectWithBothDimensionsFixed_Throws()
    {
        var root = RootRow(Rect(w: 160, h: 90, aspect: "16:9"));

        var ex = Assert.Throws<LayoutException>(() => Resolve(root));
        Assert.Contains("over-constrained", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Resolve_UndeterminedMainSize_Throws()
    {
        var root = RootRow(new RectElement());

        var ex = Assert.Throws<LayoutException>(() => Resolve(root));
        Assert.Contains("undetermined", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Resolve_GrowInsideGrid_Throws()
    {
        var root = new ContainerElement
        {
            Layout = new LayoutSpec { Mode = LayoutMode.Grid, Columns = 2 },
            Children = [Rect(grow: 1)]
        };

        var ex = Assert.Throws<LayoutException>(() => Resolve(root));
        Assert.Contains("grid", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Resolve_GridWithoutColumns_Throws()
    {
        var root = new ContainerElement
        {
            Layout = new LayoutSpec { Mode = LayoutMode.Grid },
            Children = [Rect(w: 100, h: 100)]
        };

        var ex = Assert.Throws<LayoutException>(() => Resolve(root));
        Assert.Contains("cols", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Resolve_ComponentNotExpanded_Throws()
    {
        var root = RootRow(new ComponentElement { Name = "card", Size = new SizeSpec { Width = 100, Height = 100 } });

        var ex = Assert.Throws<LayoutException>(() => Resolve(root));
        Assert.Contains("expanded", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("card", ex.Message);
    }

    [Fact]
    public void Resolve_RootWithSize_Throws()
    {
        var root = RootRow();
        var invalid = root with { Size = new SizeSpec { Width = 100 } };

        Assert.Throws<LayoutException>(() => Resolve(invalid));
    }

    [Fact]
    public void Resolve_LayoutLessChildWithoutAt_Throws()
    {
        var root = new ContainerElement { Children = [Rect(w: 100, h: 100)] };

        var ex = Assert.Throws<LayoutException>(() => Resolve(root));
        Assert.Contains("at", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Resolve_LayoutLessChildWithGrow_Throws()
    {
        var child = Rect(grow: 1) with { At = new PointSpec(0, 0) };
        var root = new ContainerElement { Children = [child] };

        var ex = Assert.Throws<LayoutException>(() => Resolve(root));
        Assert.Contains("grow", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Resolve_LayoutLessChildWithoutSize_Throws()
    {
        var child = new RectElement { At = new PointSpec(10, 10) };
        var root = new ContainerElement { Children = [child] };

        var ex = Assert.Throws<LayoutException>(() => Resolve(root));
        Assert.Contains("size", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Resolve_LayoutLessChildOverflows_ErrorPolicyThrows()
    {
        var child = Rect(w: 100, h: 100) with { At = new PointSpec(900, 500) };
        var root = new ContainerElement { Children = [child] };

        var ex = Assert.Throws<LayoutException>(() => Resolve(root));
        Assert.Contains("overflow", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Resolve_UnknownPaletteToken_Throws()
    {
        var child = new RectElement { Fill = new SolidFill("brand"), Size = new SizeSpec { Width = 10, Height = 10 } };
        var root = RootRow(child);

        var ex = Assert.Throws<LayoutException>(() => Resolve(root));
        Assert.Contains("brand", ex.Message);
    }

    [Fact]
    public void Resolve_ShrinkText_AppliesMeasuredScale()
    {
        var text = new TextElement { Value = "hello", Size = new SizeSpec { Width = 100, Height = 40 } };
        var result = Resolve(RootRow(text), new StubMeasurer(0.75));

        var resolved = Assert.IsType<ResolvedText>(result.Slides[0].Root.Children[0]);
        Assert.Equal(0.75, resolved.FontScale);
        Assert.Empty(result.Warnings);
    }

    [Fact]
    public void Resolve_ShrinkTextBelowMinScale_Warns()
    {
        var text = new TextElement { Value = "hello", Size = new SizeSpec { Width = 100, Height = 40 } };
        var result = Resolve(RootRow(text), new StubMeasurer(0.3));

        var resolved = Assert.IsType<ResolvedText>(result.Slides[0].Root.Children[0]);
        Assert.Equal(0.3, resolved.FontScale);
        var warning = Assert.Single(result.Warnings);
        Assert.Contains("MinScale", warning);
    }

    [Fact]
    public void Resolve_ErrorPolicyTextNotFitting_Throws()
    {
        var text = new TextElement
        {
            Value = "hello",
            Overflow = OverflowPolicy.Error,
            Size = new SizeSpec { Width = 100, Height = 40 }
        };

        Assert.Throws<LayoutException>(() => Resolve(RootRow(text), new StubMeasurer(0.9)));
    }

    [Fact]
    public void Resolve_NoMeasurer_ShrinkTextKeepsScaleOne()
    {
        var text = new TextElement { Value = "hello", Size = new SizeSpec { Width = 100, Height = 40 } };
        var result = Resolve(RootRow(text));

        var resolved = Assert.IsType<ResolvedText>(result.Slides[0].Root.Children[0]);
        Assert.Equal(1, resolved.FontScale);
        Assert.Empty(result.Warnings);
    }

    [Fact]
    public void Resolve_ShrinkText_PassesInsetReducedBoxToMeasurer()
    {
        var text = new TextElement
        {
            Value = "hello",
            Size = new SizeSpec { Width = 100, Height = 40 },
            Insets = EdgeInsets.Symmetric(5, 10)
        };
        var measurer = new StubMeasurer(1);
        Resolve(RootRow(text), measurer);

        Assert.NotNull(measurer.LastRequest);
        Assert.Equal(80, measurer.LastRequest!.BoxWidthPt);
        Assert.Equal(30, measurer.LastRequest.BoxHeightPt);
    }

    [Fact]
    public void Resolve_ClipTextNotFitting_PassesThroughWithWarning()
    {
        var text = new TextElement
        {
            Value = "hello",
            Overflow = OverflowPolicy.Clip,
            Size = new SizeSpec { Width = 100, Height = 40 }
        };
        var result = Resolve(RootRow(text), new StubMeasurer(0.9));

        var resolved = Assert.IsType<ResolvedText>(result.Slides[0].Root.Children[0]);
        Assert.Equal(1, resolved.FontScale); // clip never shrinks
        var warning = Assert.Single(result.Warnings);
        Assert.Contains("clip", warning, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("slides[0].children[0]", warning);
    }

    [Fact]
    public void Resolve_ClipTextFitting_NoWarning()
    {
        var text = new TextElement
        {
            Value = "hello",
            Overflow = OverflowPolicy.Clip,
            Size = new SizeSpec { Width = 100, Height = 40 }
        };
        var result = Resolve(RootRow(text), new StubMeasurer(1));

        Assert.Empty(result.Warnings);
    }

    private sealed class StubMeasurer : ITextMeasurer
    {
        private readonly double _scale;

        public StubMeasurer(double scale)
        {
            _scale = scale;
        }

        public TextMeasureRequest? LastRequest { get; private set; }

        public double FitScale(TextMeasureRequest request)
        {
            LastRequest = request;
            return _scale;
        }
    }
}
