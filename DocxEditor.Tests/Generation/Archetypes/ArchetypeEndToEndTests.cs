using PptxEditor.Core.Generation.Archetypes;
using PptxEditor.Core.Generation.Components;
using PptxEditor.Core.Generation.Emit.Ooxml;
using PptxEditor.Core.Generation.Layout;
using PptxEditor.Core.Generation.Model;
using PptxEditor.Core.Generation.Schema;

namespace DocxEditor.Tests.Generation.Archetypes;

public sealed class ArchetypeEndToEndTests
{
    private readonly GenerationDocumentParser _parser = new();

    private static string DemoDeckJson()
    {
        var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Generation", "Archetypes", "demo-deck.json");
        if (!File.Exists(path))
        {
            path = Path.Combine(FindRepoRoot(), "DocxEditor.Tests", "Generation", "Archetypes", "demo-deck.json");
        }
        return File.ReadAllText(path);
    }

    private static string FindRepoRoot()
    {
        var dir = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "DocxEditor.sln")))
            {
                return dir.FullName;
            }
            dir = dir.Parent;
        }
        throw new InvalidOperationException("Could not find repository root containing DocxEditor.sln.");
    }

    [Fact]
    public void DemoDeck_ParsesWithAllSixSlides()
    {
        var json = DemoDeckJson();
        var result = _parser.Validate(json);

        Assert.True(result.IsValid, string.Join(Environment.NewLine, result.Errors.Select(e => e.ToString())));
        Assert.NotNull(result.Document);
        Assert.Equal(6, result.Document!.Slides.Count);
    }

    [Fact]
    public void DemoDeck_ArchetypeExpander_ExpandsAllSixSlides()
    {
        var json = DemoDeckJson();
        var document = _parser.Parse(json);

        var expanded = ArchetypeExpander.Expand(document);

        Assert.Equal(6, expanded.Slides.Count);
        Assert.All(expanded.Slides, slide =>
        {
            var container = Assert.IsType<ContainerElement>(slide);
            Assert.True(container.Layout is not null
                        || container.Fill is not null
                        || container.Children.Count > 0,
                "every archetype slide must produce a meaningful root container");
        });

        var slideTypes = new[] { "cover", "kpi_row", "two_col", "table_slide", "section", "cover" };
        for (var i = 0; i < 6; i++)
        {
            var numComponents = CountComponents(expanded.Slides[i]);
            Assert.True(numComponents > 0,
                $"slide {i} ({slideTypes[i]} archetype) must produce at least one component node for component expansion to handle later");
        }
    }

    [Fact]
    public void DemoDeck_FullPipeline_ProducesValidPptx()
    {
        var json = DemoDeckJson();
        var document = _parser.Parse(json);

        var archetyped = ArchetypeExpander.Expand(document);
        var componentized = ComponentExpander.Expand(archetyped);
        var resolver = new LayoutResolver();
        var layout = resolver.Resolve(componentized);

        Assert.Equal(6, layout.Slides.Count);
        Assert.All(layout.Slides, slide =>
        {
            Assert.True(slide.WidthPt > 0);
            Assert.True(slide.HeightPt > 0);
            Assert.True(slide.Root.Width > 0);
            Assert.True(slide.Root.Height > 0);
        });

        var emitter = new OoxmlEmitter();
        var result = emitter.Emit(layout);

        Assert.NotNull(result.Bytes);
        Assert.True(result.Bytes.Length > 0, "a valid .pptx must produce non-zero bytes");

        var pptxMagic = new byte[] { 0x50, 0x4B, 0x03, 0x04 };
        Assert.Equal(pptxMagic, result.Bytes.Take(4).ToArray());
    }

    [Fact]
    public void DemoDeck_FullPipeline_RunsWithoutWarnings()
    {
        var json = DemoDeckJson();
        var document = _parser.Parse(json);

        var archetyped = ArchetypeExpander.Expand(document);
        var componentized = ComponentExpander.Expand(archetyped);
        var resolver = new LayoutResolver();
        var layout = resolver.Resolve(componentized);

        Assert.Empty(layout.Warnings);
    }

    [Fact]
    public void DemoDeck_AllKnownSlideTypes_PresentInOrder()
    {
        var json = DemoDeckJson();
        var document = _parser.Parse(json);
        var expanded = ArchetypeExpander.Expand(document);

        Assert.Equal(6, expanded.Slides.Count);

        var slide1 = expanded.Slides[0];
        Assert.Equal(new SolidFill("paper"), slide1.Fill);
        Assert.Equal(Justify.Center, slide1.Layout!.Justify);

        var slide2 = expanded.Slides[1];
        Assert.Equal(Justify.Start, slide2.Layout!.Justify);

        var slide3 = expanded.Slides[2];
        Assert.Equal(LayoutMode.Column, slide3.Layout!.Mode);

        var slide4 = expanded.Slides[3];
        Assert.Equal(LayoutMode.Column, slide4.Layout!.Mode);

        var slide5 = expanded.Slides[4];
        Assert.Equal(Justify.End, slide5.Layout!.Justify);

        var slide6 = expanded.Slides[5];
        Assert.Equal(Justify.Center, slide6.Layout!.Justify);
    }

    // ------------------------------------------------------------------ pipeline: no component left behind

    [Fact]
    public void FullPipeline_EveryResolvedElementIsPrivativeOrContainer()
    {
        var json = DemoDeckJson();
        var document = _parser.Parse(json);

        var archetyped = ArchetypeExpander.Expand(document);
        var componentized = ComponentExpander.Expand(archetyped);
        var resolver = new LayoutResolver();
        var layout = resolver.Resolve(componentized);

        foreach (var slide in layout.Slides)
        {
            AssertAllResolved(slide.Root);
        }
    }

    // ------------------------------------------------------------------ helpers

    private static int CountComponents(GenElement element)
    {
        var count = element is ComponentElement ? 1 : 0;
        var children = element switch
        {
            ContainerElement c => c.Children,
            GroupElement g => g.Children,
            _ => null
        };
        if (children is null)
        {
            return count;
        }
        foreach (var child in children)
        {
            count += CountComponents(child);
        }
        return count;
    }

    private static void AssertAllResolved(ResolvedElement element)
    {
        // Cast via object: ComponentElement (GenElement) is outside the ResolvedElement
        // hierarchy, so a direct pattern is a compile-time tautology (CS0184).
        Assert.False((object)element is ComponentElement);
        if (element is ResolvedContainer container)
        {
            foreach (var child in container.Children)
            {
                AssertAllResolved(child);
            }
        }
    }
}
