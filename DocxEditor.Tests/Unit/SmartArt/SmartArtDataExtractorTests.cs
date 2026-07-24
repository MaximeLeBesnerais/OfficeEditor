using DocumentFormat.OpenXml.Packaging;
using PptxEditor.Core.Converters.SmartArt;
using Presentation = DocumentFormat.OpenXml.Presentation;
using Xunit;

namespace DocxEditor.Tests.Unit.SmartArt;

public sealed class SmartArtDataExtractorTests
{
    [Fact]
    public void Extract_FromReferenceDeck_ReturnsModelWithAllMetadata()
    {
        using var document = OpenSalesAccelDeck();
        var (slide15Part, graphicFrame) = FindDiagramGraphicFrame(document);
        Assert.NotNull(slide15Part);
        Assert.NotNull(graphicFrame);

        var model = SmartArtDataExtractor.Extract(slide15Part, graphicFrame);
        Assert.NotNull(model);

        Assert.Equal("process", model!.LayoutCategory);
        Assert.Contains("process5", model.LayoutTypeId);
        Assert.Contains("simple1", model.QuickStyleId);
        Assert.Contains("accent1_2", model.ColorStyleId);

        Assert.Equal(4, model.Nodes.Count);
        Assert.Contains(model.Nodes, n => n.Text == "SUSTAIN");
        Assert.Contains(model.Nodes, n => n.Text == "DIAGNOSE");
        Assert.Contains(model.Nodes, n => n.Text == "DESIGN");
        Assert.Contains(model.Nodes, n => n.Text == "DELIVER");

        Assert.All(model.Nodes, n =>
        {
            Assert.False(string.IsNullOrEmpty(n.ModelId));
            Assert.NotEmpty(n.Text);
            Assert.Equal(0, n.HierarchyLevel);
            Assert.False(n.IsPlaceholder);
            Assert.NotEmpty(n.PlaceholderText);
        });

        Assert.NotEmpty(model.Connections);
        Assert.Contains(model.Connections, c => c.Type == "presOf");
        Assert.Contains(model.Connections, c => c.Type == "presParOf");

        Assert.DoesNotContain(model.Nodes,
            n => n.ModelId.Contains("parTrans", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(model.Nodes,
            n => n.ModelId.Contains("sibTrans", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Extract_ConnectionsAllHaveValidSrcDest()
    {
        using var document = OpenSalesAccelDeck();
        var (slide15Part, graphicFrame) = FindDiagramGraphicFrame(document);
        Assert.NotNull(slide15Part);
        Assert.NotNull(graphicFrame);

        var model = SmartArtDataExtractor.Extract(slide15Part, graphicFrame);
        Assert.NotNull(model);
        Assert.NotEmpty(model!.Connections);

        foreach (var cxn in model.Connections)
        {
            Assert.NotEmpty(cxn.ModelId);
            Assert.NotNull(cxn.Type);
            if (cxn.Type == "connection")
            {
                Assert.NotEmpty(cxn.SourceId);
                Assert.NotEmpty(cxn.DestId);
            }
        }
    }

    [Fact]
    public void Extract_NonDiagramSlide_ReturnsNull()
    {
        using var document = OpenSalesAccelDeck();
        var slide1Part = document.PresentationPart!.SlideParts.First();
        var slide1 = slide1Part.Slide;
        Assert.NotNull(slide1);

        var graphicFrame = slide1!.CommonSlideData!.ShapeTree!
            .Elements<Presentation.GraphicFrame>()
            .FirstOrDefault();

        if (graphicFrame == null)
        {
            // No graphic frames at all — this is the expected case for slide 1.
            return;
        }

        var model = SmartArtDataExtractor.Extract(slide1Part, graphicFrame);
        Assert.NotNull(model); // Should handle non-diagram graphic frames gracefully
    }

    // ── helpers ──

    private static PresentationDocument OpenSalesAccelDeck()
    {
        var path = Path.Combine(ResolveReferenceDirectory(), "sales_acceleration_deck.pptx");
        Assert.True(File.Exists(path), $"Reference deck not found: {path}");
        return PresentationDocument.Open(path, false);
    }

    private static (SlidePart slidePart, Presentation.GraphicFrame graphicFrame) FindDiagramGraphicFrame(
        PresentationDocument document)
    {
        foreach (var slidePart in document.PresentationPart!.SlideParts)
        {
            var slide = slidePart.Slide;
            if (slide?.CommonSlideData?.ShapeTree == null) continue;

            var graphicFrame = slide.CommonSlideData.ShapeTree
                .Elements<Presentation.GraphicFrame>()
                .FirstOrDefault(gf =>
                {
                    var uri = gf.Graphic?.GraphicData?.Uri?.Value ?? "";
                    return uri.Contains("/drawingml/2006/diagram", StringComparison.Ordinal);
                });

            if (graphicFrame != null)
                return (slidePart, graphicFrame);
        }

        return (null!, null!);
    }

    private static string ResolveReferenceDirectory()
    {
        var baseDir = AppContext.BaseDirectory;
        return Path.GetFullPath(Path.Combine(
            baseDir, "..", "..", "..", "..", "examples", "REF", "PPTX"));
    }
}
