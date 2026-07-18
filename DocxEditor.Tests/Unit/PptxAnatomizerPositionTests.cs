using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PptxEditor.Core.Builders;

namespace DocxEditor.Tests.Unit;

public class PptxAnatomizerPositionTests : IDisposable
{
    private readonly string _testDir = SlideOpsTestHelpers.CreateTempDirectory(nameof(PptxAnatomizerPositionTests));

    public void Dispose() => SlideOpsTestHelpers.BestEffortDelete(_testDir);

    [Fact]
    public void Shape_WithExplicitXfrm_ReportsPosition()
    {
        // AddText creates a shape at (0, 1440000) with extents (7200000, 3600000).
        var path = SlideOpsTestHelpers.CreateDeckWithTitledSlides(_testDir, "Title");
        using (var builder = PresentationBuilder.Open(path))
        {
            builder.GetSlide(0);
            builder.CurrentSlide.AddText("Body text");
            builder.Save();
        }

        using var reader = PresentationBuilder.Open(path);
        var anatomy = reader.Analyze();
        var body = anatomy[0].Elements.First(e => e.Text == "Body text");

        Assert.Equal(0L, body.X);
        Assert.Equal(1440000L, body.Y);
        Assert.Equal(7200000L, body.Cx);
        Assert.Equal(3600000L, body.Cy);
    }

    [Fact]
    public void GraphicFrame_ReportsTransformPosition()
    {
        var path = Path.Combine(_testDir, $"{Guid.NewGuid():N}.pptx");
        using (var builder = PresentationBuilder.Create(path))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTable(new List<List<string>> { new() { "a" } });
            builder.Save();
        }

        using var reader = PresentationBuilder.Open(path);
        var anatomy = reader.Analyze();
        var table = anatomy[0].Elements.First(e => e.Type == "Table");

        // SlideBuilder.AddTable: p:xfrm off=(0,1440000) ext=(7200000,3600000).
        Assert.Equal(0L, table.X);
        Assert.Equal(1440000L, table.Y);
        Assert.Equal(7200000L, table.Cx);
        Assert.Equal(3600000L, table.Cy);
    }

    [Fact]
    public void Shape_WithoutXfrm_ReportsNullPosition()
    {
        var path = SlideOpsTestHelpers.CreateDeckWithTitledSlides(_testDir, "Title");

        // Append a shape with no a:xfrm (placeholder-style: inherits layout position).
        using (var doc = PresentationDocument.Open(path, true))
        {
            var slidePart = SlideOpsTestHelpers.GetSlidePartsInOrder(doc)[0];
            var shapeTree = slidePart.Slide!.CommonSlideData!.ShapeTree!;
            shapeTree.Append(new P.Shape(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 77, Name = "NoXfrm" },
                    new P.NonVisualShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.ShapeProperties(),
                new P.TextBody(
                    new Drawing.BodyProperties(),
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph(new Drawing.Run(new Drawing.Text("Inherited"))))));
            slidePart.Slide.Save();
        }

        using var reader = PresentationBuilder.Open(path);
        var anatomy = reader.Analyze();
        var shape = anatomy[0].Elements.First(e => e.Id == 77);

        Assert.Equal("Inherited", shape.Text);
        Assert.Null(shape.X);
        Assert.Null(shape.Y);
        Assert.Null(shape.Cx);
        Assert.Null(shape.Cy);
    }

    [Fact]
    public void GroupShape_ReportsGroupTransform_ButIsNotRecursed()
    {
        var path = SlideOpsTestHelpers.CreateDeckWithTitledSlides(_testDir, "Title");

        using (var doc = PresentationDocument.Open(path, true))
        {
            var slidePart = SlideOpsTestHelpers.GetSlidePartsInOrder(doc)[0];
            var shapeTree = slidePart.Slide!.CommonSlideData!.ShapeTree!;
            shapeTree.Append(new P.GroupShape(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 88, Name = "Grp" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(
                    new Drawing.TransformGroup(
                        new Drawing.Offset { X = 100, Y = 200 },
                        new Drawing.Extents { Cx = 3000, Cy = 4000 },
                        new Drawing.ChildOffset { X = 0, Y = 0 },
                        new Drawing.ChildExtents { Cx = 3000, Cy = 4000 })),
                new P.Shape(
                    new P.NonVisualShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 89, Name = "Child" },
                        new P.NonVisualShapeDrawingProperties(),
                        new P.ApplicationNonVisualDrawingProperties()),
                    new P.ShapeProperties(),
                    new P.TextBody(
                        new Drawing.BodyProperties(),
                        new Drawing.ListStyle(),
                        new Drawing.Paragraph(new Drawing.Run(new Drawing.Text("ChildText")))))));
            slidePart.Slide.Save();
        }

        using var reader = PresentationBuilder.Open(path);
        var anatomy = reader.Analyze();
        var group = anatomy[0].Elements.First(e => e.Type == "Group");

        Assert.Equal(88u, group.Id);
        Assert.Equal(100L, group.X);
        Assert.Equal(200L, group.Y);
        Assert.Equal(3000L, group.Cx);
        Assert.Equal(4000L, group.Cy);

        // Groups are not recursed: the child shape is not reported.
        Assert.DoesNotContain(anatomy[0].Elements, e => e.Id == 89);
    }

    [Fact]
    public void Picture_ReportsPosition()
    {
        var imagePath = Path.Combine(_testDir, $"{Guid.NewGuid():N}.png");
        File.WriteAllBytes(imagePath, Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAACklEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg=="));

        var path = Path.Combine(_testDir, $"{Guid.NewGuid():N}.pptx");
        using (var builder = PresentationBuilder.Create(path))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddImage(imagePath);
            builder.Save();
        }

        using var reader = PresentationBuilder.Open(path);
        var anatomy = reader.Analyze();
        var image = anatomy[0].Elements.First(e => e.Type == "Image");

        // SlideBuilder.AddImage: a:xfrm off=(0,1440000) ext=(7200000,3600000).
        Assert.Equal(0L, image.X);
        Assert.Equal(1440000L, image.Y);
        Assert.Equal(7200000L, image.Cx);
        Assert.Equal(3600000L, image.Cy);
    }
}
