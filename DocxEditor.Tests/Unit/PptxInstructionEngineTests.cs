using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PptxEditor.Core.Builders;
using PptxEditor.Core.Instructions;
using PptxEditor.Core.Models;

namespace DocxEditor.Tests.Unit;

public class PptxInstructionEngineTests : IDisposable
{
    // Minimal valid PNG: 1x1 pixel, transparent.
    private static readonly byte[] PngBytes = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAACklEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==");
    private static readonly string PngBase64 = Convert.ToBase64String(PngBytes);

    private readonly string _testDir = SlideOpsTestHelpers.CreateTempDirectory(nameof(PptxInstructionEngineTests));
    private readonly PptxInstructionEngine _engine = new();

    public void Dispose() => SlideOpsTestHelpers.BestEffortDelete(_testDir);

    #region Helpers

    private string CreateDeck(params string[] titles)
        => SlideOpsTestHelpers.CreateDeckWithTitledSlides(_testDir, titles);

    private string CreateImageDeck()
    {
        var path = Path.Combine(_testDir, $"{Guid.NewGuid():N}.pptx");
        var imagePath = Path.Combine(_testDir, $"{Guid.NewGuid():N}.png");
        File.WriteAllBytes(imagePath, PngBytes);

        using (var builder = PresentationBuilder.Create(path))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddImage(imagePath);
            builder.Save();
        }

        return path;
    }

    private string CreateTableDeck()
    {
        var path = Path.Combine(_testDir, $"{Guid.NewGuid():N}.pptx");
        using (var builder = PresentationBuilder.Create(path))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTable(new List<List<string>>
            {
                new() { "h1", "h2" },
                new() { "a", "b" }
            });
            builder.Save();
        }

        return path;
    }

    /// <summary>Finds the first element id of the given anatomizer type on a 1-based slide.</summary>
    private static uint FindElementId(string path, int slide, string type)
    {
        using var builder = PresentationBuilder.Open(path);
        var anatomy = builder.Analyze();
        return anatomy[slide - 1].Elements.First(e => e.Type == type).Id;
    }

    private static PptxInstructionSet Set(params PptxInstruction[] ops)
        => new() { Operations = ops.ToList() };

    #endregion

    [Fact]
    public void Apply_ReplaceText_RoundTripsThroughOpenXml()
    {
        var path = CreateDeck("Alpha", "Beta");
        var titleId = FindElementId(path, 1, "Text");
        Guid? revision;

        using (var builder = PresentationBuilder.Open(path))
        {
            var result = _engine.Apply(builder, Set(new PptxReplaceTextInstruction
            {
                Slide = 1, ElementId = titleId, Text = "Replaced"
            }));

            Assert.True(result.Success);
            Assert.Equal(1, result.AppliedOps);
            Assert.Empty(result.FailedOps);
            Assert.Equal(new[] { 1 }, result.ChangedSlides);
            Assert.NotNull(result.Revision);
            revision = result.Revision;

            builder.Save();
        }

        // Reopen and assert via raw OpenXML.
        using var doc = PresentationDocument.Open(path, false);
        var slidePart = SlideOpsTestHelpers.GetSlidePartsInOrder(doc)[0];
        var texts = slidePart.Slide!.Descendants<Drawing.Text>().Select(t => t.Text).ToList();
        Assert.Contains("Replaced", texts);
        Assert.DoesNotContain("Alpha", texts);
    }

    [Fact]
    public void Apply_RevisionsDifferBetweenBatches()
    {
        var path = CreateDeck("Alpha");
        var titleId = FindElementId(path, 1, "Text");

        using var builder = PresentationBuilder.Open(path);
        var first = _engine.Apply(builder, Set(new PptxReplaceTextInstruction
        {
            Slide = 1, ElementId = titleId, Text = "One"
        }));
        var second = _engine.Apply(builder, Set(new PptxReplaceTextInstruction
        {
            Slide = 1, ElementId = titleId, Text = "Two"
        }));

        Assert.NotNull(first.Revision);
        Assert.NotNull(second.Revision);
        Assert.NotEqual(first.Revision, second.Revision);
    }

    [Fact]
    public void Apply_ReplaceText_PreservesRunFormatting()
    {
        var path = CreateDeck("Styled");

        // Bold the title run via raw OpenXML before editing.
        using (var doc = PresentationDocument.Open(path, true))
        {
            var slidePart = SlideOpsTestHelpers.GetSlidePartsInOrder(doc)[0];
            var run = slidePart.Slide!.Descendants<Drawing.Run>().First();
            run.RunProperties = new Drawing.RunProperties { Bold = true };
            slidePart.Slide.Save();
        }

        var titleId = FindElementId(path, 1, "Text");
        using (var builder = PresentationBuilder.Open(path))
        {
            var result = _engine.Apply(builder, Set(new PptxReplaceTextInstruction
            {
                Slide = 1, ElementId = titleId, Text = "Still Bold"
            }));
            Assert.True(result.Success);
            builder.Save();
        }

        using (var doc = PresentationDocument.Open(path, false))
        {
            var slidePart = SlideOpsTestHelpers.GetSlidePartsInOrder(doc)[0];
            var run = slidePart.Slide!.Descendants<Drawing.Run>()
                .First(r => r.Text?.Text == "Still Bold");
            Assert.NotNull(run.RunProperties);
            Assert.True(run.RunProperties.Bold?.Value);
        }
    }

    [Fact]
    public void Apply_ReplaceImage_RoundTrips()
    {
        var path = CreateImageDeck();
        var imageId = FindElementId(path, 1, "Image");

        using (var builder = PresentationBuilder.Open(path))
        {
            var result = _engine.Apply(builder, Set(new PptxReplaceImageInstruction
            {
                Slide = 1, ElementId = imageId, Image = PngBase64
            }));

            Assert.True(result.Success);
            Assert.Equal(new[] { 1 }, result.ChangedSlides);
            builder.Save();
        }

        // Deck still opens and the picture still resolves to an image part.
        using var doc = PresentationDocument.Open(path, false);
        var slidePart = SlideOpsTestHelpers.GetSlidePartsInOrder(doc)[0];
        var picture = slidePart.Slide!.Descendants<P.Picture>().First();
        var embedId = picture.BlipFill!.Blip!.Embed?.Value;
        Assert.NotNull(embedId);
        Assert.IsType<ImagePart>(slidePart.GetPartById(embedId));
    }

    [Fact]
    public void Apply_ReplaceTable_RoundTrips()
    {
        var path = CreateTableDeck();
        var tableId = FindElementId(path, 1, "Table");

        using (var builder = PresentationBuilder.Open(path))
        {
            var result = _engine.Apply(builder, Set(new PptxReplaceTableInstruction
            {
                Slide = 1,
                ElementId = tableId,
                Rows = new List<List<string>> { new() { "x", "y" }, new() { "1", "2" } }
            }));

            Assert.True(result.Success);
            Assert.Equal(new[] { 1 }, result.ChangedSlides);
            builder.Save();
        }

        using var doc = PresentationDocument.Open(path, false);
        var texts = doc.PresentationPart!.SlideParts.First().Slide!.InnerText;
        Assert.Contains("x", texts);
        Assert.DoesNotContain("h1", texts);
    }

    [Fact]
    public void Apply_MoveSlide_ReordersAndMarksWholeRange()
    {
        var path = CreateDeck("A", "B", "C");

        using (var builder = PresentationBuilder.Open(path))
        {
            var result = _engine.Apply(builder, Set(new PptxMoveSlideInstruction { From = 3, To = 1 }));

            Assert.True(result.Success);
            Assert.Equal(new[] { 1, 2, 3 }, result.ChangedSlides);
            builder.Save();
        }

        var order = SlideOpsTestHelpers.ReadSlideTextsInOrder(path);
        Assert.Equal(3, order.Count);
        Assert.StartsWith("C", order[0]);
        Assert.StartsWith("A", order[1]);
        Assert.StartsWith("B", order[2]);
    }

    [Fact]
    public void Apply_MoveSlide_Forward_MarksOnlyAffectedRange()
    {
        var path = CreateDeck("A", "B", "C", "D");

        using (var builder = PresentationBuilder.Open(path))
        {
            var result = _engine.Apply(builder, Set(new PptxMoveSlideInstruction { From = 1, To = 3 }));

            Assert.True(result.Success);
            Assert.Equal(new[] { 1, 2, 3 }, result.ChangedSlides);
            builder.Save();
        }

        var order = SlideOpsTestHelpers.ReadSlideTextsInOrder(path);
        Assert.StartsWith("B", order[0]);
        Assert.StartsWith("C", order[1]);
        Assert.StartsWith("A", order[2]);
        Assert.StartsWith("D", order[3]);
    }

    [Fact]
    public void Apply_DuplicateSlide_ProducesIdenticalCopy()
    {
        var path = CreateDeck("Original", "Other");

        using (var builder = PresentationBuilder.Open(path))
        {
            var result = _engine.Apply(builder, Set(new PptxDuplicateSlideInstruction { Slide = 1 }));

            Assert.True(result.Success);
            Assert.Equal(new[] { 2 }, result.ChangedSlides);
            Assert.Equal(3, builder.SlideCount);
            builder.Save();
        }

        var order = SlideOpsTestHelpers.ReadSlideTextsInOrder(path);
        Assert.Equal(3, order.Count);
        Assert.Equal(order[0], order[1]); // copy is verbatim
        Assert.StartsWith("Original", order[1]);
    }

    [Fact]
    public void Apply_DuplicateSlide_WithPosition_InsertsThere()
    {
        var path = CreateDeck("A", "B", "C");

        using (var builder = PresentationBuilder.Open(path))
        {
            var result = _engine.Apply(builder, Set(new PptxDuplicateSlideInstruction
            {
                Slide = 3, Position = 1
            }));

            Assert.True(result.Success);
            Assert.Equal(new[] { 1 }, result.ChangedSlides);
            builder.Save();
        }

        var order = SlideOpsTestHelpers.ReadSlideTextsInOrder(path);
        Assert.Equal(4, order.Count);
        Assert.StartsWith("C", order[0]);
        Assert.StartsWith("A", order[1]);
    }

    [Fact]
    public void Apply_DeleteSlide_RemovesAndMarksShiftedSlides()
    {
        var path = CreateDeck("A", "B", "C");

        using (var builder = PresentationBuilder.Open(path))
        {
            var result = _engine.Apply(builder, Set(new PptxDeleteSlideInstruction { Slide = 1 }));

            Assert.True(result.Success);
            Assert.Equal(new[] { 1, 2 }, result.ChangedSlides);
            builder.Save();
        }

        var order = SlideOpsTestHelpers.ReadSlideTextsInOrder(path);
        Assert.Equal(2, order.Count);
        Assert.StartsWith("B", order[0]);
        Assert.StartsWith("C", order[1]);
    }

    [Fact]
    public void Apply_DeleteLastSlide_MarksNothingButApplies()
    {
        var path = CreateDeck("A", "B");

        using (var builder = PresentationBuilder.Open(path))
        {
            var result = _engine.Apply(builder, Set(new PptxDeleteSlideInstruction { Slide = 2 }));

            Assert.True(result.Success);
            Assert.Equal(1, result.AppliedOps);
            Assert.Empty(result.ChangedSlides);
            builder.Save();
        }

        Assert.Single(SlideOpsTestHelpers.ReadSlideTextsInOrder(path));
    }

    [Fact]
    public void Apply_DeleteOnlySlide_RejectedAtValidation()
    {
        var path = CreateDeck("Only");

        using (var builder = PresentationBuilder.Open(path))
        {
            var result = _engine.Apply(builder, Set(new PptxDeleteSlideInstruction { Slide = 1 }));

            Assert.False(result.Success);
            Assert.Equal(0, result.AppliedOps);
            Assert.Null(result.Revision);
            Assert.Empty(result.ChangedSlides);
            var error = Assert.Single(result.FailedOps);
            Assert.Equal(0, error.Index);
            Assert.Equal("deleteSlide", error.Type);
            Assert.Contains("last remaining slide", error.Error);
            builder.Save();
        }

        Assert.Single(SlideOpsTestHelpers.ReadSlideTextsInOrder(path));
    }

    [Fact]
    public void Apply_TwoDeletesOnTwoSlideDeck_SecondRejected()
    {
        var path = CreateDeck("A", "B");

        using var builder = PresentationBuilder.Open(path);
        var result = _engine.Apply(builder, Set(
            new PptxDeleteSlideInstruction { Slide = 1 },
            new PptxDeleteSlideInstruction { Slide = 2 }));

        Assert.False(result.Success);
        Assert.Equal(0, result.AppliedOps);
        var error = Assert.Single(result.FailedOps);
        Assert.Equal(1, error.Index);
        Assert.Equal(2, builder.SlideCount); // nothing applied
    }

    [Fact]
    public void Apply_ValidationFailure_AppliesNothing()
    {
        var path = CreateDeck("Alpha", "Beta");
        var titleId = FindElementId(path, 1, "Text");

        using (var builder = PresentationBuilder.Open(path))
        {
            var result = _engine.Apply(builder, Set(
                new PptxReplaceTextInstruction { Slide = 1, ElementId = titleId, Text = "ShouldNotAppear" },
                new PptxReplaceTextInstruction { Slide = 99, ElementId = titleId, Text = "Nope" }));

            Assert.False(result.Success);
            Assert.Equal(0, result.AppliedOps);
            Assert.Null(result.Revision);
            Assert.Empty(result.ChangedSlides);
            var error = Assert.Single(result.FailedOps);
            Assert.Equal(1, error.Index);
            Assert.Contains("out of range", error.Error);
            builder.Save();
        }

        // First (valid) op must NOT have been applied.
        var texts = SlideOpsTestHelpers.ReadSlideTextsInOrder(path);
        Assert.StartsWith("Alpha", texts[0]);
    }

    [Fact]
    public void Apply_UnknownElement_ValidationError()
    {
        var path = CreateDeck("Alpha");

        using var builder = PresentationBuilder.Open(path);
        var result = _engine.Apply(builder, Set(new PptxReplaceTextInstruction
        {
            Slide = 1, ElementId = 9999, Text = "x"
        }));

        Assert.False(result.Success);
        Assert.Equal(0, result.AppliedOps);
        Assert.Contains("No element with id 9999", Assert.Single(result.FailedOps).Error);
    }

    [Fact]
    public void Apply_WrongElementType_ValidationError()
    {
        var path = CreateImageDeck();
        var imageId = FindElementId(path, 1, "Image");

        using var builder = PresentationBuilder.Open(path);
        var result = _engine.Apply(builder, Set(new PptxReplaceTextInstruction
        {
            Slide = 1, ElementId = imageId, Text = "x"
        }));

        Assert.False(result.Success);
        Assert.Equal(0, result.AppliedOps);
        var error = Assert.Single(result.FailedOps);
        Assert.Contains("'Image'", error.Error);
        Assert.Contains("'Text'", error.Error);
    }

    [Fact]
    public void Apply_InvalidBase64Image_ValidationError()
    {
        var path = CreateImageDeck();
        var imageId = FindElementId(path, 1, "Image");

        using var builder = PresentationBuilder.Open(path);
        var result = _engine.Apply(builder, Set(new PptxReplaceImageInstruction
        {
            Slide = 1, ElementId = imageId, Image = "!!!not-base64!!!"
        }));

        Assert.False(result.Success);
        Assert.Equal(0, result.AppliedOps);
        Assert.Contains("base64", Assert.Single(result.FailedOps).Error);
    }

    [Fact]
    public void Apply_DuplicateSlide_PositionOutOfRange_ValidationError()
    {
        var path = CreateDeck("A", "B");

        using var builder = PresentationBuilder.Open(path);
        var result = _engine.Apply(builder, Set(new PptxDuplicateSlideInstruction
        {
            Slide = 1, Position = 5
        }));

        Assert.False(result.Success);
        Assert.Equal(0, result.AppliedOps);
        Assert.Contains("out of range", Assert.Single(result.FailedOps).Error);
    }

    [Fact]
    public void Apply_ExecutionFailure_CollectedPerOp_BatchContinues()
    {
        // A shape without a text body is classified "Text" by the anatomizer but
        // the replacer fails on it at execution time → per-op error, other ops apply.
        var path = CreateDeck("Alpha", "Beta");
        uint textlessId;
        using (var doc = PresentationDocument.Open(path, true))
        {
            var slidePart = SlideOpsTestHelpers.GetSlidePartsInOrder(doc)[0];
            var shapeTree = slidePart.Slide!.CommonSlideData!.ShapeTree!;
            textlessId = 900;
            shapeTree.Append(new P.Shape(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties { Id = textlessId, Name = "Textless" },
                    new P.NonVisualShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.ShapeProperties()));
            slidePart.Slide.Save();
        }

        var titleId = FindElementId(path, 2, "Text");
        using (var builder = PresentationBuilder.Open(path))
        {
            var result = _engine.Apply(builder, Set(
                new PptxReplaceTextInstruction { Slide = 1, ElementId = textlessId, Text = "boom" },
                new PptxReplaceTextInstruction { Slide = 2, ElementId = titleId, Text = "Applied" }));

            Assert.False(result.Success);
            Assert.Equal(1, result.AppliedOps);
            var error = Assert.Single(result.FailedOps);
            Assert.Equal(0, error.Index);
            Assert.Equal("replaceText", error.Type);
            Assert.NotNull(result.Revision);
            Assert.Equal(new[] { 2 }, result.ChangedSlides);
            builder.Save();
        }

        var texts = SlideOpsTestHelpers.ReadSlideTextsInOrder(path);
        Assert.StartsWith("Applied", texts[1]);
    }

    [Fact]
    public void Apply_MixedBatch_AccumulatesChangedSlidesSorted()
    {
        var path = CreateDeck("A", "B", "C");
        var titleId = FindElementId(path, 3, "Text");

        using var builder = PresentationBuilder.Open(path);
        var result = _engine.Apply(builder, Set(
            new PptxReplaceTextInstruction { Slide = 3, ElementId = titleId, Text = "C2" },
            new PptxMoveSlideInstruction { From = 1, To = 2 }));

        Assert.True(result.Success);
        Assert.Equal(2, result.AppliedOps);
        Assert.Equal(new[] { 1, 2, 3 }, result.ChangedSlides);
    }
}
