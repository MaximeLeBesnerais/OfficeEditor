using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using PptxEditor.Core.Builders;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Tests for orphaned-part pruning in <see cref="PresentationBuilder.RemoveSlide"/>:
/// slide-only media parts must be removed from the package, shared parts must survive.
/// </summary>
public sealed class RemoveSlidePruningTests : IDisposable
{
    private readonly string _tempDirectory;

    public RemoveSlidePruningTests()
    {
        _tempDirectory = SlideOpsTestHelpers.CreateTempDirectory(nameof(RemoveSlidePruningTests));
    }

    public void Dispose()
    {
        SlideOpsTestHelpers.BestEffortDelete(_tempDirectory);
    }

    [Fact]
    public void RemoveSlide_PrunesOrphanedImagePart()
    {
        var imagePath = SlideOpsTestHelpers.WriteMinimalPng(_tempDirectory);
        var path = Path.Combine(_tempDirectory, "orphan-image.pptx");

        using (var builder = PresentationBuilder.Create(path))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("A");
            builder.CurrentSlide.AddImage(imagePath);
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("B");
            builder.Save();
        }

        int partCountBefore;
        using (var doc = PresentationDocument.Open(path, false))
        {
            var allParts = SlideOpsTestHelpers.EnumerateAllParts(doc);
            Assert.Single(allParts.OfType<ImagePart>());
            partCountBefore = allParts.Count;
        }

        using (var builder = PresentationBuilder.Open(path))
        {
            builder.RemoveSlide(0);
            Assert.Equal(1, builder.SlideCount);
            builder.Save();
        }

        using (var doc = PresentationDocument.Open(path, false))
        {
            var allParts = SlideOpsTestHelpers.EnumerateAllParts(doc);

            // The image lived only on the removed slide: it must be gone from the package.
            Assert.Empty(allParts.OfType<ImagePart>());
            Assert.True(allParts.Count < partCountBefore,
                $"Expected fewer parts after pruning (before: {partCountBefore}, after: {allParts.Count}).");
            Assert.Single(allParts.OfType<SlidePart>());
            Assert.Equal(new[] { "B" }, SlideOpsTestHelpers.ReadSlideTextsInOrder(doc));
        }
    }

    [Fact]
    public void RemoveSlide_KeepsImagePartSharedWithRemainingSlide()
    {
        var imagePath = SlideOpsTestHelpers.WriteMinimalPng(_tempDirectory);
        var path = Path.Combine(_tempDirectory, "shared-image.pptx");

        using (var builder = PresentationBuilder.Create(path))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("A");
            builder.CurrentSlide.AddImage(imagePath);
            builder.Save();
        }

        // Duplicate slide 0 so the same ImagePart is referenced by two slides.
        using (var builder = PresentationBuilder.Open(path))
        {
            builder.DuplicateSlide(0);
            builder.RemoveSlide(0);
            builder.Save();
        }

        using var doc = PresentationDocument.Open(path, false);
        var imagePart = Assert.Single(SlideOpsTestHelpers.EnumerateAllParts(doc).OfType<ImagePart>());

        // The remaining slide's image reference still resolves to the shared part.
        var remainingSlide = SlideOpsTestHelpers.GetSlidePartsInOrder(doc).Single();
        var relId = remainingSlide.Parts.Single(p => p.OpenXmlPart is ImagePart).RelationshipId;
        var resolved = (ImagePart)remainingSlide.GetPartById(relId);
        Assert.Equal(imagePart.Uri, resolved.Uri);
    }

    [Fact]
    public void RemoveSlide_PrunesOrphanedNotesSlidePart()
    {
        var path = Path.Combine(_tempDirectory, "orphan-notes.pptx");

        using (var builder = PresentationBuilder.Create(path))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("A");
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("B");
            builder.Save();
        }

        // Attach a notes part to slide 0 (the builder has no notes API).
        using (var doc = PresentationDocument.Open(path, true))
        {
            var slidePart = SlideOpsTestHelpers.GetSlidePartsInOrder(doc)[0];
            var notesPart = slidePart.AddNewPart<NotesSlidePart>();
            notesPart.NotesSlide = new NotesSlide(
                new CommonSlideData(
                    new ShapeTree(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties { Id = 0, Name = "" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new GroupShapeProperties(
                            new Drawing.TransformGroup(
                                new Drawing.Offset { X = 0, Y = 0 },
                                new Drawing.Extents { Cx = 0, Cy = 0 },
                                new Drawing.ChildOffset { X = 0, Y = 0 },
                                new Drawing.ChildExtents { Cx = 0, Cy = 0 })))));
            doc.Save();
        }

        using (var doc = PresentationDocument.Open(path, false))
        {
            Assert.Single(SlideOpsTestHelpers.EnumerateAllParts(doc).OfType<NotesSlidePart>());
        }

        using (var builder = PresentationBuilder.Open(path))
        {
            builder.RemoveSlide(0);
            builder.Save();
        }

        using (var doc = PresentationDocument.Open(path, false))
        {
            Assert.Empty(SlideOpsTestHelpers.EnumerateAllParts(doc).OfType<NotesSlidePart>());
            Assert.Equal(new[] { "B" }, SlideOpsTestHelpers.ReadSlideTextsInOrder(doc));
        }
    }
}
