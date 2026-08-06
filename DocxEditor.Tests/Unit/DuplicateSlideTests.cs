using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using PptxEditor.Core.Builders;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Tests for <see cref="PresentationBuilder.DuplicateSlide"/> — verbatim slide copy with
/// identical relationship ids, position control, and package integrity after reopen.
/// </summary>
public sealed class DuplicateSlideTests : IDisposable
{
    private readonly string _tempDirectory;

    public DuplicateSlideTests()
    {
        _tempDirectory = SlideOpsTestHelpers.CreateTempDirectory(nameof(DuplicateSlideTests));
    }

    public void Dispose()
    {
        SlideOpsTestHelpers.BestEffortDelete(_tempDirectory);
    }

    [Fact]
    public void DuplicateSlide_DefaultPosition_InsertsCopyImmediatelyAfterSource()
    {
        var path = SlideOpsTestHelpers.CreateDeckWithTitledSlides(_tempDirectory, "A", "B");

        using (var builder = PresentationBuilder.Open(path))
        {
            builder.DuplicateSlide(0);
            Assert.Equal(3, builder.SlideCount);
            builder.Save();
        }

        Assert.Equal(new[] { "A", "A", "B" }, SlideOpsTestHelpers.ReadSlideTextsInOrder(path));

        // Slide ids must stay unique.
        using var doc = PresentationDocument.Open(path, false);
        var ids = doc.PresentationPart!.Presentation!.SlideIdList!.ChildElements
            .OfType<SlideId>()
            .Select(sid => sid.Id!.Value)
            .ToList();
        Assert.Equal(ids.Count, ids.Distinct().Count());
    }

    [Theory]
    [InlineData(2, 0, new[] { "C", "A", "B", "C" })] // copy of last slide to the front
    [InlineData(0, 3, new[] { "A", "B", "C", "A" })] // position == count appends at the end
    [InlineData(1, 2, new[] { "A", "B", "B", "C" })] // same as default position for index 1
    public void DuplicateSlide_ExplicitPosition_InsertsAtRequestedIndex(
        int index, int position, string[] expectedOrder)
    {
        var path = SlideOpsTestHelpers.CreateDeckWithTitledSlides(_tempDirectory, "A", "B", "C");

        using (var builder = PresentationBuilder.Open(path))
        {
            builder.DuplicateSlide(index, position);
            Assert.Equal(4, builder.SlideCount);
            builder.Save();
        }

        Assert.Equal(expectedOrder, SlideOpsTestHelpers.ReadSlideTextsInOrder(path));
    }

    [Fact]
    public void DuplicateSlide_InvalidArguments_Throw()
    {
        var path = SlideOpsTestHelpers.CreateDeckWithTitledSlides(_tempDirectory, "A", "B");

        using var builder = PresentationBuilder.Open(path);
        Assert.Throws<ArgumentOutOfRangeException>(() => builder.DuplicateSlide(-1));
        Assert.Throws<ArgumentOutOfRangeException>(() => builder.DuplicateSlide(2));
        Assert.Throws<ArgumentOutOfRangeException>(() => builder.DuplicateSlide(0, -1));
        Assert.Throws<ArgumentOutOfRangeException>(() => builder.DuplicateSlide(0, 3));
    }

    [Fact]
    public void DuplicateSlide_WithImage_SharesImagePartUnderIdenticalRelationshipId()
    {
        var imagePath = SlideOpsTestHelpers.WriteMinimalPng(_tempDirectory);
        var path = Path.Combine(_tempDirectory, "with-image.pptx");

        using (var builder = PresentationBuilder.Create(path))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("Source");
            builder.CurrentSlide.AddImage(imagePath);
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("Other");
            builder.Save();
        }

        using (var builder = PresentationBuilder.Open(path))
        {
            builder.DuplicateSlide(0);
            builder.Save();
        }

        using var doc = PresentationDocument.Open(path, false);
        var slideParts = SlideOpsTestHelpers.GetSlidePartsInOrder(doc);
        Assert.Equal(3, slideParts.Count);

        // Content equality: the duplicated slide carries the same text as the source.
        Assert.Equal(slideParts[0].Slide!.InnerText, slideParts[1].Slide!.InnerText);

        var (sourceRelId, sourceImage) = GetSingleImageReference(slideParts[0]);
        var (dupRelId, dupImage) = GetSingleImageReference(slideParts[1]);

        // Identical relationship id (critical trick) and SHARED part, not a copy.
        Assert.Equal(sourceRelId, dupRelId);
        Assert.Equal(sourceImage.Uri, dupImage.Uri);
    }

    [Fact]
    public void DuplicateSlide_ReopenedDeckHasNoDanglingRelationshipsAndValidates()
    {
        var imagePath = SlideOpsTestHelpers.WriteMinimalPng(_tempDirectory);
        var baselinePath = Path.Combine(_tempDirectory, "baseline.pptx");
        var duplicatedPath = Path.Combine(_tempDirectory, "duplicated.pptx");

        using (var builder = PresentationBuilder.Create(baselinePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("Source");
            builder.CurrentSlide.AddImage(imagePath);
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("Other");
            builder.Save();
        }

        File.Copy(baselinePath, duplicatedPath);
        using (var builder = PresentationBuilder.Open(duplicatedPath))
        {
            builder.DuplicateSlide(0);
            builder.Save();
        }

        using var doc = PresentationDocument.Open(duplicatedPath, false);

        // No dangling relationships: every r:embed / r:id / r:link in the duplicated slide
        // XML resolves to a part or a hyperlink/external relationship (regex on OuterXml,
        // per the attribute-handling rule).
        var duplicatedSlidePart = SlideOpsTestHelpers.GetSlidePartsInOrder(doc)[1];
        var relationshipIds = Regex.Matches(
                duplicatedSlidePart.Slide!.OuterXml,
                @"\br:(?:embed|id|link)\s*=\s*""([^""]*)""")
            .Select(match => match.Groups[1].Value)
            .Distinct()
            .ToList();

        Assert.NotEmpty(relationshipIds);
        foreach (var relId in relationshipIds)
        {
            Assert.True(RelationshipExists(duplicatedSlidePart, relId),
                $"Relationship '{relId}' referenced by the duplicated slide does not exist.");
        }

        // OpenXML validation: duplicating must not introduce ANY new validation errors.
        var baselineErrors = CountValidationErrors(baselinePath);
        var duplicatedErrors = CountValidationErrors(duplicatedPath);
        Assert.Equal(baselineErrors, duplicatedErrors);
    }

    [Fact]
    public void DuplicateSlide_DoesNotCopyNotes()
    {
        // Notes are per-slide state (documented decision): a duplicate starts without a
        // notes part even when the source slide carries one.
        var path = Path.Combine(_tempDirectory, "with-notes.pptx");

        using (var builder = PresentationBuilder.Create(path))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("Source");
            builder.CurrentSlide.SetNotes("Source-only speaker notes.");
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("Other");
            builder.Save();
        }

        using (var builder = PresentationBuilder.Open(path))
        {
            builder.DuplicateSlide(0);
            builder.Save();
        }

        using var doc = PresentationDocument.Open(path, false);
        var slideParts = SlideOpsTestHelpers.GetSlidePartsInOrder(doc);
        Assert.Equal(3, slideParts.Count);

        Assert.NotNull(slideParts[0].NotesSlidePart);
        Assert.Contains("Source-only speaker notes.", slideParts[0].NotesSlidePart!.NotesSlide!.CommonSlideData!.ShapeTree!.InnerText);
        Assert.Null(slideParts[1].NotesSlidePart); // the duplicate starts without notes
        Assert.Null(slideParts[2].NotesSlidePart);
    }

    private static (string RelId, ImagePart Part) GetSingleImageReference(SlidePart slidePart)
    {
        var pair = slidePart.Parts.Single(p => p.OpenXmlPart is ImagePart);
        return (pair.RelationshipId, (ImagePart)pair.OpenXmlPart);
    }

    private static bool RelationshipExists(SlidePart slidePart, string relId)
    {
        if (slidePart.HyperlinkRelationships.Any(rel => rel.Id == relId) ||
            slidePart.ExternalRelationships.Any(rel => rel.Id == relId))
        {
            return true;
        }

        try
        {
            slidePart.GetPartById(relId);
            return true;
        }
        catch (ArgumentOutOfRangeException)
        {
            return false;
        }
    }

    private static int CountValidationErrors(string path)
    {
        using var document = PresentationDocument.Open(path, false);
        var validator = new OpenXmlValidator();
        return validator.Validate(document).Count();
    }
}
