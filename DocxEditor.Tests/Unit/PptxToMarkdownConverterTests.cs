using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PptxEditor.Core.Builders;
using PptxEditor.Core.Converters;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Tests for the PPTX → Markdown outline writer (<see cref="PptxToMarkdownConverter"/>):
/// slide headings, title/body placeholders, nested bullets, pipe tables, free-floating
/// text paragraphs, speaker notes, image references/extraction, escaping and hidden-shape
/// skipping.
/// </summary>
public sealed class PptxToMarkdownConverterTests : IDisposable
{
    private readonly string _tempDir;

    public PptxToMarkdownConverterTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), nameof(PptxToMarkdownConverterTests), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        SlideOpsTestHelpers.BestEffortDelete(_tempDir);
    }

    [Fact]
    public void Convert_TitleAndContentSlide_ProducesStructuredOutline()
    {
        var path = BuildOutlineDeck();

        using var document = PresentationDocument.Open(path, false);
        var markdown = PptxToMarkdownConverter.Convert(document);

        // Slide 1 with a title placeholder becomes the document title (no "## Slide 1").
        Assert.Contains("# Northwind Quarterly Review", markdown);
        Assert.Contains("Prepared by the finance team", markdown);
        Assert.DoesNotContain("## Slide 1", markdown);

        // Slide 2: slide heading, then the title placeholder text as a heading.
        Assert.Contains("## Slide 2", markdown);
        Assert.Contains("### Revenue Highlights", markdown);

        // Body bullets respect lvl nesting (2 spaces per level).
        Assert.Contains("- Revenue grew 24% YoY", markdown);
        Assert.Contains("  - Driven by EMEA expansion", markdown);
        Assert.Contains("    - New enterprise segment closed", markdown);

        // Table becomes a pipe table with a header row.
        Assert.Contains("| Metric | Q1 | Q2 |", markdown);
        Assert.Contains("| --- | --- | --- |", markdown);
        Assert.Contains("| Revenue | 12.4M | 15.1M |", markdown);

        // Free-floating text box becomes a paragraph.
        Assert.Contains("Full breakdown available in the appendix.", markdown);

        // Speaker notes are a [!note] callout.
        Assert.Contains("> [!note]", markdown);
        Assert.Contains("> Remember to emphasise the EMEA growth story.", markdown);

        // Images are referenced from assets/.
        Assert.Contains("![Image", markdown);
        Assert.Contains("(assets/image_1.png)", markdown);
    }

    [Fact]
    public void ConvertToFile_WritesMarkdownAndExtractsMediaNextToOutput()
    {
        var sourcePath = BuildOutlineDeck();
        var outputDirectory = Path.Combine(_tempDir, "out");
        var outputPath = Path.Combine(outputDirectory, "deck.md");

        PptxToMarkdownConverter.ConvertToFile(sourcePath, outputPath);

        Assert.True(File.Exists(outputPath), $"Expected Markdown file at {outputPath}");
        var markdown = File.ReadAllText(outputPath);
        Assert.Contains("# Northwind Quarterly Review", markdown);
        Assert.Contains("## Slide 2", markdown);
        Assert.Contains("![Image", markdown);

        // The referenced image is extracted into assets/ next to the output.
        var extracted = Path.Combine(outputDirectory, "assets", "image_1.png");
        Assert.True(File.Exists(extracted), $"Expected extracted image at {extracted}");
        var expectedBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");
        Assert.Equal(expectedBytes, File.ReadAllBytes(extracted));
    }

    [Fact]
    public void ConvertToFile_MissingSource_ThrowsFileNotFound()
    {
        var missing = Path.Combine(_tempDir, "does-not-exist.pptx");
        Assert.Throws<FileNotFoundException>(() => PptxToMarkdownConverter.ConvertToFile(missing, Path.Combine(_tempDir, "out.md")));
    }

    [Fact]
    public void Convert_EscapesMarkdownSpecialCharacters()
    {
        var path = Path.Combine(_tempDir, $"{Guid.NewGuid():N}.pptx");
        using (var builder = PresentationBuilder.Create(path))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddText("R&D *budget*: <$500K> [link] (parens) ~tilde~");
            builder.Save();
        }

        using var document = PresentationDocument.Open(path, false);
        var markdown = PptxToMarkdownConverter.Convert(document);

        Assert.Contains("R&D \\*budget\\*: <$500K> \\[link\\] (parens) \\~tilde\\~", markdown);
    }

    [Fact]
    public void Convert_SkipsShapesHiddenInSelectionPane()
    {
        var path = Path.Combine(_tempDir, $"{Guid.NewGuid():N}.pptx");
        using (var builder = PresentationBuilder.Create(path))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("Visible Title");
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("Hidden Title");
            builder.Save();
        }

        // Hide the second slide's title via cNvPr show="0" (selection-pane hidden shape).
        using (var editDocument = PresentationDocument.Open(path, true))
        {
            var hiddenTitle = SlideOpsTestHelpers.GetSlidePartsInOrder(editDocument)[1].Slide!
                .Descendants<P.Shape>()
                .First(s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "Title");
            hiddenTitle.NonVisualShapeProperties!.NonVisualDrawingProperties!
                .SetAttribute(new OpenXmlAttribute("show", "", "0"));
            editDocument.Save();
        }

        using var document = PresentationDocument.Open(path, false);
        var markdown = PptxToMarkdownConverter.Convert(document);

        Assert.Contains("## Slide 1", markdown);
        Assert.Contains("Visible Title", markdown);
        // The slide itself still exists, but its hidden title contributes nothing.
        Assert.Contains("## Slide 2", markdown);
        Assert.DoesNotContain("Hidden Title", markdown);
    }

    /// <summary>Builds a two-slide deck: title slide + content slide with nested bullets,
    /// a table, an image, a text box and speaker notes. Placeholder types and bullet
    /// nesting are applied by hand because the fluent builder creates plain shapes.</summary>
    private string BuildOutlineDeck()
    {
        var path = Path.Combine(_tempDir, $"{Guid.NewGuid():N}.pptx");
        var pngPath = SlideOpsTestHelpers.WriteMinimalPng(_tempDir);

        using (var builder = PresentationBuilder.Create(path))
        {
            // Slide 1: title slide.
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("Northwind Quarterly Review");
            builder.CurrentSlide.AddSubtitle("Prepared by the finance team");

            // Slide 2: content slide.
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("Revenue Highlights");
            builder.CurrentSlide.AddBulletList(new[]
            {
                "Revenue grew 24% YoY",
                "Driven by EMEA expansion",
                "New enterprise segment closed"
            });
            builder.CurrentSlide.AddTable(new List<List<string>>
            {
                new() { "Metric", "Q1", "Q2" },
                new() { "Revenue", "12.4M", "15.1M" },
                new() { "Margin", "31%", "33%" }
            });
            builder.CurrentSlide.AddImage(pngPath);
            builder.CurrentSlide.AddText("Full breakdown available in the appendix.");

            MarkPlaceholder(GetShape(builder, 0, "Title"), PlaceholderValues.Title);
            MarkPlaceholder(GetShape(builder, 0, "Subtitle"), PlaceholderValues.SubTitle);
            MarkPlaceholder(GetShape(builder, 1, "Title"), PlaceholderValues.Title);

            var bulletShape = GetShape(builder, 1, "Bullet List");
            MarkPlaceholder(bulletShape, PlaceholderValues.Body);
            var bulletParagraphs = bulletShape.TextBody!.Elements<Drawing.Paragraph>().ToList();
            bulletParagraphs[1].ParagraphProperties!.Level = 1;
            bulletParagraphs[2].ParagraphProperties!.Level = 2;

            // Speaker notes on slide 2.
            var slide2Part = ((SlideBuilder)builder.GetSlide(1)).SlidePart;
            var notesPart = slide2Part.AddNewPart<NotesSlidePart>();
            notesPart.NotesSlide = BuildNotesSlide("Remember to emphasise the EMEA growth story.");

            builder.Save();
        }

        return path;
    }

    private static P.Shape GetShape(IPresentationBuilder builder, int slideIndex, string name)
    {
        var slidePart = ((SlideBuilder)builder.GetSlide(slideIndex)).SlidePart;
        return slidePart.Slide!.Descendants<P.Shape>()
            .First(s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == name);
    }

    private static void MarkPlaceholder(P.Shape shape, PlaceholderValues type)
        => shape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .Append(new PlaceholderShape { Type = type });

    private static NotesSlide BuildNotesSlide(string text)
    {
        return new NotesSlide(
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
                            new Drawing.ChildExtents { Cx = 0, Cy = 0 })),
                    new P.Shape(
                        new NonVisualShapeProperties(
                            new NonVisualDrawingProperties { Id = 1, Name = "Notes" },
                            new NonVisualShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties(
                                new PlaceholderShape { Type = PlaceholderValues.Body, Index = 1 })),
                        new ShapeProperties(),
                        new TextBody(
                            new Drawing.BodyProperties(),
                            new Drawing.ListStyle(),
                            new Drawing.Paragraph(new Drawing.Run(new Drawing.Text(text))))))));
    }
}
