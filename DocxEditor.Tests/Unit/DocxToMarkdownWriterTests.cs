using DocumentFormat.OpenXml.Packaging;
using DocxEditor.Core.Builders;
using DocxEditor.Core.Generation.Assets;
using DocxEditor.Core.Markdown;
using DocxEditor.Core.Markdown.Rendering;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Round-trip coverage for <see cref="DocxToMarkdownWriter"/>: documents built with
/// <see cref="DocumentBuilder.AddMarkdown"/> are converted back to Markdown and the emitted
/// constructs are asserted. The writer is the reverse of the rich markdown renderer, so a
/// document that renders md → docx must convert back into Markdown that contains the original
/// constructs.
/// </summary>
public sealed class DocxToMarkdownWriterTests : IDisposable
{
    private readonly string tempDirectory = Path.Combine(Path.GetTempPath(), "DocxToMarkdownWriterTests", Guid.NewGuid().ToString("N"));

    public DocxToMarkdownWriterTests()
    {
        Directory.CreateDirectory(tempDirectory);
    }

    public void Dispose()
    {
        if (Directory.Exists(tempDirectory))
        {
            Directory.Delete(tempDirectory, recursive: true);
        }
    }

    [Fact]
    public void Convert_WithHeadings_EmitsAtxHeadings()
    {
        string path = BuildDocument("# Title\n\n## Subtitle", "headings.docx");

        string markdown = ConvertToMarkdown(path);

        Assert.Contains("# Title", markdown);
        Assert.Contains("## Subtitle", markdown);
    }

    [Fact]
    public void Convert_WithParagraphFormatting_EmitsEmphasisStrikethroughAndInlineCode()
    {
        string path = BuildDocument("A **bold**, *italic*, ~~struck~~ and `code` mix.", "formatting.docx");

        string markdown = ConvertToMarkdown(path);

        Assert.Contains("**bold**", markdown);
        Assert.Contains("*italic*", markdown);
        Assert.Contains("~~struck~~", markdown);
        Assert.Contains("`code`", markdown);
    }

    [Fact]
    public void Convert_WithLists_EmitsOrderedUnorderedAndNestedLists()
    {
        string path = BuildDocument("- one\n- two\n  - nested\n\n1. first\n2. second", "lists.docx");

        string markdown = ConvertToMarkdown(path);

        Assert.Contains("- one", markdown);
        Assert.Contains("- two", markdown);
        Assert.Contains("    - nested", markdown);
        Assert.Contains("1. first", markdown);
        Assert.Contains("2. second", markdown);
    }

    [Fact]
    public void Convert_WithTable_EmitsPipeTableWithEscapedPipes()
    {
        // The escaped pipe in the source produces a literal "|" inside the cell, which the
        // writer must re-escape so the pipe table round-trips.
        string path = BuildDocument("| Col A | Col B |\n| --- | --- |\n| 1 | a\\|b |", "table.docx");

        string markdown = ConvertToMarkdown(path);

        Assert.Contains("| Col A | Col B |", markdown);
        Assert.Contains("| --- | --- |", markdown);
        Assert.Contains("| 1 | a\\|b |", markdown);
    }

    [Fact]
    public void Convert_WithBlockquote_EmitsQuoteMarker()
    {
        string path = BuildDocument("> quoted text", "quote.docx");

        string markdown = ConvertToMarkdown(path);

        Assert.Contains("> quoted text", markdown);
    }

    [Fact]
    public void Convert_WithHorizontalRule_EmitsThematicBreak()
    {
        string path = BuildDocument("---", "hr.docx");

        string markdown = ConvertToMarkdown(path);

        Assert.Contains("---", markdown);
    }

    [Fact]
    public void Convert_WithHyperlink_EmitsMarkdownLink()
    {
        string path = BuildDocument("[docs](https://example.com)", "link.docx");

        string markdown = ConvertToMarkdown(path);

        Assert.Contains("[docs](https://example.com)", markdown);
    }

    [Fact]
    public void Convert_WithLiteralSpecialCharacters_EscapesMarkdownSyntax()
    {
        string path = Path.Combine(tempDirectory, "escapes.docx");
        using (IDocumentBuilder builder = DocumentBuilder.Create(path))
        {
            // AddParagraph writes literal text (no markdown interpretation).
            builder.AddParagraph("Special: * _ ` [ ] < > # ~ ^");
            builder.AddParagraph("1. ordered-like");
            builder.AddParagraph("- dash");
            builder.Save();
        }

        string markdown = ConvertToMarkdown(path);

        // '>' is only escaped at line start (mid-paragraph it is literal Markdown); the other
        // special characters are always escaped.
        Assert.Contains(@"Special: \* \_ \` \[ \] \< > \# \~ \^", markdown);
        Assert.Contains(@"1\. ordered-like", markdown);
        Assert.Contains(@"\- dash", markdown);
    }

    [Fact]
    public void Convert_WithImage_EmbedsSelfContainedDataUri()
    {
        string source = BuildDocumentWithImage("embedded-image.docx");

        string markdown = ConvertToMarkdown(source);

        Assert.Contains("![tiny pixel](data:image/png;base64,", markdown);
    }

    [Fact]
    public void ConvertToFile_WithImage_ExtractsAssetNextToOutput()
    {
        string source = BuildDocumentWithImage("file-image.docx");
        string outputDirectory = Path.Combine(tempDirectory, "out");
        Directory.CreateDirectory(outputDirectory);
        string output = Path.Combine(outputDirectory, "converted.md");

        DocxToMarkdownResult result = DocxToMarkdownWriter.ConvertToFile(source, output);

        Assert.True(File.Exists(output));
        Assert.Contains("![tiny pixel](converted_assets/image-1.png)", result.Markdown);
        Assert.Contains("![tiny pixel](converted_assets/image-1.png)", File.ReadAllText(output));
        Assert.Contains(Path.Combine(outputDirectory, "converted_assets", "image-1.png"), result.ExtractedAssetPaths);
        Assert.True(File.Exists(Path.Combine(outputDirectory, "converted_assets", "image-1.png")));
    }

    [Fact]
    public void Convert_AfterBuildWriteAndReopen_RoundTripsConstructs()
    {
        // End-to-end: build → save to disk → reopen → convert to Markdown.
        string path = BuildDocument(
            "# Report\n\nIntro with **bold**, *italic* and `code`.\n\n- a\n- b\n\n| K | V |\n| --- | --- |\n| x | y |",
            "roundtrip.docx");

        string markdown = ConvertToMarkdown(path);

        Assert.Contains("# Report", markdown);
        Assert.Contains("**bold**", markdown);
        Assert.Contains("*italic*", markdown);
        Assert.Contains("`code`", markdown);
        Assert.Contains("- a", markdown);
        Assert.Contains("- b", markdown);
        Assert.Contains("| K | V |", markdown);
        Assert.Contains("| x | y |", markdown);
    }

    [Fact]
    public void Convert_WithEmptyBody_ProducesEmptyMarkdown()
    {
        string path = Path.Combine(tempDirectory, "empty.docx");
        using (IDocumentBuilder builder = DocumentBuilder.Create(path))
        {
            builder.Save();
        }

        string markdown = ConvertToMarkdown(path);

        Assert.Equal(string.Empty, markdown);
    }

    private string BuildDocument(string markdown, string fileName)
    {
        string path = Path.Combine(tempDirectory, fileName);
        using (IDocumentBuilder builder = DocumentBuilder.Create(path))
        {
            builder.AddMarkdown(markdown);
            builder.Save();
        }

        return path;
    }

    private string BuildDocumentWithImage(string fileName)
    {
        // 1×1 transparent PNG.
        string imagePath = Path.Combine(tempDirectory, "pixel.png");
        File.WriteAllBytes(
            imagePath,
            Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII="));

        string source = Path.Combine(tempDirectory, fileName);
        using (IDocumentBuilder builder = DocumentBuilder.Create(source))
        {
            var options = new MarkdownRenderOptions
            {
                ImageSourceOptions = new ImageSourceOptions { AllowedRoot = tempDirectory }
            };
            builder.AddRichMarkdown("![tiny pixel](pixel.png)", options);
            builder.Save();
        }

        return source;
    }

    private string ConvertToMarkdown(string path)
    {
        using WordprocessingDocument document = WordprocessingDocument.Open(path, false);
        return new DocxToMarkdownWriter(document).Convert().Markdown;
    }
}
