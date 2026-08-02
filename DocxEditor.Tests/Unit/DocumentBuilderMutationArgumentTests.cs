using DocxEditor.Core.Builders;
using DocxEditor.Core.Models;
using DocumentFormat.OpenXml.Packaging;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Null/empty argument contract for the public <see cref="DocumentBuilder"/> mutation APIs.
/// Every find-by-text operation funnels through the <c>FindParagraphByText</c> choke point,
/// which rejects null and empty targets: an empty target would silently match the first
/// paragraph of the document (every paragraph contains ""), deleting/replacing/inserting
/// around the wrong element. The rich-content, variable-merge, and batch-merge APIs reject
/// null lists/dictionaries and null entries so a null public argument fails predictably with
/// an argument exception instead of an NRE or a partial render that claims success while
/// dropping blocks.
/// </summary>
public class DocumentBuilderMutationArgumentTests : IDisposable
{
    private readonly List<string> _tempFiles = new();

    // IDocumentBuilder declares static abstract members, so it cannot appear as a generic
    // type argument (CS8920); helpers therefore use the concrete DocumentBuilder type.
    private static DocumentBuilder CreateTwoParagraphBuilder()
    {
        var builder = (DocumentBuilder)DocumentBuilder.Create();
        builder.AddParagraph("First");
        builder.AddParagraph("Second");
        return builder;
    }

    private static List<DocumentFormat.OpenXml.Wordprocessing.Paragraph> ReadParagraphs(DocumentBuilder builder)
    {
        using var doc = WordprocessingDocument.Open(new MemoryStream(builder.SaveToBytes()), false);
        return doc.MainDocumentPart!.Document!.Body!
            .Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().ToList();
    }

    // ---------------------------------------------------------------- empty-target choke point

    private static void AssertEmptyTargetRejected(DocumentBuilder builder, Action<DocumentBuilder> action)
    {
        var ex = Assert.Throws<ArgumentException>(() => action(builder));
        Assert.Contains("non-empty", ex.Message);

        var paragraphs = ReadParagraphs(builder);
        Assert.Equal(2, paragraphs.Count);
        Assert.Equal("First", paragraphs[0].InnerText);
        Assert.Equal("Second", paragraphs[1].InnerText);
    }

    [Fact]
    public void DeleteParagraph_WithEmptyTarget_ThrowsAndPreservesDocument()
    {
        using var builder = CreateTwoParagraphBuilder();
        AssertEmptyTargetRejected(builder, b => b.DeleteParagraph(""));
    }

    [Fact]
    public void InsertAfter_WithEmptyTarget_ThrowsAndPreservesDocument()
    {
        using var builder = CreateTwoParagraphBuilder();
        AssertEmptyTargetRejected(builder, b => b.InsertAfter("", "Inserted"));
    }

    [Fact]
    public void InsertBefore_WithEmptyTarget_ThrowsAndPreservesDocument()
    {
        using var builder = CreateTwoParagraphBuilder();
        AssertEmptyTargetRejected(builder, b => b.InsertBefore("", "Inserted"));
    }

    [Fact]
    public void ReplaceParagraph_WithEmptyTarget_ThrowsAndPreservesDocument()
    {
        using var builder = CreateTwoParagraphBuilder();
        AssertEmptyTargetRejected(builder, b => b.ReplaceParagraph("", "Replacement"));
    }

    [Fact]
    public void ReplaceWithRichContent_WithEmptyTarget_ThrowsAndPreservesDocument()
    {
        using var builder = CreateTwoParagraphBuilder();
        AssertEmptyTargetRejected(builder, b => b.ReplaceWithRichContent("", new List<ContentBlock> { new ParagraphBlock { Text = "Replacement" } }));
    }

    [Fact]
    public void ReplaceWithMarkdown_WithEmptyTarget_ThrowsAndPreservesDocument()
    {
        using var builder = CreateTwoParagraphBuilder();
        AssertEmptyTargetRejected(builder, b => b.ReplaceWithMarkdown("", "# Replacement"));
    }

    [Fact]
    public void DeleteParagraph_WithNullTarget_ThrowsArgumentNullException()
    {
        using var builder = CreateTwoParagraphBuilder();
        Assert.Throws<ArgumentNullException>(() => builder.DeleteParagraph(null!));
    }

    [Fact]
    public void ReplaceParagraph_WithNullTarget_ThrowsArgumentNullException()
    {
        using var builder = CreateTwoParagraphBuilder();
        Assert.Throws<ArgumentNullException>(() => builder.ReplaceParagraph(null!, "Replacement"));
    }

    [Fact]
    public void DeleteParagraph_WithValidTarget_StillDeletesTargetParagraph()
    {
        using var builder = CreateTwoParagraphBuilder();
        builder.DeleteParagraph("First");

        var paragraphs = ReadParagraphs(builder);
        Assert.Single(paragraphs);
        Assert.Equal("Second", paragraphs[0].InnerText);
    }

    // ---------------------------------------------------------------- rich-content blocks

    [Fact]
    public void AddRichContent_WithNullBlocks_ThrowsArgumentNullException()
    {
        using var builder = DocumentBuilder.Create();
        Assert.Throws<ArgumentNullException>(() => builder.AddRichContent(null!));
    }

    [Fact]
    public void ReplaceWithRichContent_WithNullBlocks_ThrowsArgumentNullException()
    {
        using var builder = CreateTwoParagraphBuilder();
        Assert.Throws<ArgumentNullException>(() => builder.ReplaceWithRichContent("First", null!));
    }

    [Fact]
    public void AddRichContent_WithNullBlockEntry_ThrowsAndDoesNotPartiallyRender()
    {
        using var builder = CreateTwoParagraphBuilder();
        var blocks = new List<ContentBlock> { new ParagraphBlock { Text = "ok" }, null! };

        var ex = Assert.Throws<ArgumentException>(() => builder.AddRichContent(blocks));
        Assert.Contains("blocks[1]", ex.Message);

        var paragraphs = ReadParagraphs(builder);
        Assert.Equal(2, paragraphs.Count);
        Assert.Equal("First", paragraphs[0].InnerText);
    }

    [Fact]
    public void ReplaceWithRichContent_WithNullBlockEntry_ThrowsAndDoesNotPartiallyRender()
    {
        using var builder = CreateTwoParagraphBuilder();
        var blocks = new List<ContentBlock> { new ParagraphBlock { Text = "ok" }, null! };

        var ex = Assert.Throws<ArgumentException>(() => builder.ReplaceWithRichContent("First", blocks));
        Assert.Contains("blocks[1]", ex.Message);

        var paragraphs = ReadParagraphs(builder);
        Assert.Equal(2, paragraphs.Count);
        Assert.Equal("First", paragraphs[0].InnerText);
        Assert.Equal("Second", paragraphs[1].InnerText);
    }

    [Fact]
    public void AddRichContent_WithValidBlocks_StillRendersAllBlocks()
    {
        using var builder = DocumentBuilder.Create();
        builder.AddRichContent(new List<ContentBlock>
        {
            new ParagraphBlock { Text = "hello" },
            new HeadingBlock { Level = 1, Text = "Title" }
        });

        using var doc = WordprocessingDocument.Open(new MemoryStream(builder.SaveToBytes()), false);
        var text = doc.MainDocumentPart!.Document!.Body!.InnerText;
        Assert.Contains("hello", text);
        Assert.Contains("Title", text);
    }

    // ---------------------------------------------------------------- variables & batch merge

    [Fact]
    public void MergeVariables_WithNullData_ThrowsArgumentNullException()
    {
        using var builder = DocumentBuilder.Create();
        Assert.Throws<ArgumentNullException>(() => builder.MergeVariables(null!));
    }

    [Fact]
    public void MergeBatch_WithNullRecords_ThrowsArgumentNullException()
    {
        using var builder = DocumentBuilder.Create();
        Assert.Throws<ArgumentNullException>(() => builder.MergeBatch(null!, "out_{index}.docx"));
    }

    [Fact]
    public void MergeBatch_WithNullRecordEntry_ThrowsAndWritesNoFiles()
    {
        var outputPattern = Path.Combine(Path.GetTempPath(), $"batch_{Guid.NewGuid():N}_{{index}}.docx");
        var expectedA = outputPattern.Replace("{index}", "0");
        _tempFiles.Add(expectedA);
        using var builder = DocumentBuilder.Create();

        var ex = Assert.Throws<ArgumentException>(() => builder.MergeBatch(
            new List<Dictionary<string, string>> { new() { ["name"] = "Ada" }, null! }, outputPattern));
        Assert.Contains("records[1]", ex.Message);
        Assert.False(File.Exists(expectedA));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void MergeBatch_WithNullOrEmptyOrWhitespaceOutputPattern_Throws(string? outputPattern)
    {
        using var builder = DocumentBuilder.Create();
        var records = new List<Dictionary<string, string>> { new() { ["name"] = "x" } };

        if (outputPattern is null)
        {
            Assert.Throws<ArgumentNullException>(() => builder.MergeBatch(records, outputPattern!));
        }
        else
        {
            Assert.Throws<ArgumentException>(() => builder.MergeBatch(records, outputPattern));
        }
    }

    public void Dispose()
    {
        foreach (var file in _tempFiles)
        {
            try
            {
                if (File.Exists(file))
                {
                    File.Delete(file);
                }
            }
            catch (Exception)
            {
                // best-effort temp cleanup
            }
        }
    }
}
