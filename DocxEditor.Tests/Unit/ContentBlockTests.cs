using DocxEditor.Core.Builders;
using DocxEditor.Core.Content;
using DocxEditor.Core.Models;
using DocumentFormat.OpenXml.Packaging;

namespace DocxEditor.Tests.Unit;

public class ContentBlockTests : IDisposable
{
    private readonly string _testFilePath = Path.Combine(Path.GetTempPath(), $"test_blocks_{Guid.NewGuid()}.docx");

    [Fact]
    public void AddRichContent_ShouldAddMultipleBlocks()
    {
        // Arrange
        var blocks = new ContentBlockBuilder()
            .AddHeading(1, "Title")
            .AddParagraph("Paragraph 1")
            .AddList(false, new List<string> { "Item 1", "Item 2" })
            .Build();

        // Act
        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.AddRichContent(blocks);
            builder.Save();
        }

        // Assert
        using (var doc = WordprocessingDocument.Open(_testFilePath, false))
        {
            var body = doc.MainDocumentPart!.Document.Body!;
            var paragraphs = body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().ToList();
            
            Assert.True(paragraphs.Count >= 3);
            Assert.Contains(paragraphs, p => p.InnerText == "Title");
            Assert.Contains(paragraphs, p => p.InnerText == "Paragraph 1");
        }
    }

    [Fact]
    public void AddRichContent_ShouldAddTable()
    {
        // Arrange
        var blocks = new ContentBlockBuilder()
            .AddTable(new List<List<string>>
            {
                new List<string> { "Header 1", "Header 2" },
                new List<string> { "Cell 1", "Cell 2" }
            })
            .Build();

        // Act
        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.AddRichContent(blocks);
            builder.Save();
        }

        // Assert
        using (var doc = WordprocessingDocument.Open(_testFilePath, false))
        {
            var body = doc.MainDocumentPart!.Document.Body!;
            var tables = body.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>().ToList();
            
            Assert.Single(tables);
        }
    }

    [Fact]
    public void ReplaceWithRichContent_ShouldReplaceTargetParagraph()
    {
        // Arrange
        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.AddParagraph("Old content");
            builder.AddParagraph("Keep this");
            builder.Save();
        }

        var blocks = new ContentBlockBuilder()
            .AddHeading(1, "New Title")
            .AddParagraph("New paragraph")
            .Build();

        // Act
        using (var builder = DocumentBuilder.Open(_testFilePath))
        {
            builder.ReplaceWithRichContent("Old content", blocks);
            builder.Save();
        }

        // Assert
        using (var doc = WordprocessingDocument.Open(_testFilePath, false))
        {
            var body = doc.MainDocumentPart!.Document.Body!;
            var text = body.InnerText;
            
            Assert.DoesNotContain("Old content", text);
            Assert.Contains("New Title", text);
            Assert.Contains("New paragraph", text);
            Assert.Contains("Keep this", text);
        }
    }

    public void Dispose()
    {
        if (File.Exists(_testFilePath))
        {
            File.Delete(_testFilePath);
        }
    }
}
