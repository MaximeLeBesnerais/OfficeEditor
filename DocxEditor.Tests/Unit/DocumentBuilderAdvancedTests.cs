using DocxEditor.Core.Builders;
using DocumentFormat.OpenXml.Packaging;

namespace DocxEditor.Tests.Unit;

public class DocumentBuilderAdvancedTests : IDisposable
{
    private readonly string _testFilePath = Path.Combine(Path.GetTempPath(), $"test_advanced_{Guid.NewGuid()}.docx");

    [Fact]
    public void InsertAfter_ShouldInsertParagraphAfterTarget()
    {
        // Arrange
        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.AddParagraph("First paragraph");
            builder.AddParagraph("Target paragraph");
            builder.AddParagraph("Last paragraph");
            
            // Act
            builder.InsertAfter("Target", "Inserted paragraph");
            builder.Save();
        }

        // Assert
        using (var doc = WordprocessingDocument.Open(_testFilePath, false))
        {
            var body = doc.MainDocumentPart!.Document.Body!;
            var paragraphs = body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().ToList();
            
            Assert.Equal(4, paragraphs.Count);
            Assert.Equal("First paragraph", paragraphs[0].InnerText);
            Assert.Equal("Target paragraph", paragraphs[1].InnerText);
            Assert.Equal("Inserted paragraph", paragraphs[2].InnerText);
            Assert.Equal("Last paragraph", paragraphs[3].InnerText);
        }
    }

    [Fact]
    public void InsertBefore_ShouldInsertParagraphBeforeTarget()
    {
        // Arrange
        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.AddParagraph("First paragraph");
            builder.AddParagraph("Target paragraph");
            
            // Act
            builder.InsertBefore("Target", "Inserted paragraph");
            builder.Save();
        }

        // Assert
        using (var doc = WordprocessingDocument.Open(_testFilePath, false))
        {
            var body = doc.MainDocumentPart!.Document.Body!;
            var paragraphs = body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().ToList();
            
            Assert.Equal(3, paragraphs.Count);
            Assert.Equal("First paragraph", paragraphs[0].InnerText);
            Assert.Equal("Inserted paragraph", paragraphs[1].InnerText);
            Assert.Equal("Target paragraph", paragraphs[2].InnerText);
        }
    }

    [Fact]
    public void ReplaceParagraph_ShouldReplaceContentAndPreserveStyle()
    {
        // Arrange
        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.AddParagraph("Old content", "Heading1");
            
            // Act - replace without specifying style
            builder.ReplaceParagraph("Old content", "New content");
            builder.Save();
        }

        // Assert
        using (var doc = WordprocessingDocument.Open(_testFilePath, false))
        {
            var body = doc.MainDocumentPart!.Document.Body!;
            var paragraph = body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().First();
            
            Assert.Equal("New content", paragraph.InnerText);
            Assert.Equal("Heading1", paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
        }
    }

    [Fact]
    public void DeleteParagraph_ShouldRemoveParagraph()
    {
        // Arrange
        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.AddParagraph("Keep this");
            builder.AddParagraph("Delete this");
            builder.AddParagraph("Keep this too");
            
            // Act
            builder.DeleteParagraph("Delete this");
            builder.Save();
        }

        // Assert
        using (var doc = WordprocessingDocument.Open(_testFilePath, false))
        {
            var body = doc.MainDocumentPart!.Document.Body!;
            var paragraphs = body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().ToList();
            
            Assert.Equal(2, paragraphs.Count);
            Assert.DoesNotContain(paragraphs, p => p.InnerText == "Delete this");
        }
    }

    [Fact]
    public void OpenExistingDocument_ShouldPreserveContent()
    {
        // Arrange - create initial document
        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.AddParagraph("Existing content");
            builder.Save();
        }

        // Act - open and add more content
        using (var builder = DocumentBuilder.Open(_testFilePath))
        {
            builder.AddParagraph("New content");
            builder.Save();
        }

        // Assert
        using (var doc = WordprocessingDocument.Open(_testFilePath, false))
        {
            var body = doc.MainDocumentPart!.Document.Body!;
            var paragraphs = body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().ToList();
            
            Assert.Equal(2, paragraphs.Count);
            Assert.Equal("Existing content", paragraphs[0].InnerText);
            Assert.Equal("New content", paragraphs[1].InnerText);
        }
    }

    [Fact]
    public void AddParagraph_WithStyle_ShouldApplyStyle()
    {
        // Arrange & Act
        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.AddParagraph("Styled paragraph", "Heading1");
            builder.Save();
        }

        // Assert
        using (var doc = WordprocessingDocument.Open(_testFilePath, false))
        {
            var body = doc.MainDocumentPart!.Document.Body!;
            var paragraph = body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().First();
            
            Assert.Equal("Heading1", paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
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
