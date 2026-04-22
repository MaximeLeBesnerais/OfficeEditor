using DocxEditor.Core.Builders;
using DocumentFormat.OpenXml.Packaging;

namespace DocxEditor.Tests.Unit;

public class DocumentBuilderTests : IDisposable
{
    private readonly string _testFilePath = Path.Combine(Path.GetTempPath(), $"test_output_{Guid.NewGuid()}.docx");

    [Fact]
    public void Create_ShouldCreateNewDocument()
    {
        // Act
        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.AddParagraph("Hello World");
            builder.Save();
        }

        // Assert
        Assert.True(File.Exists(_testFilePath));
        using (var doc = WordprocessingDocument.Open(_testFilePath, false))
        {
            Assert.NotNull(doc.MainDocumentPart?.Document.Body);
        }
    }

    [Fact]
    public void AddParagraph_ShouldAddTextToDocument()
    {
        // Arrange
        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            // Act
            builder.AddParagraph("Test paragraph");
            builder.Save();
        }

        // Assert
        using (var doc = WordprocessingDocument.Open(_testFilePath, false))
        {
            var body = doc.MainDocumentPart!.Document.Body!;
            var paragraphs = body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();
            Assert.Contains(paragraphs, p => p.InnerText == "Test paragraph");
        }
    }

    [Fact]
    public void ReplaceText_ShouldReplaceTextInDocument()
    {
        // Arrange
        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.AddParagraph("Hello {{NAME}}");
            
            // Act
            builder.ReplaceText("{{NAME}}", "World");
            builder.Save();
        }

        // Assert
        using (var doc = WordprocessingDocument.Open(_testFilePath, false))
        {
            var body = doc.MainDocumentPart!.Document.Body!;
            var paragraphs = body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();
            Assert.Contains(paragraphs, p => p.InnerText == "Hello World");
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
