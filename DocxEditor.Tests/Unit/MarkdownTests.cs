using DocxEditor.Core.Builders;
using OfficeEditor.Core.Models;
using DocumentFormat.OpenXml.Packaging;

namespace DocxEditor.Tests.Unit;

public class MarkdownTests : IDisposable
{
    private readonly string _testFilePath = Path.Combine(Path.GetTempPath(), $"test_md_{Guid.NewGuid()}.docx");

    [Fact]
    public void AddMarkdown_ShouldConvertHeadings()
    {
        // Arrange
        var markdown = "# Title\n\n## Subtitle";

        // Act
        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.AddMarkdown(markdown);
            builder.Save();
        }

        // Assert
        using (var doc = WordprocessingDocument.Open(_testFilePath, false))
        {
            var mainPart = doc.MainDocumentPart;
            Assert.NotNull(mainPart);
            var document = mainPart.Document;
            Assert.NotNull(document);
            var body = document.Body;
            Assert.NotNull(body);
            var text = body.InnerText;
            
            Assert.Contains("Title", text);
            Assert.Contains("Subtitle", text);
        }
    }

    [Fact]
    public void AddMarkdown_ShouldConvertParagraphs()
    {
        // Arrange
        var markdown = "This is a paragraph with **bold** text.";

        // Act
        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.AddMarkdown(markdown);
            builder.Save();
        }

        // Assert
        using (var doc = WordprocessingDocument.Open(_testFilePath, false))
        {
            var mainPart = doc.MainDocumentPart;
            Assert.NotNull(mainPart);
            var document = mainPart.Document;
            Assert.NotNull(document);
            var body = document.Body;
            Assert.NotNull(body);
            var text = body.InnerText;
            Assert.Contains("This is a paragraph with", text);
        }
    }

    [Fact]
    public void AddMarkdown_ShouldConvertLists()
    {
        // Arrange
        var markdown = "- Item 1\n- Item 2\n- Item 3";

        // Act
        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.AddMarkdown(markdown);
            builder.Save();
        }

        // Assert
        using (var doc = WordprocessingDocument.Open(_testFilePath, false))
        {
            var mainPart = doc.MainDocumentPart;
            Assert.NotNull(mainPart);
            var document = mainPart.Document;
            Assert.NotNull(document);
            var body = document.Body;
            Assert.NotNull(body);
            var text = body.InnerText;
            
            Assert.Contains("Item 1", text);
            Assert.Contains("Item 2", text);
            Assert.Contains("Item 3", text);
        }
    }

    [Fact]
    public void AddMarkdown_WithStyleMap_ShouldApplyCustomStyles()
    {
        // Arrange
        var markdown = "# Title";
        var styleMap = new StyleMapping
        {
            StyleMap = new Dictionary<string, string>
            {
                ["heading1"] = "CustomTitle"
            }
        };

        // Act
        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.AddMarkdown(markdown, styleMap);
            builder.Save();
        }

        // Assert
        using (var doc = WordprocessingDocument.Open(_testFilePath, false))
        {
            var mainPart = doc.MainDocumentPart;
            Assert.NotNull(mainPart);
            var document = mainPart.Document;
            Assert.NotNull(document);
            var body = document.Body;
            Assert.NotNull(body);
            var paragraph = body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().First();
            
            Assert.Equal("CustomTitle", paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
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
