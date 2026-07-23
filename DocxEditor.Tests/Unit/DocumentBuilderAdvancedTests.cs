using DocxEditor.Core.Builders;
using DocumentFormat.OpenXml.Packaging;

namespace DocxEditor.Tests.Unit;

public class DocumentBuilderAdvancedTests : IDisposable
{
    private readonly string _testFilePath = Path.Combine(Path.GetTempPath(), $"test_advanced_{Guid.NewGuid()}.docx");
    private readonly List<string> _additionalFiles = new();

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
            var mainPart = doc.MainDocumentPart;
            Assert.NotNull(mainPart);
            var document = mainPart.Document;
            Assert.NotNull(document);
            var body = document.Body;
            Assert.NotNull(body);
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
            var mainPart = doc.MainDocumentPart;
            Assert.NotNull(mainPart);
            var document = mainPart.Document;
            Assert.NotNull(document);
            var body = document.Body;
            Assert.NotNull(body);
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
            var mainPart = doc.MainDocumentPart;
            Assert.NotNull(mainPart);
            var document = mainPart.Document;
            Assert.NotNull(document);
            var body = document.Body;
            Assert.NotNull(body);
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
            var mainPart = doc.MainDocumentPart;
            Assert.NotNull(mainPart);
            var document = mainPart.Document;
            Assert.NotNull(document);
            var body = document.Body;
            Assert.NotNull(body);
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
            var mainPart = doc.MainDocumentPart;
            Assert.NotNull(mainPart);
            var document = mainPart.Document;
            Assert.NotNull(document);
            var body = document.Body;
            Assert.NotNull(body);
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
            var mainPart = doc.MainDocumentPart;
            Assert.NotNull(mainPart);
            var document = mainPart.Document;
            Assert.NotNull(document);
            var body = document.Body;
            Assert.NotNull(body);
            var paragraph = body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().First();
            
            Assert.Equal("Heading1", paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
        }
    }

    [Fact]
    public void MissingTargetOperations_ShouldThrowHelpfulErrors()
    {
        // Arrange
        using var builder = DocumentBuilder.Create(_testFilePath);
        builder.AddParagraph("Existing paragraph");

        // Act & Assert
        var after = Assert.Throws<InvalidOperationException>(() => builder.InsertAfter("Missing", "Text"));
        var before = Assert.Throws<InvalidOperationException>(() => builder.InsertBefore("Missing", "Text"));
        var replace = Assert.Throws<InvalidOperationException>(() => builder.ReplaceParagraph("Missing", "Text"));
        var rich = Assert.Throws<InvalidOperationException>(() => builder.ReplaceWithRichContent("Missing", new()));

        Assert.Contains("Missing", after.Message);
        Assert.Contains("Missing", before.Message);
        Assert.Contains("Missing", replace.Message);
        Assert.Contains("Missing", rich.Message);
    }

    [Fact]
    public void DeleteParagraph_WhenTargetMissing_ShouldLeaveDocumentUnchanged()
    {
        // Arrange & Act
        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.AddParagraph("Keep this");
            builder.DeleteParagraph("Missing");
            builder.Save();
        }

        // Assert
        using var doc = WordprocessingDocument.Open(_testFilePath, false);
        var paragraphs = doc.MainDocumentPart!.Document!.Body!.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().ToList();
        Assert.Single(paragraphs);
        Assert.Equal("Keep this", paragraphs[0].InnerText);
    }

    [Fact]
    public void ApplyStyle_ShouldNoOpWithoutParagraphAndStyleLastParagraphWhenPresent()
    {
        // Arrange & Act
        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.ApplyStyle("Heading2");
            builder.AddParagraph("First");
            builder.AddParagraph("Second");
            builder.ApplyStyle("Quote");
            builder.Save();
        }

        // Assert
        using var doc = WordprocessingDocument.Open(_testFilePath, false);
        var paragraphs = doc.MainDocumentPart!.Document!.Body!.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().ToList();
        Assert.Equal(2, paragraphs.Count);
        Assert.Null(paragraphs[0].ParagraphProperties?.ParagraphStyleId);
        Assert.Equal("Quote", paragraphs[1].ParagraphProperties?.ParagraphStyleId?.Val?.Value);
    }

    [Fact]
    public void Save_WithDifferentPath_ShouldCloneDocument()
    {
        // Arrange
        var clonePath = Path.Combine(Path.GetTempPath(), $"test_advanced_clone_{Guid.NewGuid()}.docx");
        _additionalFiles.Add(clonePath);

        // Act
        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.AddParagraph("Clone me");
            builder.Save(clonePath);
        }

        // Assert
        Assert.True(File.Exists(_testFilePath));
        Assert.True(File.Exists(clonePath));
        using var doc = WordprocessingDocument.Open(clonePath, false);
        Assert.Contains("Clone me", doc.MainDocumentPart!.Document!.Body!.InnerText);
    }

    [Fact]
    public void MergeBatch_ShouldCreateOneDocumentPerRecord()
    {
        // Arrange
        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.AddParagraph("Hello {{name}}");
            builder.Save();
        }

        var outputPattern = Path.Combine(Path.GetTempPath(), $"merge_{Guid.NewGuid()}_{{index}}_{{name}}.docx");
        var expectedA = outputPattern.Replace("{index}", "0").Replace("{name}", "Ada");
        var expectedB = outputPattern.Replace("{index}", "1").Replace("{name}", "Bob");
        _additionalFiles.Add(expectedA);
        _additionalFiles.Add(expectedB);

        // Act
        var batchRunnerPath = Path.Combine(Path.GetTempPath(), $"merge_runner_{Guid.NewGuid()}.docx");
        _additionalFiles.Add(batchRunnerPath);
        using (var builder = DocumentBuilder.Create(batchRunnerPath))
        {
            builder.MergeBatch(new List<Dictionary<string, string>>
            {
                new() { ["name"] = "Ada" },
                new() { ["name"] = "Bob" }
            }, outputPattern, _testFilePath);
        }

        // Assert
        Assert.True(File.Exists(expectedA));
        Assert.True(File.Exists(expectedB));
        using var firstDoc = WordprocessingDocument.Open(expectedA, false);
        using var secondDoc = WordprocessingDocument.Open(expectedB, false);
        Assert.Contains("Hello Ada", firstDoc.MainDocumentPart!.Document!.Body!.InnerText);
        Assert.Contains("Hello Bob", secondDoc.MainDocumentPart!.Document!.Body!.InnerText);
    }

    public void Dispose()
    {
        if (File.Exists(_testFilePath))
        {
            File.Delete(_testFilePath);
        }

        foreach (var file in _additionalFiles)
        {
            if (File.Exists(file))
            {
                File.Delete(file);
            }
        }
    }
}
