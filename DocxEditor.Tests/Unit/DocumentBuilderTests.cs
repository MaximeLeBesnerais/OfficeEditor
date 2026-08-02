using DocxEditor.Core.Builders;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

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
            var mainPart = doc.MainDocumentPart;
            Assert.NotNull(mainPart);
            var document = mainPart.Document;
            Assert.NotNull(document);
            var body = document.Body;
            Assert.NotNull(body);
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
            var mainPart = doc.MainDocumentPart;
            Assert.NotNull(mainPart);
            var document = mainPart.Document;
            Assert.NotNull(document);
            var body = document.Body;
            Assert.NotNull(body);
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
            var mainPart = doc.MainDocumentPart;
            Assert.NotNull(mainPart);
            var document = mainPart.Document;
            Assert.NotNull(document);
            var body = document.Body;
            Assert.NotNull(body);
            var paragraphs = body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();
            Assert.Contains(paragraphs, p => p.InnerText == "Hello World");
        }
    }

    [Fact]
    public void ReplaceText_ShouldReplaceInsideTablesHeadersAndFooters()
    {
        // Arrange: body paragraph, a table cell, a header, and a footer all carry the token.
        using (var doc = WordprocessingDocument.Create(_testFilePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var table = new Table(
                new TableProperties(),
                new TableGrid(new GridColumn()),
                new TableRow(new TableCell(new Paragraph(new Run(new Text("Cell {{TOKEN}}"))))));
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Body {{TOKEN}}"))),
                table));

            var headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new Header(new Paragraph(new Run(new Text("Header {{TOKEN}}"))));
            headerPart.Header.Save();

            var footerPart = mainPart.AddNewPart<FooterPart>();
            footerPart.Footer = new Footer(new Paragraph(new Run(new Text("Footer {{TOKEN}}"))));
            footerPart.Footer.Save();

            mainPart.Document.Save();
        }

        // Act
        using (var builder = DocumentBuilder.Open(_testFilePath))
        {
            builder.ReplaceText("{{TOKEN}}", "Done");
            builder.Save();
        }

        // Assert
        using (var doc = WordprocessingDocument.Open(_testFilePath, false))
        {
            var mainPart = doc.MainDocumentPart!;
            Assert.DoesNotContain("{{TOKEN}}", mainPart.Document!.Body!.InnerText);
            Assert.Contains("Cell Done", mainPart.Document.Body!.Descendants<TableCell>().Single().InnerText);
            Assert.Contains("Header Done", mainPart.HeaderParts.Single().Header!.InnerText);
            Assert.Contains("Footer Done", mainPart.FooterParts.Single().Footer!.InnerText);
        }
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void ReplaceText_WithEmptyFind_ShouldThrowArgumentException(string? find)
    {
        // Empty find would insert the replacement between every character of the document.
        using var builder = DocumentBuilder.Create(_testFilePath);
        builder.AddParagraph("Some text");

        Assert.Throws<ArgumentException>(() => builder.ReplaceText(find!, "x"));
    }

    [Fact]
    public void Create_SaveToBytes_ReturnsValidDocx()
    {
        // Arrange
        byte[] bytes;

        // Act
        using (var builder = DocumentBuilder.Create())
        {
            builder.AddParagraph("Hello bytes");
            bytes = builder.SaveToBytes();
        }

        // Assert
        Assert.NotEmpty(bytes);
        using var stream = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        Assert.NotNull(doc.MainDocumentPart);
        Assert.NotNull(doc.MainDocumentPart!.Document);
        Assert.NotNull(doc.MainDocumentPart.Document.Body);
    }

    [Fact]
    public void Create_Save_Stream_ReturnsValidDocx()
    {
        // Arrange
        using var outputStream = new MemoryStream();

        // Act
        using (var builder = DocumentBuilder.Create())
        {
            builder.AddParagraph("Hello stream");
            builder.Save(outputStream);
        }

        // Assert
        Assert.True(outputStream.CanRead);
        Assert.True(outputStream.CanWrite);
        Assert.True(outputStream.Length > 0);
        outputStream.Position = 0;
        using var doc = WordprocessingDocument.Open(outputStream, false);
        Assert.NotNull(doc.MainDocumentPart);
        var body = doc.MainDocumentPart!.Document!.Body!;
        Assert.Contains(body.Elements<Paragraph>(), p => p.InnerText == "Hello stream");
    }

    [Fact]
    public void Open_FromBytes_AndSaveToBytes_RoundtripsContent()
    {
        // Arrange
        byte[] originalBytes;
        using (var builder = DocumentBuilder.Create())
        {
            builder.AddParagraph("Original text");
            originalBytes = builder.SaveToBytes();
        }

        // Act
        byte[] modifiedBytes;
        using (var builder = DocumentBuilder.Open(originalBytes))
        {
            builder.AddParagraph("Added text");
            modifiedBytes = builder.SaveToBytes();
        }

        // Assert
        using var stream = new MemoryStream(modifiedBytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        Assert.Contains(body.Elements<Paragraph>(), p => p.InnerText == "Original text");
        Assert.Contains(body.Elements<Paragraph>(), p => p.InnerText == "Added text");
    }

    [Fact]
    public void Save_WithNoPath_OnPathlessDocument_ThrowsInvalidOperationException()
    {
        // Arrange
        using var builder = DocumentBuilder.Create();
        builder.AddParagraph("No path");

        // Act & Assert
        Assert.Throws<InvalidOperationException>(() => builder.Save());
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void Save_WithEmptyOrWhitespacePath_ThrowsArgumentExceptionAndPreservesSource(string path)
    {
        // Gap fixed: an empty destination was a silent no-op on file-backed documents while
        // whitespace materialized a junk file named after the whitespace; both are now rejected
        // as argument errors up front, matching the Create/Open path contract.
        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.AddParagraph("Keep");
            Assert.Throws<ArgumentException>(() => builder.Save(path));
        }

        using var doc = WordprocessingDocument.Open(_testFilePath, false);
        Assert.Contains("Keep", doc.MainDocumentPart!.Document!.Body!.InnerText);
    }

    [Fact]
    public void Save_WithWhitespacePath_CreatesNoJunkFile()
    {
        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.AddParagraph("Keep");
            Assert.Throws<ArgumentException>(() => builder.Save("   "));
        }

        Assert.False(File.Exists("   "), "A rejected whitespace path must not create a file.");
    }

    [Fact]
    public void SaveToBytes_AfterAddingMarkdown_ReturnsValidDocx()
    {
        // Arrange
        byte[] bytes;
        var markdown = "# Markdown Title\n\nMarkdown paragraph.";

        // Act
        using (var builder = DocumentBuilder.Create())
        {
            builder.AddMarkdown(markdown);
            bytes = builder.SaveToBytes();
        }

        // Assert
        Assert.NotEmpty(bytes);
        using var stream = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        var text = body.InnerText;
        Assert.Contains("Markdown Title", text);
        Assert.Contains("Markdown paragraph", text);
    }

    public void Dispose()
    {
        if (File.Exists(_testFilePath))
        {
            File.Delete(_testFilePath);
        }
    }
}
