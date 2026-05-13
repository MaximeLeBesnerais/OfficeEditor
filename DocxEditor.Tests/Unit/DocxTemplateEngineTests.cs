using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Variables;

namespace DocxEditor.Tests.Unit;

public sealed class DocxTemplateEngineTests : IDisposable
{
    private readonly List<string> _tempFiles = [];

    [Fact]
    public void Process_RewritesParagraph_WhenTemplateChangesText()
    {
        // Arrange
        var path = CreateDocument(body => body.Append(CreateParagraph("Before ", "{{#if show}}Visible{{/if}}")));

        // Act
        Process(path, new Dictionary<string, object> { ["show"] = true });

        // Assert
        using var document = WordprocessingDocument.Open(path, false);
        var paragraph = GetFirstParagraph(document);
        Assert.Equal("Before Visible", paragraph.InnerText);
        var run = Assert.Single(paragraph.Elements<Run>());
        Assert.Equal("Before Visible", run.GetFirstChild<Text>()!.Text);
    }

    [Fact]
    public void Process_LeavesParagraphUnchanged_WhenItDoesNotContainTemplateMarkers()
    {
        // Arrange
        var path = CreateDocument(body => body.Append(CreateParagraph("Plain ", "text")));

        // Act
        Process(path, new Dictionary<string, object> { ["show"] = true });

        // Assert
        using var document = WordprocessingDocument.Open(path, false);
        var paragraph = GetFirstParagraph(document);
        Assert.Equal("Plain text", paragraph.InnerText);
        Assert.Equal(["Plain ", "text"], paragraph.Elements<Run>().Select(run => run.GetFirstChild<Text>()!.Text).ToArray());
    }

    [Theory]
    [InlineData("{{#if enabled}}Yes{{/if}}", true, "Yes")]
    [InlineData("{{#if enabled}}Yes{{/if}}", false, "")]
    [InlineData("{{#ifnot enabled}}No{{/ifnot}}", true, "")]
    [InlineData("{{#ifnot enabled}}No{{/ifnot}}", false, "No")]
    public void Process_EvaluatesIfAndIfNot_ForBooleanValues(string template, bool enabled, string expected)
    {
        // Arrange
        var path = CreateDocument(body => body.Append(CreateParagraph(template)));

        // Act
        Process(path, new Dictionary<string, object> { ["enabled"] = enabled });

        // Assert
        using var document = WordprocessingDocument.Open(path, false);
        Assert.Equal(expected, GetFirstParagraph(document).InnerText);
    }

    [Theory]
    [InlineData("text", "Yes")]
    [InlineData("", "")]
    [InlineData("false", "")]
    [InlineData("FALSE", "")]
    [InlineData("0", "")]
    public void Process_EvaluatesStringTruthiness(string value, string expected)
    {
        // Arrange
        var path = CreateDocument(body => body.Append(CreateParagraph("{{#if label}}Yes{{/if}}")));

        // Act
        Process(path, new Dictionary<string, object> { ["label"] = value });

        // Assert
        using var document = WordprocessingDocument.Open(path, false);
        Assert.Equal(expected, GetFirstParagraph(document).InnerText);
    }

    [Theory]
    [InlineData("{{#if count == 5}}Equal{{/if}}", 5, "Equal")]
    [InlineData("{{#if count != 4}}Not equal{{/if}}", 5, "Not equal")]
    [InlineData("{{#if count > 4}}Greater{{/if}}", 5, "Greater")]
    [InlineData("{{#if count < 6}}Less{{/if}}", 5, "Less")]
    [InlineData("{{#if count >= 5}}Greater or equal{{/if}}", 5, "Greater or equal")]
    [InlineData("{{#if count <= 5}}Less or equal{{/if}}", 5, "Less or equal")]
    public void Process_EvaluatesNumericComparisons(string template, int count, string expected)
    {
        // Arrange
        var path = CreateDocument(body => body.Append(CreateParagraph(template)));

        // Act
        Process(path, new Dictionary<string, object> { ["count"] = count });

        // Assert
        using var document = WordprocessingDocument.Open(path, false);
        Assert.Equal(expected, GetFirstParagraph(document).InnerText);
    }

    [Theory]
    [InlineData("{{#if status == 'open'}}Open{{/if}}", "open", "Open")]
    [InlineData("{{#if status != \"closed\"}}Not closed{{/if}}", "open", "Not closed")]
    [InlineData("{{#if status > alpha}}After alpha{{/if}}", "beta", "After alpha")]
    [InlineData("{{#if status < omega}}Before omega{{/if}}", "beta", "Before omega")]
    [InlineData("{{#if status >= beta}}At least beta{{/if}}", "beta", "At least beta")]
    [InlineData("{{#if status <= beta}}At most beta{{/if}}", "beta", "At most beta")]
    public void Process_EvaluatesStringComparisons(string template, string status, string expected)
    {
        // Arrange
        var path = CreateDocument(body => body.Append(CreateParagraph(template)));

        // Act
        Process(path, new Dictionary<string, object> { ["status"] = status });

        // Assert
        using var document = WordprocessingDocument.Open(path, false);
        Assert.Equal(expected, GetFirstParagraph(document).InnerText);
    }

    [Fact]
    public void Process_ReturnsFalse_ForComparisonWithMissingVariable()
    {
        // Arrange
        var path = CreateDocument(body => body.Append(CreateParagraph("Before {{#if missing > 1}}Visible{{/if}} After")));

        // Act
        Process(path, []);

        // Assert
        using var document = WordprocessingDocument.Open(path, false);
        Assert.Equal("Before  After", GetFirstParagraph(document).InnerText);
    }

    [Fact]
    public void Process_ExpandsLoop_ForTwoRowsAndNullValues()
    {
        // Arrange
        var path = CreateDocument(body => body.Append(CreateParagraph("{{#each rows}}{name}:{value}{{/each}}")));
        var data = new Dictionary<string, object>
        {
            ["rows"] = new List<Dictionary<string, object>>
            {
                new() { ["name"] = "First", ["value"] = 1 },
                new() { ["name"] = "Second", ["value"] = null! }
            }
        };

        // Act
        Process(path, data);

        // Assert
        using var document = WordprocessingDocument.Open(path, false);
        Assert.Equal("First:1\nSecond:", GetFirstParagraph(document).InnerText);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Process_RemovesLoopContent_WhenSourceIsMissingOrNotAList(bool includeNonListSource)
    {
        // Arrange
        var path = CreateDocument(body => body.Append(CreateParagraph("Before {{#each rows}}{name}{{/each}} After")));
        var data = includeNonListSource
            ? new Dictionary<string, object> { ["rows"] = "not a list" }
            : [];

        // Act
        Process(path, data);

        // Assert
        using var document = WordprocessingDocument.Open(path, false);
        Assert.Equal("Before  After", GetFirstParagraph(document).InnerText);
    }

    [Fact]
    public void Process_DoesNotThrow_WhenMainDocumentPartIsMissing()
    {
        // Arrange
        var path = CreateWordprocessingDocument(_ => { });

        // Act / Assert
        using var document = WordprocessingDocument.Open(path, true);
        var exception = Record.Exception(() => new DocxTemplateEngine().Process(document, []));
        Assert.Null(exception);
    }

    [Fact]
    public void Process_DoesNotThrow_WhenBodyIsMissing()
    {
        // Arrange
        var path = CreateWordprocessingDocument(document => AddMainDocumentPart(document).Document = new Document());

        // Act / Assert
        using var wordDocument = WordprocessingDocument.Open(path, true);
        var exception = Record.Exception(() => new DocxTemplateEngine().Process(wordDocument, []));
        Assert.Null(exception);
    }

    [Fact]
    public void Process_DoesNotThrow_WhenBodyIsEmpty()
    {
        // Arrange
        var path = CreateDocument(_ => { });

        // Act / Assert
        using var document = WordprocessingDocument.Open(path, true);
        var exception = Record.Exception(() => new DocxTemplateEngine().Process(document, []));
        Assert.Null(exception);
    }

    public void Dispose()
    {
        foreach (var file in _tempFiles)
        {
            if (File.Exists(file))
            {
                File.Delete(file);
            }
        }
    }

    private string CreateDocument(Action<Body> configureBody)
    {
        return CreateWordprocessingDocument(document =>
        {
            var mainDocumentPart = AddMainDocumentPart(document);
            var body = new Body();
            configureBody(body);
            mainDocumentPart.Document = new Document(body);
        });
    }

    private string CreateWordprocessingDocument(Action<WordprocessingDocument> configure)
    {
        var path = Path.Combine(Path.GetTempPath(), $"docx_template_engine_{Guid.NewGuid():N}.docx");
        _tempFiles.Add(path);

        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        configure(document);
        return path;
    }

    private static MainDocumentPart AddMainDocumentPart(WordprocessingDocument document)
    {
        return document.AddMainDocumentPart();
    }

    private static void Process(string path, Dictionary<string, object> data)
    {
        using var document = WordprocessingDocument.Open(path, true);
        new DocxTemplateEngine().Process(document, data);
    }

    private static Paragraph CreateParagraph(params string[] runs)
    {
        return new Paragraph(runs.Select(CreateRun));
    }

    private static Run CreateRun(string text)
    {
        return new Run(CreateText(text));
    }

    private static Text CreateText(string text)
    {
        return new Text(text) { Space = SpaceProcessingModeValues.Preserve };
    }

    private static Paragraph GetFirstParagraph(WordprocessingDocument document)
    {
        return document.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();
    }
}
