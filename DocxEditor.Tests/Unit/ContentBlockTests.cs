using DocxEditor.Core.Builders;
using DocxEditor.Core.Content;
using DocxEditor.Core.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeEditor.Core.Models;

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
            var mainPart = doc.MainDocumentPart;
            Assert.NotNull(mainPart);
            var document = mainPart.Document;
            Assert.NotNull(document);
            var body = document.Body;
            Assert.NotNull(body);
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
            var mainPart = doc.MainDocumentPart;
            Assert.NotNull(mainPart);
            var document = mainPart.Document;
            Assert.NotNull(document);
            var body = document.Body;
            Assert.NotNull(body);
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
            var mainPart = doc.MainDocumentPart;
            Assert.NotNull(mainPart);
            var document = mainPart.Document;
            Assert.NotNull(document);
            var body = document.Body;
            Assert.NotNull(body);
            var text = body.InnerText;
            
            Assert.DoesNotContain("Old content", text);
            Assert.Contains("New Title", text);
            Assert.Contains("New paragraph", text);
            Assert.Contains("Keep this", text);
        }
    }

    [Fact]
    public void ContentBlockBuilder_ShouldBuildRichBlockTypesAndModelDefaults()
    {
        // Arrange
        var inlineFormats = new List<InlineFormat>
        {
            new() { Type = "bold", Text = "Bold" },
            new() { Type = "italic", Text = "Italic" }
        };

        // Act
        var blocks = new ContentBlockBuilder()
            .AddParagraph("Plain", "CustomParagraph")
            .AddParagraph("Formatted", inlineFormats, "FormattedStyle")
            .AddHeading(3, "Heading", "CustomHeading")
            .AddList(true, new List<string> { "One", "Two" }, "ListStyle")
            .AddTable(new List<List<string>> { new() { "A", "B" } })
            .AddBlockquote("Quote", "QuoteStyle")
            .AddCode("var x = 1;", "csharp", "CodeStyle")
            .AddHorizontalRule()
            .AddCustom("tip", "Tip text", "TipStyle")
            .Build();

        // Assert
        Assert.Collection(blocks,
            block =>
            {
                var paragraph = Assert.IsType<ParagraphBlock>(block);
                Assert.Equal("paragraph", paragraph.Type);
                Assert.Equal("Plain", paragraph.Text);
                Assert.Equal("CustomParagraph", paragraph.Style);
            },
            block =>
            {
                var paragraph = Assert.IsType<ParagraphBlock>(block);
                Assert.Same(inlineFormats, paragraph.InlineFormats);
            },
            block =>
            {
                var heading = Assert.IsType<HeadingBlock>(block);
                Assert.Equal("heading", heading.Type);
                Assert.Equal(3, heading.Level);
            },
            block =>
            {
                var list = Assert.IsType<ListBlock>(block);
                Assert.Equal("list", list.Type);
                Assert.True(list.Ordered);
            },
            block =>
            {
                var table = Assert.IsType<TableBlock>(block);
                Assert.Equal("table", table.Type);
                Assert.Equal("A", table.Rows[0].Cells[0].Text);
            },
            block => Assert.Equal("blockquote", Assert.IsType<BlockquoteBlock>(block).Type),
            block =>
            {
                var code = Assert.IsType<CodeBlock>(block);
                Assert.Equal("code", code.Type);
                Assert.Equal("csharp", code.Language);
            },
            block => Assert.Equal("horizontalRule", Assert.IsType<HorizontalRuleBlock>(block).Type),
            block =>
            {
                var custom = Assert.IsType<CustomBlock>(block);
                Assert.Equal("custom", custom.Type);
                Assert.Equal("tip", custom.CustomType);
            });
    }

    [Fact]
    public void ContentBlockRenderer_ShouldRenderRichBlocksAndInlineFormats()
    {
        // Arrange
        var ensuredStyles = new List<string>();
        var renderer = new ContentBlockRenderer(StyleMapping.Default, ensureStyle: ensuredStyles.Add);
        var body = new Body();
        var blocks = new List<ContentBlock>
        {
            new ParagraphBlock
            {
                InlineFormats = new List<InlineFormat>
                {
                    new() { Type = "bold", Text = "B" },
                    new() { Type = "italic", Text = "I" },
                    new() { Type = "underline", Text = "U" },
                    new() { Type = "strikethrough", Text = "S" },
                    new() { Type = "code", Text = "C" },
                    new() { Type = "unknown", Text = "?" }
                }
            },
            new BlockquoteBlock { Text = "Quote" },
            new CodeBlock { Text = "Code" },
            new HorizontalRuleBlock(),
            new CustomBlock { CustomType = "warning", Text = "Warning" }
        };

        // Act
        renderer.Render(body, blocks);

        // Assert
        var paragraphs = body.Elements<Paragraph>().ToList();
        Assert.Equal(5, paragraphs.Count);
        Assert.Equal("BIUSC?", paragraphs[0].InnerText);
        Assert.NotNull(paragraphs[0].Elements<Run>().ElementAt(0).RunProperties?.Bold);
        Assert.NotNull(paragraphs[0].Elements<Run>().ElementAt(1).RunProperties?.Italic);
        Assert.NotNull(paragraphs[0].Elements<Run>().ElementAt(2).RunProperties?.Underline);
        Assert.NotNull(paragraphs[0].Elements<Run>().ElementAt(3).RunProperties?.Strike);
        Assert.NotNull(paragraphs[0].Elements<Run>().ElementAt(4).RunProperties?.RunStyle);
        Assert.Equal("Quote", paragraphs[1].ParagraphProperties?.ParagraphStyleId?.Val?.Value);
        Assert.Equal("Code", paragraphs[2].ParagraphProperties?.ParagraphStyleId?.Val?.Value);
        Assert.NotNull(paragraphs[3].ParagraphProperties?.ParagraphBorders?.BottomBorder);
        Assert.Equal("Warning", paragraphs[4].ParagraphProperties?.ParagraphStyleId?.Val?.Value);
        Assert.Contains("Normal", ensuredStyles);
        Assert.Contains("Warning", ensuredStyles);
    }

    [Fact]
    public void ContentBlockRenderer_WithEmptyStyleMapping_ShouldRenderWithoutParagraphStyles()
    {
        // Arrange
        var renderer = new ContentBlockRenderer(new StyleMapping());
        var body = new Body();

        // Act
        renderer.Render(body, new List<ContentBlock>
        {
            new ParagraphBlock { Text = "Plain" },
            new HeadingBlock { Level = 2, Text = "Heading" },
            new ListBlock { Ordered = false, Items = new List<string> { "Item" } },
            new CustomBlock { CustomType = "missing", Text = "Custom" }
        });

        // Assert
        var paragraphs = body.Elements<Paragraph>().ToList();
        Assert.Equal(4, paragraphs.Count);
        Assert.All(paragraphs.Where(p => p.InnerText != "Item"), p => Assert.Null(p.ParagraphProperties));
        Assert.NotNull(paragraphs[2].ParagraphProperties?.NumberingProperties);
        Assert.Null(paragraphs[2].ParagraphProperties?.ParagraphStyleId);
    }

    [Fact]
    public void AddRichContent_WithLists_ShouldCreateNumberingDefinitions()
    {
        var blocks = new ContentBlockBuilder()
            .AddList(false, ["Bullet item"])
            .AddList(true, ["Numbered item"])
            .Build();

        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.AddRichContent(blocks);
            builder.Save();
        }

        using var doc = WordprocessingDocument.Open(_testFilePath, false);
        var numberingPart = doc.MainDocumentPart!.NumberingDefinitionsPart;
        Assert.NotNull(numberingPart);

        var numbering = numberingPart.Numbering;
        Assert.NotNull(numbering);

        var abstractNums = numbering.Elements<AbstractNum>().ToList();
        Assert.Equal(2, abstractNums.Count);

        var instances = numbering.Elements<NumberingInstance>().ToList();
        Assert.Equal(2, instances.Count);

        Assert.Contains(abstractNums, a => a.AbstractNumberId?.Value == 1);
        Assert.Contains(abstractNums, a => a.AbstractNumberId?.Value == 2);
        Assert.Contains(instances, i => i.NumberID?.Value == 1);
        Assert.Contains(instances, i => i.NumberID?.Value == 2);
    }

    [Fact]
    public void AddRichContent_WithLists_ShouldSetNumberingPropertiesOnParagraphs()
    {
        var blocks = new ContentBlockBuilder()
            .AddList(false, ["Unordered"])
            .AddList(true, ["Ordered"])
            .Build();

        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.AddRichContent(blocks);
            builder.Save();
        }

        using var doc = WordprocessingDocument.Open(_testFilePath, false);
        var paragraphs = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ToList();

        Assert.Equal(2, paragraphs.Count);

        var bulletProps = paragraphs[0].ParagraphProperties?.NumberingProperties;
        Assert.NotNull(bulletProps);
        Assert.Equal(2, bulletProps!.NumberingId?.Val?.Value);

        var numberProps = paragraphs[1].ParagraphProperties?.NumberingProperties;
        Assert.NotNull(numberProps);
        Assert.Equal(1, numberProps!.NumberingId?.Val?.Value);
    }

    public void Dispose()
    {
        if (File.Exists(_testFilePath))
        {
            File.Delete(_testFilePath);
        }
    }
}
