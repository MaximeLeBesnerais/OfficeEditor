using DocxEditor.Core.Builders;
using DocxEditor.Core.Content;
using DocxEditor.Core.Models;
using DocumentFormat.OpenXml;
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
        Assert.Contains("Code", ensuredStyles);
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

        // One abstractNum + one numbering instance per list, with collision-free ids.
        var abstractNums = numbering.Elements<AbstractNum>().ToList();
        Assert.Equal(2, abstractNums.Count);
        Assert.Equal(2, abstractNums.Select(a => a.AbstractNumberId?.Value).Distinct().Count());

        var instances = numbering.Elements<NumberingInstance>().ToList();
        Assert.Equal(2, instances.Count);
        Assert.Equal(2, instances.Select(i => i.NumberID?.Value).Distinct().Count());

        // Every instance must reference an abstractNum that exists.
        var abstractIds = abstractNums.Select(a => a.AbstractNumberId?.Value).ToHashSet();
        Assert.All(instances, i => Assert.Contains(i.AbstractNumId?.Val?.Value, abstractIds));

        // Each list restarts at 1 (own abstractNum with its own start value).
        Assert.All(abstractNums, a =>
            Assert.Equal(1, a.Elements<Level>().First().StartNumberingValue?.Val?.Value));

        // CT_Numbering sequence: all abstractNum elements precede all num elements.
        var children = numbering.ChildElements.ToList();
        var lastAbstractIndex = children.FindLastIndex(c => c is AbstractNum);
        var firstInstanceIndex = children.FindIndex(c => c is NumberingInstance);
        Assert.True(lastAbstractIndex < firstInstanceIndex,
            "abstractNum elements must precede num elements in numbering.xml");

        OpenXmlAssert.NoDocxValidationErrors(_testFilePath);
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

        var bulletNumId = paragraphs[0].ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value;
        var orderedNumId = paragraphs[1].ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value;
        Assert.NotNull(bulletNumId);
        Assert.NotNull(orderedNumId);
        Assert.NotEqual(bulletNumId, orderedNumId);

        // Each numId must resolve to an abstractNum with the matching format —
        // binding bullets to a decimal definition (or vice versa) was the collision bug.
        var numbering = doc.MainDocumentPart!.NumberingDefinitionsPart!.Numbering;
        Assert.NotNull(numbering);
        Assert.Equal(NumberFormatValues.Bullet, ResolveNumberingFormat(numbering, bulletNumId!.Value));
        Assert.Equal(NumberFormatValues.Decimal, ResolveNumberingFormat(numbering, orderedNumId!.Value));
    }

    [Fact]
    public void AddRichContent_WithMultipleLists_ShouldReuseAbstractDefinitionsPerKind()
    {
        var blocks = new ContentBlockBuilder()
            .AddList(false, ["Bullet A1", "Bullet A2"])
            .AddList(true, ["Ordered B1", "Ordered B2"])
            .AddList(false, ["Bullet C1"])
            .AddList(true, ["Ordered D1", "Ordered D2"])
            .Build();

        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.AddRichContent(blocks);
            builder.Save();
        }

        using (var doc = WordprocessingDocument.Open(_testFilePath, false))
        {
            var numbering = doc.MainDocumentPart!.NumberingDefinitionsPart!.Numbering;
            Assert.NotNull(numbering);
            var instances = numbering.Elements<NumberingInstance>().ToList();
            var abstractNums = numbering.Elements<AbstractNum>().ToList();

            // Every list got its own instance so its counter restarts at 1.
            Assert.Equal(4, instances.Count);
            Assert.Equal(4, instances.Select(i => i.NumberID?.Value).Distinct().Count());

            // Same-kind lists share one generated abstract definition instead of
            // duplicating it: one bullet definition, one ordered definition.
            Assert.Equal(2, abstractNums.Count);
            var instanceAbstractIds = instances.Select(i => i.AbstractNumId?.Val?.Value).ToList();
            Assert.Equal(2, instanceAbstractIds.Distinct().Count());

            // Every list paragraph's numbering id resolves to a real instance that
            // resolves to a real abstract definition of the matching format.
            var paragraphs = doc.MainDocumentPart.Document!.Body!.Elements<Paragraph>().ToList();
            Assert.Equal(7, paragraphs.Count);
            var listIds = new[] { 0, 2, 4, 5 }
                .Select(i => paragraphs[i].ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value)
                .ToList();
            Assert.All(listIds, id => Assert.NotNull(id));

            for (int i = 0; i < listIds.Count; i++)
            {
                var listId = listIds[i]!.Value;
                var instance = instances.Single(x => x.NumberID?.Value == listId);
                var abstractNum = abstractNums.Single(a => a.AbstractNumberId?.Value == instance.AbstractNumId?.Val?.Value);
                var expected = i % 2 == 0 ? NumberFormatValues.Bullet : NumberFormatValues.Decimal;
                Assert.Equal(expected, abstractNum.Elements<Level>().First().NumberingFormat?.Val?.Value);
            }

            // Bullet lists (paragraphs 0, 4) share one definition; ordered lists (2, 5) share the other.
            var bulletIds = new[] { 0, 4 }
                .Select(i => instances.Single(x => x.NumberID?.Value == paragraphs[i].ParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value).AbstractNumId!.Val!.Value)
                .Distinct()
                .ToList();
            var orderedIds = new[] { 2, 5 }
                .Select(i => instances.Single(x => x.NumberID?.Value == paragraphs[i].ParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value).AbstractNumId!.Val!.Value)
                .Distinct()
                .ToList();
            Assert.Single(bulletIds);
            Assert.Single(orderedIds);
            Assert.NotEqual(bulletIds[0], orderedIds[0]);
        }

        OpenXmlAssert.NoDocxValidationErrors(_testFilePath);
    }

    [Fact]
    public void AddRichContent_WithPreExistingNumbering_ShouldAllocateCollisionFreeIdsAndRestart()
    {
        // Fixture: document already has a decimal list (abstractNumId 5, numId 7).
        using (var doc = WordprocessingDocument.Create(_testFilePath, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var existingAbstract = new AbstractNum(
                new Level(
                    new StartNumberingValue { Val = 1 },
                    new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new LevelText { Val = "%1." }
                ) { LevelIndex = 0 }
            ) { AbstractNumberId = 5 };
            var existingInstance = new NumberingInstance(
                new AbstractNumId { Val = 5 }
            ) { NumberID = 7 };
            var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = new Numbering(existingAbstract, existingInstance);

            var existingItem = new Paragraph(
                new ParagraphProperties(
                    new NumberingProperties(
                        new NumberingLevelReference { Val = 0 },
                        new NumberingId { Val = 7 })),
                new Run(new Text("Existing item")));
            mainPart.Document = new Document(new Body(existingItem));
            mainPart.Document.Save();
        }

        var blocks = new ContentBlockBuilder()
            .AddList(true, ["New A1", "New A2"])
            .AddList(true, ["New B1", "New B2"])
            .Build();

        using (var builder = DocumentBuilder.Open(_testFilePath))
        {
            builder.AddRichContent(blocks);
            builder.Save();
        }

        using (var doc = WordprocessingDocument.Open(_testFilePath, false))
        {
            var numbering = doc.MainDocumentPart!.NumberingDefinitionsPart!.Numbering;
            Assert.NotNull(numbering);
            var instances = numbering.Elements<NumberingInstance>().ToList();
            var abstractNums = numbering.Elements<AbstractNum>().ToList();

            // Pre-existing ids remain untouched; new ids are max+1 allocations.
            Assert.Contains(instances, i => i.NumberID?.Value == 7 && i.AbstractNumId?.Val?.Value == 5);
            var newInstances = instances.Where(i => i.NumberID?.Value > 7).ToList();
            Assert.Equal(2, newInstances.Count);
            Assert.All(newInstances, i => Assert.True(i.AbstractNumId?.Val?.Value > 5));

            // The two new lists are independent: distinct instances, sharing one generated
            // abstract definition (same semantics = same definition, reused not duplicated),
            // each list restarting at 1 instead of continuing 1,2,3,4.
            var newAbstractIds = newInstances.Select(i => i.AbstractNumId?.Val?.Value).ToList();
            Assert.All(newAbstractIds, id => Assert.NotNull(id));
            Assert.Single(newAbstractIds.Distinct());
            Assert.Equal(2, newInstances.Select(i => i.NumberID?.Value).Distinct().Count());
            foreach (var abstractId in newAbstractIds.Distinct())
            {
                var abstractNum = abstractNums.Single(a => a.AbstractNumberId?.Value == abstractId);
                Assert.Equal(1, abstractNum.Elements<Level>().First().StartNumberingValue?.Val?.Value);
            }

            // The pre-existing list paragraph still binds to its original instance.
            var paragraphs = doc.MainDocumentPart.Document!.Body!.Elements<Paragraph>().ToList();
            Assert.Equal(7, paragraphs[0].ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value);

            // Each new list's items share one fresh instance id.
            var listAId = paragraphs[1].ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value;
            var listBId = paragraphs[3].ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value;
            Assert.NotNull(listAId);
            Assert.NotNull(listBId);
            Assert.NotEqual(listAId, listBId);
            Assert.Equal(listAId, paragraphs[2].ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value);
            Assert.Equal(listBId, paragraphs[4].ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value);
        }

        OpenXmlAssert.NoDocxValidationErrors(_testFilePath);
    }

    [Fact]
    public void AddRichContent_WithTable_ShouldEmitTblGridMatchingColumnCount()
    {
        var blocks = new ContentBlockBuilder()
            .AddTable(new List<List<string>>
            {
                new() { "H1", "H2", "H3" },
                new() { "A", "B", "C" }
            })
            .Build();

        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.AddRichContent(blocks);
            builder.Save();
        }

        using (var doc = WordprocessingDocument.Open(_testFilePath, false))
        {
            var table = doc.MainDocumentPart!.Document!.Body!.Elements<Table>().Single();

            // CT_Tbl order: tblPr, tblGrid, then rows.
            Assert.IsType<TableProperties>(table.ChildElements[0]);
            var grid = Assert.IsType<TableGrid>(table.ChildElements[1]);
            Assert.Equal(3, grid.Elements<GridColumn>().Count());
        }

        OpenXmlAssert.NoDocxValidationErrors(_testFilePath);
    }

    [Fact]
    public void AddRichContent_WithEmptyTable_ShouldRenderNothing()
    {
        var blocks = new List<ContentBlock>
        {
            new ParagraphBlock { Text = "Before" },
            new TableBlock { Rows = [] },
            new ParagraphBlock { Text = "After" }
        };

        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.AddRichContent(blocks);
            builder.Save();
        }

        using (var doc = WordprocessingDocument.Open(_testFilePath, false))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            Assert.Empty(body.Elements<Table>());
            Assert.Equal(2, body.Elements<Paragraph>().Count());
        }

        OpenXmlAssert.NoDocxValidationErrors(_testFilePath);
    }

    private static NumberFormatValues? ResolveNumberingFormat(Numbering numbering, int numberingId)
    {
        var instance = numbering.Elements<NumberingInstance>().Single(i => i.NumberID?.Value == numberingId);
        var abstractNum = numbering.Elements<AbstractNum>().Single(a => a.AbstractNumberId?.Value == instance.AbstractNumId?.Val?.Value);
        return abstractNum.Elements<Level>().First().NumberingFormat?.Val?.Value;
    }

    public void Dispose()
    {
        if (File.Exists(_testFilePath))
        {
            File.Delete(_testFilePath);
        }
    }
}
