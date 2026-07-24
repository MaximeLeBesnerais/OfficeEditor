using DocxEditor.Core.Models;
using DocxEditor.Core.Serialization;

namespace DocxEditor.Tests.Unit;

public class InstructionParserBranchTests
{
    [Theory]
    [InlineData("null")]
    [InlineData("{}")]
    [InlineData(@"{ ""operations"": null }")]
    public void JsonParser_WithMissingOperations_ShouldThrowArgumentException(string json)
    {
        var parser = new DocxJsonInstructionParser();

        var exception = Assert.Throws<ArgumentException>(() => parser.Parse(json));
        Assert.Contains("Invalid JSON", exception.Message);
    }

    [Theory]
    [InlineData(@"{ ""operations"": [ { ""type"": ""create"" } ] }")]
    [InlineData(@"{ ""OPERATIONS"": [ { ""TYPE"": ""CREATE"" } ] }")]
    public void JsonParser_WithCreateInstructionAndCaseVariations_ShouldParse(string json)
    {
        var instructions = new DocxJsonInstructionParser().Parse(json);

        Assert.Single(instructions.Operations);
        Assert.IsType<CreateDocumentInstruction>(instructions.Operations[0]);
    }

    [Fact]
    public void JsonParser_WithAllSupportedInstructionsAndOptionalFieldsMissing_ShouldParse()
    {
        var json = """
        {
          "operations": [
            { "type": "addParagraph", "text": "No style" },
            { "type": "replaceText", "find": "old", "replace": "" },
            { "type": "insertAfter", "target": "marker", "content": { "text": "Inserted", "style": "Body" } }
          ]
        }
        """;

        var instructions = new DocxJsonInstructionParser().Parse(json);

        Assert.Collection(instructions.Operations,
            first =>
            {
                var add = Assert.IsType<AddParagraphInstruction>(first);
                Assert.Equal("No style", add.Text);
                Assert.Null(add.Style);
            },
            second =>
            {
                var replace = Assert.IsType<ReplaceTextInstruction>(second);
                Assert.Equal("old", replace.Find);
                Assert.Equal("", replace.Replace);
            },
            third =>
            {
                var insert = Assert.IsType<InsertAfterInstruction>(third);
                Assert.Equal("marker", insert.Target);
                Assert.Equal("Inserted", insert.Content.Text);
                Assert.Equal("Body", insert.Content.Style);
            });
    }

    [Theory]
    [InlineData(@"{ ""operations"": [ { ""type"": null } ] }", typeof(NotSupportedException))]
    [InlineData(@"{ ""operations"": [ { ""type"": ""addParagraph"" } ] }", typeof(ArgumentException))]
    [InlineData(@"{ ""operations"": [ { ""type"": ""replaceText"", ""replace"": ""x"" } ] }", typeof(ArgumentException))]
    [InlineData(@"{ ""operations"": [ { ""type"": ""replaceText"", ""find"": ""x"" } ] }", typeof(ArgumentException))]
    [InlineData(@"{ ""operations"": [ { ""type"": ""insertAfter"", ""content"": { ""text"": ""x"" } } ] }", typeof(ArgumentException))]
    [InlineData(@"{ ""operations"": [ { ""type"": ""insertAfter"", ""target"": ""x"" } ] }", typeof(ArgumentException))]
    [InlineData(@"{ ""operations"": [ { ""type"": ""mystery"" } ] }", typeof(NotSupportedException))]
    public void JsonParser_WithInvalidInstructionShapes_ShouldThrowExpectedException(string json, Type expectedException)
    {
        var exception = Record.Exception(() => new DocxJsonInstructionParser().Parse(json));

        Assert.NotNull(exception);
        Assert.IsType(expectedException, exception);
    }

    [Theory]
    [InlineData("")]
    [InlineData("operations:")]
    public void YamlParser_WithMissingOperations_ShouldThrowArgumentException(string yaml)
    {
        var parser = new DocxYamlInstructionParser();

        var exception = Assert.Throws<ArgumentException>(() => parser.Parse(yaml));
        Assert.Contains("Invalid YAML", exception.Message);
    }

    [Fact]
    public void YamlParser_WithAllSupportedInstructionsAndOptionalFields_ShouldParse()
    {
        var yaml = """
        operations:
          - type: CREATE
          - type: addParagraph
            text: Styled text
            style: Heading1
          - type: replaceText
            find: old
            replace: new
          - type: insertAfter
            target: marker
            content:
              text: Inserted
              style: Body
        """;

        var instructions = new DocxYamlInstructionParser().Parse(yaml);

        Assert.Collection(instructions.Operations,
            first => Assert.IsType<CreateDocumentInstruction>(first),
            second =>
            {
                var add = Assert.IsType<AddParagraphInstruction>(second);
                Assert.Equal("Styled text", add.Text);
                Assert.Equal("Heading1", add.Style);
            },
            third =>
            {
                var replace = Assert.IsType<ReplaceTextInstruction>(third);
                Assert.Equal("old", replace.Find);
                Assert.Equal("new", replace.Replace);
            },
            fourth =>
            {
                var insert = Assert.IsType<InsertAfterInstruction>(fourth);
                Assert.Equal("marker", insert.Target);
                Assert.Equal("Inserted", insert.Content.Text);
            });
    }

    [Theory]
    [InlineData("operations:\n  - type:", typeof(NotSupportedException))]
    [InlineData("operations:\n  - type: unknown", typeof(NotSupportedException))]
    [InlineData("operations:\n  - type: addParagraph", typeof(ArgumentException))]
    [InlineData("operations:\n  - type: replaceText\n    replace: x", typeof(ArgumentException))]
    [InlineData("operations:\n  - type: replaceText\n    find: x", typeof(ArgumentException))]
    [InlineData("operations:\n  - type: insertAfter\n    content:\n      text: x", typeof(ArgumentException))]
    [InlineData("operations:\n  - type: insertAfter\n    target: x", typeof(ArgumentException))]
    public void YamlParser_WithInvalidInstructionShapes_ShouldThrowExpectedException(string yaml, Type expectedException)
    {
        var exception = Record.Exception(() => new DocxYamlInstructionParser().Parse(yaml));

        Assert.NotNull(exception);
        Assert.IsType(expectedException, exception);
    }

    [Theory]
    [InlineData("not: [valid", typeof(Exception))]
    [InlineData("operations: scalar", typeof(Exception))]
    public void YamlParser_WithMalformedYaml_ShouldThrow(string yaml, Type expectedException)
    {
        var exception = Record.Exception(() => new DocxYamlInstructionParser().Parse(yaml));

        Assert.NotNull(exception);
        Assert.IsAssignableFrom(expectedException, exception);
    }

    [Fact]
    public void JsonParser_WithAddRichContent_ShouldParseContentBlocks()
    {
        var json = """
        {
          "operations": [
            {
              "type": "addRichContent",
              "blocks": [
                { "type": "paragraph", "text": "Hello" },
                { "type": "heading", "level": 1, "text": "Title" },
                { "type": "list", "ordered": true, "items": ["One", "Two"] },
                { "type": "blockquote", "text": "Quote text" },
                { "type": "code", "text": "var x = 1;", "language": "csharp" },
                { "type": "horizontalRule" },
                { "type": "custom", "customType": "callout", "text": "Note", "style": "Note" }
              ]
            }
          ]
        }
        """;

        var instructions = new DocxJsonInstructionParser().Parse(json);

        Assert.Single(instructions.Operations);
        var rich = Assert.IsType<AddRichContentInstruction>(instructions.Operations[0]);
        Assert.Equal(7, rich.Blocks.Count);
        Assert.IsType<ParagraphBlock>(rich.Blocks[0]);
        Assert.IsType<HeadingBlock>(rich.Blocks[1]);
        Assert.IsType<ListBlock>(rich.Blocks[2]);
        Assert.IsType<BlockquoteBlock>(rich.Blocks[3]);
        Assert.IsType<CodeBlock>(rich.Blocks[4]);
        Assert.IsType<HorizontalRuleBlock>(rich.Blocks[5]);
        Assert.IsType<CustomBlock>(rich.Blocks[6]);
    }

    [Fact]
    public void JsonParser_WithReplaceWithRichContent_ShouldParse()
    {
        var json = """
        {
          "operations": [
            {
              "type": "replaceWithRichContent",
              "target": "find me",
              "blocks": [
                { "type": "paragraph", "text": "Replacement content" }
              ]
            }
          ]
        }
        """;

        var instructions = new DocxJsonInstructionParser().Parse(json);

        Assert.Single(instructions.Operations);
        var replace = Assert.IsType<ReplaceWithRichContentInstruction>(instructions.Operations[0]);
        Assert.Equal("find me", replace.Target);
        Assert.Single(replace.Blocks);
        Assert.IsType<ParagraphBlock>(replace.Blocks[0]);
    }

    [Fact]
    public void JsonParser_WithAddRichContentMissingBlocks_ShouldThrow()
    {
        var json = """
        { "operations": [ { "type": "addRichContent" } ] }
        """;

        Assert.Throws<ArgumentException>(() => new DocxJsonInstructionParser().Parse(json));
    }

    [Fact]
    public void JsonParser_WithTableBlock_ShouldParseCells()
    {
        var json = """
        {
          "operations": [
            {
              "type": "addRichContent",
              "blocks": [
                {
                  "type": "table",
                  "rows": [
                    { "cells": [ { "text": "A" }, { "text": "B" } ] },
                    { "cells": [ { "text": "C" }, { "text": "D" } ] }
                  ]
                }
              ]
            }
          ]
        }
        """;

        var instructions = new DocxJsonInstructionParser().Parse(json);

        var rich = Assert.IsType<AddRichContentInstruction>(instructions.Operations[0]);
        var table = Assert.IsType<TableBlock>(rich.Blocks[0]);
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(2, table.Rows[0].Cells.Count);
        Assert.Equal("A", table.Rows[0].Cells[0].Text);
        Assert.Equal("D", table.Rows[1].Cells[1].Text);
    }

    [Fact]
    public void JsonParser_WithUnknownBlockType_ShouldThrow()
    {
        var json = """
        { "operations": [ { "type": "addRichContent", "blocks": [ { "type": "image" } ] } ] }
        """;

        var ex = Assert.Throws<NotSupportedException>(() => new DocxJsonInstructionParser().Parse(json));
        Assert.Contains("image", ex.Message);
        Assert.Contains("Valid types", ex.Message);
    }

    [Fact]
    public void JsonParser_WithBlockMissingType_ShouldThrow()
    {
        var json = """
        { "operations": [ { "type": "addRichContent", "blocks": [ { "text": "no type" } ] } ] }
        """;

        Assert.Throws<ArgumentException>(() => new DocxJsonInstructionParser().Parse(json));
    }

    [Fact]
    public void YamlParser_WithAddRichContent_ShouldParseContentBlocks()
    {
        var yaml = """
        operations:
          - type: addRichContent
            blocks:
              - type: paragraph
                text: Hello YAML
              - type: heading
                level: 2
                text: YAML Heading
              - type: list
                ordered: false
                items:
                  - Apple
                  - Banana
              - type: blockquote
                text: Think different
              - type: code
                text: print("hello")
                language: python
              - type: horizontalRule
              - type: custom
                customType: tip
                text: Pro tip
                style: Tip
        """;

        var instructions = new DocxYamlInstructionParser().Parse(yaml);

        Assert.Single(instructions.Operations);
        var rich = Assert.IsType<AddRichContentInstruction>(instructions.Operations[0]);
        Assert.Equal(7, rich.Blocks.Count);
        Assert.IsType<ParagraphBlock>(rich.Blocks[0]);
        Assert.IsType<HeadingBlock>(rich.Blocks[1]);
        Assert.IsType<ListBlock>(rich.Blocks[2]);
        Assert.IsType<BlockquoteBlock>(rich.Blocks[3]);
        Assert.IsType<CodeBlock>(rich.Blocks[4]);
        Assert.IsType<HorizontalRuleBlock>(rich.Blocks[5]);
        Assert.IsType<CustomBlock>(rich.Blocks[6]);

        var list = Assert.IsType<ListBlock>(rich.Blocks[2]);
        Assert.Equal(["Apple", "Banana"], list.Items);
    }

    [Fact]
    public void YamlParser_WithReplaceWithRichContent_ShouldParse()
    {
        var yaml = """
        operations:
          - type: replaceWithRichContent
            target: old paragraph
            blocks:
              - type: paragraph
                text: New content
        """;

        var instructions = new DocxYamlInstructionParser().Parse(yaml);

        Assert.Single(instructions.Operations);
        var replace = Assert.IsType<ReplaceWithRichContentInstruction>(instructions.Operations[0]);
        Assert.Equal("old paragraph", replace.Target);
        Assert.Single(replace.Blocks);
    }

    [Fact]
    public void YamlParser_WithAddRichContentMissingBlocks_ShouldThrow()
    {
        var yaml = "operations:\n  - type: addRichContent";

        Assert.Throws<ArgumentException>(() => new DocxYamlInstructionParser().Parse(yaml));
    }

    [Fact]
    public void DocxInstructionValidator_WithValidJson_ShouldReturnNoErrors()
    {
        var json = """
        {
          "operations": [
            { "type": "create" },
            { "type": "addParagraph", "text": "Hello" },
            { "type": "replaceText", "find": "{{x}}", "replace": "y" },
            { "type": "insertAfter", "target": "t", "content": { "text": "ins" } }
          ]
        }
        """;

        var errors = new DocxInstructionValidator().Validate(json);

        Assert.Empty(errors);
    }

    [Fact]
    public void DocxInstructionValidator_WithUnknownOpType_ShouldReportError()
    {
        var json = """
        { "operations": [ { "type": "badOp" } ] }
        """;

        var errors = new DocxInstructionValidator().Validate(json);

        Assert.Single(errors);
        Assert.Contains("badOp", errors[0]);
        Assert.Contains("supported", errors[0]);
    }

    [Fact]
    public void DocxInstructionValidator_WithMissingRequiredFields_ShouldReportErrors()
    {
        var json = """
        {
          "operations": [
            { "type": "addParagraph" },
            { "type": "replaceText", "replace": "x" },
            { "type": "addRichContent" },
            { "type": "replaceWithRichContent", "target": "x" }
          ]
        }
        """;

        var errors = new DocxInstructionValidator().Validate(json);

        Assert.Equal(4, errors.Count);
        Assert.Contains(errors, e => e.Contains("text") && e.Contains("addParagraph"));
        Assert.Contains(errors, e => e.Contains("find") && e.Contains("replaceText"));
        Assert.Contains(errors, e => e.Contains("blocks") && e.Contains("addRichContent"));
        Assert.Contains(errors, e => e.Contains("blocks") && e.Contains("replaceWithRichContent"));
    }

    [Fact]
    public void DocxInstructionValidator_WithUnknownFields_ShouldReportErrors()
    {
        var json = """
        {
          "operations": [
            { "type": "addParagraph", "text": "ok", "weirdField": 42 }
          ]
        }
        """;

        var errors = new DocxInstructionValidator().Validate(json);

        Assert.Single(errors);
        Assert.Contains("weirdField", errors[0]);
    }

    [Fact]
    public void DocxInstructionValidator_WithUnknownBlockType_ShouldReportError()
    {
        var json = """
        {
          "operations": [
            {
              "type": "addRichContent",
              "blocks": [
                { "type": "paragraph", "text": "ok" },
                { "type": "unknownBlock" }
              ]
            }
          ]
        }
        """;

        var errors = new DocxInstructionValidator().Validate(json);

        Assert.Single(errors);
        Assert.Contains("unknownBlock", errors[0]);
        Assert.Contains("blocks[1]", errors[0]);
    }

    [Fact]
    public void DocxInstructionValidator_WithInvalidJson_ShouldReportParseError()
    {
        var errors = new DocxInstructionValidator().Validate("not json");

        Assert.Single(errors);
        Assert.Contains("Invalid JSON", errors[0]);
    }

    [Fact]
    public void DocxInstructionValidator_WithMissingOperations_ShouldReportError()
    {
        var errors = new DocxInstructionValidator().Validate("{}");

        Assert.Single(errors);
        Assert.Contains("operations", errors[0]);
    }

    [Theory]
    [InlineData(@"{ ""operations"": [ { ""type"": ""replaceText"", ""find"": """", ""replace"": ""x"" } ] }")]
    public void JsonParser_WithEmptyFind_ShouldThrowArgumentException(string json)
    {
        // Empty find would insert the replacement between every character of the document.
        var ex = Assert.Throws<ArgumentException>(() => new DocxJsonInstructionParser().Parse(json));
        Assert.Contains("Find", ex.Message);
        Assert.Contains("non-empty", ex.Message);
    }

    [Fact]
    public void YamlParser_WithEmptyFind_ShouldThrowArgumentException()
    {
        var yaml = "operations:\n  - type: replaceText\n    find: ''\n    replace: x";

        var ex = Assert.Throws<ArgumentException>(() => new DocxYamlInstructionParser().Parse(yaml));
        Assert.Contains("non-empty", ex.Message);
    }

    [Fact]
    public void YamlParser_WithInsertAfterMissingContentText_ShouldThrowArgumentException()
    {
        // YamlDotNet ignores C# 'required', so content without text must be rejected explicitly.
        var yaml = "operations:\n  - type: insertAfter\n    target: t\n    content:\n      style: Body";

        var ex = Assert.Throws<ArgumentException>(() => new DocxYamlInstructionParser().Parse(yaml));
        Assert.Contains("text", ex.Message);
    }

    [Fact]
    public void Parsers_HeadingWithoutLevel_ShouldDefaultToLevelOne()
    {
        var json = """
        { "operations": [ { "type": "addRichContent", "blocks": [ { "type": "heading", "text": "J" } ] } ] }
        """;
        var yaml = """
        operations:
          - type: addRichContent
            blocks:
              - type: heading
                text: Y
        """;

        var jsonHeading = Assert.IsType<HeadingBlock>(
            Assert.IsType<AddRichContentInstruction>(new DocxJsonInstructionParser().Parse(json).Operations[0]).Blocks[0]);
        var yamlHeading = Assert.IsType<HeadingBlock>(
            Assert.IsType<AddRichContentInstruction>(new DocxYamlInstructionParser().Parse(yaml).Operations[0]).Blocks[0]);

        Assert.Equal(1, jsonHeading.Level);
        Assert.Equal(1, yamlHeading.Level);
    }

    [Theory]
    // level: 1.5 — raw GetInt32 would throw a framework FormatException
    [InlineData(@"{ ""operations"": [ { ""type"": ""addRichContent"", ""blocks"": [ { ""type"": ""heading"", ""level"": 1.5, ""text"": ""x"" } ] } ] }", "level")]
    // items: [1] — raw GetString would throw InvalidOperationException
    [InlineData(@"{ ""operations"": [ { ""type"": ""addRichContent"", ""blocks"": [ { ""type"": ""list"", ""items"": [1] } ] } ] }", "items[0]")]
    // rows: "x" — raw EnumerateArray would throw InvalidOperationException
    [InlineData(@"{ ""operations"": [ { ""type"": ""addRichContent"", ""blocks"": [ { ""type"": ""table"", ""rows"": ""x"" } ] } ] }", "rows")]
    public void JsonParser_WithMalformedBlockValueKinds_ShouldThrowDescriptiveArgumentException(string json, string fieldHint)
    {
        var ex = Assert.Throws<ArgumentException>(() => new DocxJsonInstructionParser().Parse(json));
        Assert.Contains(fieldHint, ex.Message);
        Assert.Contains("blocks[0]", ex.Message);
    }

    [Theory]
    [InlineData("operations:\n  - type: addRichContent\n    blocks:\n      - type: heading\n        level: abc\n        text: x", "level")]
    [InlineData("operations:\n  - type: addRichContent\n    blocks:\n      - type: heading\n        level: 1.5\n        text: x", "level")]
    [InlineData("operations:\n  - type: addRichContent\n    blocks:\n      - type: table\n        rows: notalist", "rows")]
    [InlineData("operations:\n  - type: addRichContent\n    blocks:\n      - type: list\n        ordered: maybe", "ordered")]
    public void YamlParser_WithMalformedBlockValueKinds_ShouldThrowDescriptiveArgumentException(string yaml, string fieldHint)
    {
        var ex = Assert.Throws<ArgumentException>(() => new DocxYamlInstructionParser().Parse(yaml));
        Assert.Contains(fieldHint, ex.Message);
    }

    [Fact]
    public void DocxInstructionValidator_WithEmptyFind_ShouldReportError()
    {
        var json = """
        { "operations": [ { "type": "replaceText", "find": "", "replace": "x" } ] }
        """;

        var errors = new DocxInstructionValidator().Validate(json);

        Assert.Single(errors);
        Assert.Contains("find", errors[0]);
        Assert.Contains("empty", errors[0]);
    }

    [Fact]
    public void DocxInstructionValidator_HeadingLevelOptionalButMustBeInteger()
    {
        var missingLevel = """
        { "operations": [ { "type": "addRichContent", "blocks": [ { "type": "heading", "text": "H" } ] } ] }
        """;
        var fractionalLevel = """
        { "operations": [ { "type": "addRichContent", "blocks": [ { "type": "heading", "level": 1.5, "text": "H" } ] } ] }
        """;

        Assert.Empty(new DocxInstructionValidator().Validate(missingLevel));

        var errors = new DocxInstructionValidator().Validate(fractionalLevel);
        Assert.Single(errors);
        Assert.Contains("level", errors[0]);
    }

    [Fact]
    public void DocxInstructionValidator_ShouldReturnIndependentErrorLists()
    {
        var validator = new DocxInstructionValidator();

        var first = validator.Validate("{}");
        var second = validator.Validate("{ \"operations\": [] }");

        Assert.Single(first);
        Assert.Empty(second);
    }

    [Fact]
    public void DocxInstructionValidator_WithRichContentAndValidBlocks_ShouldReturnNoErrors()
    {
        var json = """
        {
          "operations": [
            {
              "type": "addRichContent",
              "blocks": [
                { "type": "paragraph", "text": "P1" },
                { "type": "heading", "level": 1, "text": "H1" },
                { "type": "list", "ordered": false, "items": ["a", "b"] },
                { "type": "table", "rows": [ { "cells": [ { "text": "c1" } ] } ] },
                { "type": "blockquote", "text": "Q" },
                { "type": "code", "text": "x" },
                { "type": "horizontalRule" },
                { "type": "custom", "customType": "tip", "text": "T" }
              ]
            }
          ]
        }
        """;

        var errors = new DocxInstructionValidator().Validate(json);

        Assert.Empty(errors);
    }
}
