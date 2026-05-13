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
}
