using DocxEditor.Core.Serialization;
using DocxEditor.Core.Models;
using OfficeEditor.Core.Exceptions;

namespace DocxEditor.Tests.Unit;

public class SerializationTests
{
    [Fact]
    public void JsonParser_ShouldParseValidJson()
    {
        // Arrange
        var json = @"{
            ""operations"": [
                {
                    ""type"": ""addParagraph"",
                    ""text"": ""Hello World"",
                    ""style"": ""Heading1""
                },
                {
                    ""type"": ""replaceText"",
                    ""find"": ""{{NAME}}"",
                    ""replace"": ""John""
                }
            ]
        }";

        var parser = new DocxJsonInstructionParser();

        // Act
        var instructions = parser.Parse(json);

        // Assert
        Assert.Equal(2, instructions.Operations.Count);
        Assert.IsType<AddParagraphInstruction>(instructions.Operations[0]);
        Assert.IsType<ReplaceTextInstruction>(instructions.Operations[1]);
        
        var addParagraph = (AddParagraphInstruction)instructions.Operations[0];
        Assert.Equal("Hello World", addParagraph.Text);
        Assert.Equal("Heading1", addParagraph.Style);
        
        var replaceText = (ReplaceTextInstruction)instructions.Operations[1];
        Assert.Equal("{{NAME}}", replaceText.Find);
        Assert.Equal("John", replaceText.Replace);
    }

    [Fact]
    public void YamlParser_ShouldParseValidYaml()
    {
        // Arrange
        var yaml = @"
operations:
  - type: addParagraph
    text: Hello World
    style: Heading1
  - type: replaceText
    find: '{{NAME}}'
    replace: John
";

        var parser = new DocxYamlInstructionParser();

        // Act
        var instructions = parser.Parse(yaml);

        // Assert
        Assert.Equal(2, instructions.Operations.Count);
        Assert.IsType<AddParagraphInstruction>(instructions.Operations[0]);
        Assert.IsType<ReplaceTextInstruction>(instructions.Operations[1]);
    }

    [Fact]
    public void JsonParser_ShouldThrowOnInvalidJson()
    {
        // Malformed JSON is normalized to the domain exception with the original
        // System.Text.Json failure preserved as the inner exception — never a raw
        // JsonException leaking to callers.
        var json = "invalid json";
        var parser = new DocxJsonInstructionParser();

        var ex = Assert.Throws<OfficeEditorException>(() => parser.Parse(json));
        Assert.Contains("Invalid JSON", ex.Message);
        Assert.IsType<System.Text.Json.JsonException>(ex.InnerException);
    }

    [Fact]
    public void YamlParser_ShouldThrowOnMalformedYaml()
    {
        // Malformed YAML is normalized to the domain exception with the original
        // YamlDotNet failure preserved as the inner exception.
        var yaml = "not: [valid";
        var parser = new DocxYamlInstructionParser();

        var ex = Assert.Throws<OfficeEditorException>(() => parser.Parse(yaml));
        Assert.Contains("Invalid YAML", ex.Message);
        Assert.IsType<YamlDotNet.Core.YamlException>(ex.InnerException);
    }

    [Fact]
    public void JsonParser_ShouldThrowOnUnsupportedType()
    {
        // Arrange
        var json = @"{
            ""operations"": [
                {
                    ""type"": ""unsupportedType""
                }
            ]
        }";

        var parser = new DocxJsonInstructionParser();

        // Act & Assert
        Assert.Throws<NotSupportedException>(() => parser.Parse(json));
    }
}
