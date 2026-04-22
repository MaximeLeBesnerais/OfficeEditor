using DocxEditor.Core.Serialization;
using DocxEditor.Core.Models;

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

        var parser = new JsonInstructionParser();

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

        var parser = new YamlInstructionParser();

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
        // Arrange
        var json = "invalid json";
        var parser = new JsonInstructionParser();

        // Act & Assert
        Assert.ThrowsAny<Exception>(() => parser.Parse(json));
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

        var parser = new JsonInstructionParser();

        // Act & Assert
        Assert.Throws<NotSupportedException>(() => parser.Parse(json));
    }
}
