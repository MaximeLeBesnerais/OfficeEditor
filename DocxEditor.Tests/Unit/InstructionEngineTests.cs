using DocxEditor.Core.Builders;
using DocxEditor.Core.Instructions;
using DocxEditor.Core.Models;
using OfficeEditor.Core.Models;

namespace DocxEditor.Tests.Unit;

public class InstructionEngineTests
{
    [Fact]
    public void Execute_ShouldExecuteMultipleOperationsInOrder()
    {
        // Arrange
        var builder = new RecordingDocumentBuilder();
        var instructions = new DocumentInstructions
        {
            Operations =
            [
                new AddParagraphInstruction { Text = "First", Style = "Heading1" },
                new ReplaceTextInstruction { Find = "old", Replace = "new" },
                new InsertAfterInstruction
                {
                    Target = "First",
                    Content = new ParagraphContent { Text = "Second", Style = "Normal" }
                }
            ]
        };
        var engine = new InstructionEngine();

        // Act
        engine.Execute(builder, instructions);

        // Assert
        Assert.Equal(
        [
            "AddParagraph:First:Heading1",
            "ReplaceText:old:new",
            "InsertAfter:First:Second:Normal"
        ], builder.Calls);
    }

    [Fact]
    public void Execute_WithAddParagraphInstruction_ShouldCallAddParagraph()
    {
        // Arrange
        var builder = new RecordingDocumentBuilder();
        var instructions = new DocumentInstructions
        {
            Operations = [new AddParagraphInstruction { Text = "Paragraph text", Style = "BodyText" }]
        };
        var engine = new InstructionEngine();

        // Act
        engine.Execute(builder, instructions);

        // Assert
        Assert.Equal(["AddParagraph:Paragraph text:BodyText"], builder.Calls);
    }

    [Fact]
    public void Execute_WithReplaceTextInstruction_ShouldCallReplaceText()
    {
        // Arrange
        var builder = new RecordingDocumentBuilder();
        var instructions = new DocumentInstructions
        {
            Operations = [new ReplaceTextInstruction { Find = "{{Name}}", Replace = "Ada" }]
        };
        var engine = new InstructionEngine();

        // Act
        engine.Execute(builder, instructions);

        // Assert
        Assert.Equal(["ReplaceText:{{Name}}:Ada"], builder.Calls);
    }

    [Fact]
    public void Execute_WithInsertAfterInstruction_ShouldCallInsertAfter()
    {
        // Arrange
        var builder = new RecordingDocumentBuilder();
        var instructions = new DocumentInstructions
        {
            Operations =
            [
                new InsertAfterInstruction
                {
                    Target = "Target paragraph",
                    Content = new ParagraphContent { Text = "Inserted paragraph", Style = "Quote" }
                }
            ]
        };
        var engine = new InstructionEngine();

        // Act
        engine.Execute(builder, instructions);

        // Assert
        Assert.Equal(["InsertAfter:Target paragraph:Inserted paragraph:Quote"], builder.Calls);
    }

    [Fact]
    public void Execute_WithCreateDocumentInstruction_ShouldNotCallBuilder()
    {
        // Arrange
        var builder = new RecordingDocumentBuilder();
        var instructions = new DocumentInstructions
        {
            Operations = [new CreateDocumentInstruction()]
        };
        var engine = new InstructionEngine();

        // Act
        engine.Execute(builder, instructions);

        // Assert
        Assert.Empty(builder.Calls);
    }

    [Fact]
    public void Execute_WithUnsupportedInstruction_ShouldThrowNotSupportedExceptionIncludingInstructionType()
    {
        // Arrange
        var builder = new RecordingDocumentBuilder();
        var instructions = new DocumentInstructions
        {
            Operations = [new UnsupportedInstruction { Type = "unsupportedType" }]
        };
        var engine = new InstructionEngine();

        // Act
        var exception = Assert.Throws<NotSupportedException>(() => engine.Execute(builder, instructions));

        // Assert
        Assert.Contains("unsupportedType", exception.Message);
    }

    private sealed record UnsupportedInstruction : Instruction;

    private sealed class RecordingDocumentBuilder : IDocumentBuilder
    {
        public List<string> Calls { get; } = [];

        public IDocumentBuilder AddParagraph(string text, string? style = null)
        {
            Calls.Add($"AddParagraph:{text}:{style}");
            return this;
        }

        public IDocumentBuilder InsertAfter(string targetText, string text, string? style = null)
        {
            Calls.Add($"InsertAfter:{targetText}:{text}:{style}");
            return this;
        }

        public IDocumentBuilder ReplaceText(string find, string replace)
        {
            Calls.Add($"ReplaceText:{find}:{replace}");
            return this;
        }

        public IDocumentBuilder InsertBefore(string targetText, string text, string? style = null) => this;

        public IDocumentBuilder ReplaceParagraph(string targetText, string newText, string? style = null) => this;

        public IDocumentBuilder DeleteParagraph(string targetText) => this;

        public IDocumentBuilder ApplyStyle(string styleId) => this;

        public IDocumentBuilder AddRichContent(List<ContentBlock> blocks) => this;

        public IDocumentBuilder ReplaceWithRichContent(string targetText, List<ContentBlock> blocks) => this;

        public IDocumentBuilder AddMarkdown(string markdown, StyleMapping? styleMap = null) => this;

        public IDocumentBuilder ReplaceWithMarkdown(string targetText, string markdown, StyleMapping? styleMap = null) => this;

        public List<VariableInfo> DetectVariables() => [];

        public IDocumentBuilder MergeVariables(Dictionary<string, string> data) => this;

        public void MergeBatch(List<Dictionary<string, string>> records, string outputPattern, string? templatePath = null)
        {
        }

        public void Save(string? path = null)
        {
        }

        public void Save(Stream stream)
        {
        }

        public byte[] SaveToBytes() => [];

        public static IDocumentBuilder Create() => throw new NotImplementedException();

        public static IDocumentBuilder Open(Stream stream) => throw new NotImplementedException();

        public static IDocumentBuilder Open(byte[] bytes) => throw new NotImplementedException();

        public void Dispose()
        {
        }
    }
}
