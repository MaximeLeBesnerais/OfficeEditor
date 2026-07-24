using DocxEditor.Core.Builders;
using DocxEditor.Core.Instructions;
using DocxEditor.Core.Models;
using DocxEditor.Core.Serialization;
using OfficeEditor.Core.Models;

namespace DocxEditor.Tests.Unit;

public class InstructionEngineTests : IDisposable
{
    private readonly string _testOutputDir;
    private readonly List<string> _tempFiles = [];

    public InstructionEngineTests()
    {
        _testOutputDir = Path.Combine(Path.GetTempPath(), $"docx_instructions_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_testOutputDir);
    }

    public void Dispose()
    {
        foreach (var file in _tempFiles)
        {
            if (File.Exists(file)) File.Delete(file);
        }

        if (Directory.Exists(_testOutputDir))
        {
            try { Directory.Delete(_testOutputDir, true); } catch { }
        }
    }
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

    [Fact]
    public void SampleJson_ShouldParseAndExecuteEndToEnd()
    {
        var samplePath = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory, "..", "..", "..", "..",
            "examples", "Docx", "instructions", "sample.json"));
        Assert.True(File.Exists(samplePath), $"sample.json not found at {samplePath}");

        var json = File.ReadAllText(samplePath);
        var parser = new DocxJsonInstructionParser();
        var instructions = parser.Parse(json);

        Assert.Equal(9, instructions.Operations.Count);

        var outputPath = Path.Combine(_testOutputDir, "sample_output.docx");
        _tempFiles.Add(outputPath);

        {
            using var builder = (DocumentBuilder)DocumentBuilder.Create(outputPath);
            builder.AddParagraph("Company: {{companyName}} — Date: {{date}}", "Normal");

            var engine = new InstructionEngine();
            engine.Execute(builder, instructions);
            builder.Save();
        }

        using var opened = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(outputPath, false);
        var body = opened.MainDocumentPart!.Document!.Body!;
        var fullText = body.InnerText;

        Assert.Contains("Company: Acme Corporation", fullText);
        Assert.Contains("Date: 2024-01-15", fullText);
        Assert.Contains("This paragraph was added via JSON", fullText);
        Assert.Contains("This line was inserted after the first paragraph", fullText);
        Assert.Contains("Section Added by Instructions", fullText);
        Assert.Contains("First bullet point", fullText);
        Assert.Contains("Second bullet point", fullText);
        Assert.Contains("Third bullet point", fullText);
    }

    [Fact]
    public void Execute_WithAddRichContentInstruction_ShouldCallAddRichContent()
    {
        var builder = new RecordingDocumentBuilder();
        var blocks = new List<ContentBlock>
        {
            new ParagraphBlock { Text = "Hello" },
            new HeadingBlock { Level = 1, Text = "Title" }
        };
        var instructions = new DocumentInstructions
        {
            Operations = [new AddRichContentInstruction { Blocks = blocks }]
        };
        var engine = new InstructionEngine();

        engine.Execute(builder, instructions);

        Assert.Equal(["AddRichContent:2"], builder.Calls);
    }

    [Fact]
    public void Execute_WithReplaceWithRichContentInstruction_ShouldCallReplaceWithRichContent()
    {
        var builder = new RecordingDocumentBuilder();
        var blocks = new List<ContentBlock>
        {
            new ParagraphBlock { Text = "Replacement" }
        };
        var instructions = new DocumentInstructions
        {
            Operations = [new ReplaceWithRichContentInstruction { Target = "old text", Blocks = blocks }]
        };
        var engine = new InstructionEngine();

        engine.Execute(builder, instructions);

        Assert.Equal(["ReplaceWithRichContent:old text:1"], builder.Calls);
    }

    [Fact]
    public void Execute_WithAllSixOps_ShouldDispatchAll()
    {
        var builder = new RecordingDocumentBuilder();
        var instructions = new DocumentInstructions
        {
            Operations =
            [
                new AddParagraphInstruction { Text = "Para" },
                new ReplaceTextInstruction { Find = "{{x}}", Replace = "y" },
                new InsertAfterInstruction { Target = "Para", Content = new ParagraphContent { Text = "After" } },
                new AddRichContentInstruction { Blocks = [new ParagraphBlock { Text = "Rich" }] },
                new ReplaceWithRichContentInstruction { Target = "old", Blocks = [new ParagraphBlock { Text = "New" }] },
                new CreateDocumentInstruction()
            ]
        };
        var engine = new InstructionEngine();

        engine.Execute(builder, instructions);

        Assert.Equal(5, builder.Calls.Count);
        Assert.Equal("AddParagraph:Para:", builder.Calls[0]);
        Assert.Equal("ReplaceText:{{x}}:y", builder.Calls[1]);
        Assert.Equal("InsertAfter:Para:After:", builder.Calls[2]);
        Assert.Equal("AddRichContent:1", builder.Calls[3]);
        Assert.Equal("ReplaceWithRichContent:old:1", builder.Calls[4]);
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

        public IDocumentBuilder AddRichContent(List<ContentBlock> blocks)
        {
            Calls.Add($"AddRichContent:{blocks.Count}");
            return this;
        }

        public IDocumentBuilder ReplaceWithRichContent(string targetText, List<ContentBlock> blocks)
        {
            Calls.Add($"ReplaceWithRichContent:{targetText}:{blocks.Count}");
            return this;
        }

        public IDocumentBuilder AddHyperlink(string url, string displayText, string? style = null)
        {
            Calls.Add($"AddHyperlink:{url}:{displayText}:{style}");
            return this;
        }

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
