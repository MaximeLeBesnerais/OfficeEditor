using DocxEditor.Core.Models;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace DocxEditor.Core.Serialization;

public class YamlInstructionParser
{
    private readonly IDeserializer _deserializer;

    public YamlInstructionParser()
    {
        _deserializer = new DeserializerBuilder()
            .WithNamingConvention(CamelCaseNamingConvention.Instance)
            .Build();
    }

    public DocumentInstructions Parse(string yaml)
    {
        var wrapper = _deserializer.Deserialize<YamlInstructionWrapper>(yaml);
        if (wrapper?.Operations == null)
        {
            throw new ArgumentException("Invalid YAML instruction file.");
        }

        var instructions = new List<Instruction>();
        foreach (var op in wrapper.Operations)
        {
            instructions.Add(ParseInstruction(op));
        }

        return new DocumentInstructions { Operations = instructions };
    }

    private Instruction ParseInstruction(YamlInstructionDto dto)
    {
        return dto.Type?.ToLowerInvariant() switch
        {
            "create" => new CreateDocumentInstruction(),
            "addparagraph" => new AddParagraphInstruction
            {
                Text = dto.Text ?? throw new ArgumentException("Text is required for addParagraph."),
                Style = dto.Style
            },
            "replacetext" => new ReplaceTextInstruction
            {
                Find = dto.Find ?? throw new ArgumentException("Find is required for replaceText."),
                Replace = dto.Replace ?? throw new ArgumentException("Replace is required for replaceText.")
            },
            "insertafter" => new InsertAfterInstruction
            {
                Target = dto.Target ?? throw new ArgumentException("Target is required for insertAfter."),
                Content = dto.Content ?? throw new ArgumentException("Content is required for insertAfter.")
            },
            _ => throw new NotSupportedException($"Instruction type '{dto.Type}' is not supported.")
        };
    }

    private class YamlInstructionWrapper
    {
        public List<YamlInstructionDto>? Operations { get; set; }
    }

    private class YamlInstructionDto
    {
        public string? Type { get; set; }
        public string? Text { get; set; }
        public string? Style { get; set; }
        public string? Find { get; set; }
        public string? Replace { get; set; }
        public string? Target { get; set; }
        public ParagraphContent? Content { get; set; }
    }
}
