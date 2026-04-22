using System.Text.Json;
using System.Text.Json.Serialization;
using DocxEditor.Core.Models;

namespace DocxEditor.Core.Serialization;

public class DocxJsonInstructionParser
{
    private readonly JsonSerializerOptions _options;

    public DocxJsonInstructionParser()
    {
        _options = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
        };
    }

    public DocumentInstructions Parse(string json)
    {
        var wrapper = JsonSerializer.Deserialize<JsonInstructionWrapper>(json, _options);
        if (wrapper?.Operations == null)
        {
            throw new ArgumentException("Invalid JSON instruction file.");
        }

        var instructions = new List<Instruction>();
        foreach (var op in wrapper.Operations)
        {
            instructions.Add(ParseInstruction(op));
        }

        return new DocumentInstructions { Operations = instructions };
    }

    private Instruction ParseInstruction(JsonInstructionDto dto)
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

    private class JsonInstructionWrapper
    {
        public List<JsonInstructionDto>? Operations { get; set; }
    }

    private class JsonInstructionDto
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
