namespace DocxEditor.Core.Models;

public abstract record Instruction
{
    public string Type { get; init; } = string.Empty;
}

public record CreateDocumentInstruction : Instruction
{
    public CreateDocumentInstruction()
    {
        Type = "create";
    }
}

public record AddParagraphInstruction : Instruction
{
    public AddParagraphInstruction()
    {
        Type = "addParagraph";
    }

    public required string Text { get; init; }
    public string? Style { get; init; }
}

public record AddRichContentInstruction : Instruction
{
    public AddRichContentInstruction()
    {
        Type = "addRichContent";
    }

    public required List<ContentBlock> Blocks { get; init; }
}

public record ReplaceWithRichContentInstruction : Instruction
{
    public ReplaceWithRichContentInstruction()
    {
        Type = "replaceWithRichContent";
    }

    public required string Target { get; init; }
    public required List<ContentBlock> Blocks { get; init; }
}

public record ReplaceTextInstruction : Instruction
{
    public ReplaceTextInstruction()
    {
        Type = "replaceText";
    }

    public required string Find { get; init; }
    public required string Replace { get; init; }
}

public record InsertAfterInstruction : Instruction
{
    public InsertAfterInstruction()
    {
        Type = "insertAfter";
    }

    public required string Target { get; init; }
    public required ParagraphContent Content { get; init; }
}

public record ParagraphContent
{
    public required string Text { get; init; }
    public string? Style { get; init; }
}

public record DocumentInstructions
{
    public required List<Instruction> Operations { get; init; }
}
