using DocxEditor.Core.Builders;
using DocxEditor.Core.Models;

namespace DocxEditor.Core.Instructions;

public class InstructionEngine
{
    /// <summary>
    /// Executes the instruction set against the builder. All arguments are required, and every
    /// operation is prevalidated before the first mutation runs: a malformed programmatically
    /// constructed instruction (a null operation, a null rich-content block, or a null
    /// <see cref="InsertAfterInstruction.Content"/> — which would NRE in the dispatch switch)
    /// must never partially apply the earlier operations of a batch. Valid instructions parsed
    /// from JSON are always fully formed and pass unchanged; unsupported types still surface as
    /// <see cref="NotSupportedException"/> from dispatch.
    /// </summary>
    public void Execute(IDocumentBuilder builder, DocumentInstructions instructions)
    {
        ArgumentNullException.ThrowIfNull(builder);
        ArgumentNullException.ThrowIfNull(instructions);
        ArgumentNullException.ThrowIfNull(instructions.Operations);

        for (int i = 0; i < instructions.Operations.Count; i++)
        {
            ValidateOperation(instructions.Operations[i], i);
        }

        foreach (var instruction in instructions.Operations)
        {
            ExecuteInstruction(builder, instruction);
        }
    }

    private static void ValidateOperation(Instruction instruction, int index)
    {
        if (instruction is null)
        {
            throw new ArgumentException($"operations[{index}]: each operation must be non-null.", nameof(instruction));
        }

        switch (instruction)
        {
            case AddParagraphInstruction addParagraph:
                ArgumentNullException.ThrowIfNull(addParagraph.Text);
                break;
            case ReplaceTextInstruction replaceText:
                // Empty find corrupts the document (string.Replace("", x) inserts between
                // every character); mirror the DocumentBuilder.ReplaceText contract.
                if (string.IsNullOrEmpty(replaceText.Find))
                {
                    throw new ArgumentException("Find must be a non-empty string for replaceText.", nameof(replaceText.Find));
                }
                ArgumentNullException.ThrowIfNull(replaceText.Replace);
                break;
            case InsertAfterInstruction insertAfter:
                if (string.IsNullOrEmpty(insertAfter.Target))
                {
                    throw new ArgumentException("Target must be a non-empty string for insertAfter.", nameof(insertAfter.Target));
                }
                ArgumentNullException.ThrowIfNull(insertAfter.Content);
                ArgumentNullException.ThrowIfNull(insertAfter.Content.Text);
                break;
            case AddRichContentInstruction addRich:
                ValidateBlocks(addRich.Blocks, $"operations[{index}].blocks");
                break;
            case ReplaceWithRichContentInstruction replaceRich:
                if (string.IsNullOrEmpty(replaceRich.Target))
                {
                    throw new ArgumentException("Target must be a non-empty string for replaceWithRichContent.", nameof(replaceRich.Target));
                }
                ValidateBlocks(replaceRich.Blocks, $"operations[{index}].blocks");
                break;
            case CreateDocumentInstruction:
                break;
            default:
                // Unsupported types are reported by the dispatch switch so the error includes
                // the instruction type; nothing to prevalidate here.
                break;
        }
    }

    private static void ValidateBlocks(List<ContentBlock> blocks, string path)
    {
        ArgumentNullException.ThrowIfNull(blocks);
        for (int i = 0; i < blocks.Count; i++)
        {
            // A null entry would be silently skipped by the content renderer's type switch,
            // rendering fewer blocks than requested; reject it before any mutation.
            if (blocks[i] is null)
            {
                throw new ArgumentException($"{path}[{i}]: each content block must be non-null.", nameof(blocks));
            }
        }
    }

    private void ExecuteInstruction(IDocumentBuilder builder, Instruction instruction)
    {
        switch (instruction)
        {
            case AddParagraphInstruction addParagraph:
                builder.AddParagraph(addParagraph.Text, addParagraph.Style);
                break;
            case ReplaceTextInstruction replaceText:
                builder.ReplaceText(replaceText.Find, replaceText.Replace);
                break;
            case InsertAfterInstruction insertAfter:
                builder.InsertAfter(insertAfter.Target, insertAfter.Content.Text, insertAfter.Content.Style);
                break;
            case AddRichContentInstruction addRich:
                builder.AddRichContent(addRich.Blocks);
                break;
            case ReplaceWithRichContentInstruction replaceRich:
                builder.ReplaceWithRichContent(replaceRich.Target, replaceRich.Blocks);
                break;
            case CreateDocumentInstruction:
                // Create is handled by the builder initialization
                break;
            default:
                throw new NotSupportedException($"Instruction type '{instruction.Type}' is not supported.");
        }
    }
}
