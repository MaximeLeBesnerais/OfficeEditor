using DocxEditor.Core.Builders;
using DocxEditor.Core.Models;

namespace DocxEditor.Core.Instructions;

public class InstructionEngine
{
    public void Execute(IDocumentBuilder builder, DocumentInstructions instructions)
    {
        foreach (var instruction in instructions.Operations)
        {
            ExecuteInstruction(builder, instruction);
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
            case CreateDocumentInstruction:
                // Create is handled by the builder initialization
                break;
            default:
                throw new NotSupportedException($"Instruction type '{instruction.Type}' is not supported.");
        }
    }
}
