using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Generation.Design;
using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Flow;

/// <summary>
/// Dispatches flow blocks into a container (body, header, footer). Blocks are resolved
/// positionally via <c>is</c> pattern matching; the flow <c>group</c> container is flattened
/// into the parent flow per the vocabulary's authoring-sugar contract. An optional resolved
/// page format is forwarded so blocks (tables) can measure guardrails against the real text width.
/// </summary>
internal static class FlowBlockEmitter
{
    public static void EmitBlocks(
        OoxmlEmitContext context,
        OpenXmlCompositeElement container,
        IReadOnlyList<FlowBlock> blocks,
        string path,
        ResolvedPageFormat? pageFormat = null)
    {
        for (var i = 0; i < blocks.Count; i++)
        {
            EmitBlock(context, container, blocks[i], $"{path}[{i}]", pageFormat);
        }
    }

    private static void EmitBlock(
        OoxmlEmitContext context,
        OpenXmlCompositeElement container,
        FlowBlock block,
        string path,
        ResolvedPageFormat? pageFormat)
    {
        switch (block)
        {
            case ParagraphBlock paragraph:
                ParagraphEmitter.EmitParagraph(context, container, paragraph.Content, paragraph.Style, null, path);
                break;
            case HeadingBlock heading:
                ParagraphEmitter.EmitHeading(context, container, heading, path);
                break;
            case ListBlock list:
                ListEmitter.EmitList(context, container, list, path);
                break;
            case TableBlock table:
                TableEmitter.EmitTable(context, container, table, path, pageFormat);
                break;
            case ImageElement image:
                ImageEmitter.EmitInlineImage(context, container, image, path);
                break;
            case CalloutBlock callout:
                CalloutEmitter.EmitCallout(context, container, callout, path);
                break;
            case PageBreakBlock:
                EmitPageBreak(container);
                break;
            case FlowContainerBlock group:
                // Group is authoring sugar: flatten nested blocks into the parent flow.
                EmitBlocks(context, container, group.Blocks, $"{path}.blocks", pageFormat);
                break;
            default:
                context.Warn(path, $"unsupported flow block '{block.GetType().Name}'; the block was skipped.");
                break;
        }
    }

    /// <summary>Emits a hard page break as its own paragraph (<c>w:br w:type="page"</c>).</summary>
    public static void EmitPageBreak(OpenXmlCompositeElement container)
    {
        container.Append(new Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(new Break { Type = BreakValues.Page })));
    }
}
