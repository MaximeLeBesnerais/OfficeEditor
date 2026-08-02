using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Flow;

/// <summary>
/// Emits single-level bullet/ordered lists. Every list allocates a fresh numbering instance
/// (see <see cref="NumberingAllocator"/>) so separate lists restart and never bind to the
/// template's numbering; ordered lists honor <see cref="ListBlock.StartIndex"/>.
/// </summary>
internal static class ListEmitter
{
    public static void EmitList(OoxmlEmitContext context, OpenXmlCompositeElement container, ListBlock list, string path)
    {
        var ordered = list.Kind == ListKind.Ordered;
        var start = ordered ? list.StartIndex ?? 1 : 1;
        var numberingId = context.AllocateNumbering(ordered, start);

        for (var i = 0; i < list.Items.Count; i++)
        {
            var item = list.Items[i];
            var paragraph = ParagraphEmitter.BuildParagraph(context, item, list.Style, null, $"{path}.items[{i}]");
            paragraph.ParagraphProperties ??= new ParagraphProperties();
            paragraph.ParagraphProperties.NumberingProperties = new NumberingProperties(
                new NumberingLevelReference { Val = 0 },
                new NumberingId { Val = numberingId });
            container.Append(paragraph);
        }
    }
}
