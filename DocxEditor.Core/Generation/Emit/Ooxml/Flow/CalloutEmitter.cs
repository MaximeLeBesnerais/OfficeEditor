using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Flow;

/// <summary>
/// Emits flow callouts as a shaded single-cell table with a thick tone-accented left bar.
/// The tone drives the fill/accent from a fixed fallback palette (the vocabulary has no
/// callout color tokens), so callouts always render with a clear visual treatment.
/// </summary>
internal static class CalloutEmitter
{
    public static void EmitCallout(OoxmlEmitContext context, OpenXmlCompositeElement container, CalloutBlock callout, string path)
    {
        var (fill, accent) = ToneColors(callout.Tone);

        var tableProperties = new TableProperties(
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4, Color = "D9D9D9" },
                new LeftBorder { Val = BorderValues.Single, Size = 24, Color = accent },
                new BottomBorder { Val = BorderValues.Single, Size = 4, Color = "D9D9D9" },
                new RightBorder { Val = BorderValues.Single, Size = 4, Color = "D9D9D9" }))
        {
            TableWidth = new TableWidth { Width = "auto", Type = TableWidthUnitValues.Auto }
        };

        var table = new Table(tableProperties, new TableGrid(new GridColumn()));
        var cellProperties = new TableCellProperties(
            new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = fill });
        var paragraph = ParagraphEmitter.BuildParagraph(context, callout.Content, callout.Style, null, path);
        var cell = new DocumentFormat.OpenXml.Wordprocessing.TableCell(cellProperties, paragraph);
        var row = new DocumentFormat.OpenXml.Wordprocessing.TableRow(cell);
        table.Append(row);

        container.Append(table);
    }

    private static (string Fill, string Accent) ToneColors(CalloutTone tone) => tone switch
    {
        CalloutTone.Tip => ("E8F5E9", "2E7D32"),
        CalloutTone.Warning => ("FFF8E1", "F9A825"),
        CalloutTone.Error => ("FDECEA", "C62828"),
        _ => ("F2F5FA", "1F4E79")
    };
}
