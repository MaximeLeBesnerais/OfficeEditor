using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxEditor.Core.Builders;

/// <summary>
/// Builds the OpenXML notes-slide payload for a slide: a <see cref="NotesSlide"/> whose
/// shape tree holds one body-placeholder shape (the same shape PowerPoint persists for
/// typed notes) carrying the notes text, one paragraph per line. Shared by the OOXML
/// generation emitter and the edit-path <see cref="ISlideBuilder.SetNotes"/> so both
/// write byte-identical notes parts.
/// </summary>
internal static class NotesSlideWriter
{
    /// <summary>Creates a <see cref="NotesSlide"/> carrying <paramref name="notesText"/>.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="notesText"/> is null.</exception>
    public static NotesSlide Create(string notesText)
    {
        ArgumentNullException.ThrowIfNull(notesText);

        // One paragraph per line (a:br inside a run would carry the same layout but
        // PowerPoint writes multi-line notes as separate paragraphs).
        var paragraphs = notesText
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Split('\n')
            .Select(LineParagraph)
            .ToList();

        var textBody = new P.TextBody(new Drawing.BodyProperties(), new Drawing.ListStyle());
        foreach (var paragraph in paragraphs)
        {
            textBody.Append(paragraph);
        }

        var shape = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1, Name = "Notes" },
                new P.NonVisualShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties(
                    new PlaceholderShape { Type = PlaceholderValues.Body, Index = 1 })),
            new ShapeProperties(),
            textBody);

        // Notes slide scaffolding mirrors the generation emitter's empty shape tree:
        // NonVisualGroupShapeProperties + GroupShapeProperties then the notes shape.
        var shapeTree = new ShapeTree(
            new NonVisualGroupShapeProperties(
                new NonVisualDrawingProperties { Id = 0, Name = "" },
                new NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(
                new Drawing.TransformGroup(
                    new Drawing.Offset { X = 0, Y = 0 },
                    new Drawing.Extents { Cx = 0, Cy = 0 },
                    new Drawing.ChildOffset { X = 0, Y = 0 },
                    new Drawing.ChildExtents { Cx = 0, Cy = 0 })),
            shape);

        return new NotesSlide(new CommonSlideData(shapeTree));
    }

    private static Drawing.Paragraph LineParagraph(string line)
    {
        var text = new Drawing.Text(line);
        if (line != line.Trim())
        {
            text.SetAttribute(new OpenXmlAttribute("xml:space", "http://www.w3.org/XML/1998/namespace", "preserve"));
        }
        return new Drawing.Paragraph(new Drawing.Run(text));
    }
}
