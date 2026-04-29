using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxEditor.Core.Services;

public class PptxElementReplacer
{
    public void ReplaceText(SlidePart slidePart, uint elementId, string newText)
    {
        var slide = slidePart.Slide;
        if (slide?.CommonSlideData?.ShapeTree == null) return;

        var shape = FindShapeById(slide.CommonSlideData.ShapeTree, elementId);
        if (shape == null) return;

        var textBody = shape.TextBody;
        if (textBody == null) return;

        // Clear existing paragraphs and add new text
        textBody.RemoveAllChildren<Drawing.Paragraph>();
        textBody.Append(new Drawing.Paragraph(new Drawing.Run(new Drawing.Text(newText))));
    }

    public void ReplaceTableData(SlidePart slidePart, uint elementId, List<List<string>> newData)
    {
        var slide = slidePart.Slide;
        if (slide?.CommonSlideData?.ShapeTree == null) return;

        var graphicFrame = FindGraphicFrameById(slide.CommonSlideData.ShapeTree, elementId);
        if (graphicFrame == null) return;

        var graphicData = graphicFrame.Graphic?.GraphicData;
        if (graphicData == null) return;

        var existingTable = graphicData.Elements<Drawing.Table>().FirstOrDefault();
        if (existingTable == null) return;

        // Get column count from first row
        if (newData.Count == 0 || newData[0].Count == 0) return;
        
        var numCols = newData[0].Count;
        var numRows = newData.Count;

        // Create new table with updated data
        var newTable = new Drawing.Table(
            new Drawing.TableProperties { FirstRow = true },
            new Drawing.TableGrid(Enumerable.Range(0, numCols).Select(_ => new Drawing.GridColumn { Width = 7200000 / numCols }))
        );

        foreach (var row in newData)
        {
            var tableRow = new Drawing.TableRow { Height = 370840 };
            foreach (var cell in row)
            {
                var tableCell = new Drawing.TableCell(
                    new Drawing.TextBody(
                        new Drawing.BodyProperties(),
                        new Drawing.ListStyle(),
                        new Drawing.Paragraph(new Drawing.Run(new Drawing.Text(cell)))
                    ),
                    new Drawing.TableCellProperties()
                );
                tableRow.Append(tableCell);
            }
            newTable.Append(tableRow);
        }

        // Replace the old table
        graphicData.RemoveAllChildren<Drawing.Table>();
        graphicData.Append(newTable);
    }

    public void ReplaceImage(SlidePart slidePart, uint elementId, string newImagePath)
    {
        if (!File.Exists(newImagePath))
        {
            throw new FileNotFoundException($"Image not found: {newImagePath}");
        }

        var slide = slidePart.Slide;
        if (slide?.CommonSlideData?.ShapeTree == null) return;

        var picture = FindPictureById(slide.CommonSlideData.ShapeTree, elementId);
        if (picture == null) return;

        var blipFill = picture.BlipFill;
        if (blipFill == null) return;

        var blip = blipFill.Blip;
        if (blip == null) return;

        var oldEmbedId = blip.Embed?.Value;
        if (oldEmbedId == null) return;

        // Remove old image part
        var oldImagePart = (ImagePart?)slidePart.GetPartById(oldEmbedId);
        if (oldImagePart != null)
        {
            slidePart.DeletePart(oldImagePart);
        }

        // Add new image part
        var imagePart = slidePart.AddImagePart(ImagePartType.Jpeg);
        using (var stream = new FileStream(newImagePath, FileMode.Open))
        {
            imagePart.FeedData(stream);
        }

        // Update blip reference
        blip.Embed = slidePart.GetIdOfPart(imagePart);
    }

    private P.Shape? FindShapeById(ShapeTree shapeTree, uint elementId)
    {
        foreach (var element in shapeTree.ChildElements)
        {
            if (element is P.Shape shape)
            {
                var id = GetElementId(shape.NonVisualShapeProperties);
                if (id == elementId) return shape;
            }
        }
        return null;
    }

    private P.GraphicFrame? FindGraphicFrameById(ShapeTree shapeTree, uint elementId)
    {
        foreach (var element in shapeTree.ChildElements)
        {
            if (element is P.GraphicFrame graphicFrame)
            {
                var id = GetElementId(graphicFrame.NonVisualGraphicFrameProperties);
                if (id == elementId) return graphicFrame;
            }
        }
        return null;
    }

    private P.Picture? FindPictureById(ShapeTree shapeTree, uint elementId)
    {
        foreach (var element in shapeTree.ChildElements)
        {
            if (element is P.Picture picture)
            {
                var id = GetElementId(picture.NonVisualPictureProperties);
                if (id == elementId) return picture;
            }
        }
        return null;
    }

    private uint GetElementId(OpenXmlElement? nvProperties)
    {
        if (nvProperties == null) return 0;

        // The cNvPr element might be parsed as OpenXmlUnknownElement due to namespace issues
        // in OpenXML SDK v3.5.1 when Drawing elements are nested inside Presentation elements
        var cnvPr = nvProperties.ChildElements.FirstOrDefault(e => e.LocalName == "cNvPr");
        if (cnvPr == null) return 0;

        var idAttr = cnvPr.GetAttribute("id", "");
        if (!string.IsNullOrEmpty(idAttr.Value) && uint.TryParse(idAttr.Value, out var parsedId))
        {
            return parsedId;
        }

        return 0;
    }
}
