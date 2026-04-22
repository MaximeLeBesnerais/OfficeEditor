using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PptxEditor.Core.Models;

namespace PptxEditor.Core.Services;

public class PptxAnatomizer
{
    public List<SlideAnatomy> Analyze(PresentationDocument document)
    {
        var anatomy = new List<SlideAnatomy>();
        var presentation = document.PresentationPart!.Presentation;
        var slideIdList = presentation.SlideIdList;
        
        if (slideIdList == null) return anatomy;

        int slideIndex = 1;
        foreach (var slideId in slideIdList.ChildElements.OfType<SlideId>())
        {
            var slidePart = (SlidePart)document.PresentationPart.GetPartById(slideId.RelationshipId!);
            var slide = slidePart.Slide;
            
            var slideAnatomy = new SlideAnatomy { SlideIndex = slideIndex };
            
            if (slide?.CommonSlideData?.ShapeTree != null)
            {
                foreach (var element in slide.CommonSlideData.ShapeTree.ChildElements)
                {
                    var slideElement = AnalyzeElement(element, slideIndex);
                    if (slideElement != null)
                    {
                        slideAnatomy.Elements.Add(slideElement);
                    }
                }
            }
            
            anatomy.Add(slideAnatomy);
            slideIndex++;
        }

        return anatomy;
    }

    private SlideElement? AnalyzeElement(OpenXmlElement element, int slideIndex)
    {
        return element switch
        {
            P.Shape shape => AnalyzeShape(shape, slideIndex),
            P.Picture picture => AnalyzePicture(picture, slideIndex),
            P.GraphicFrame graphicFrame => AnalyzeGraphicFrame(graphicFrame, slideIndex),
            P.GroupShape groupShape => AnalyzeGroupShape(groupShape, slideIndex),
            _ => null
        };
    }

    private SlideElement AnalyzeShape(P.Shape shape, int slideIndex)
    {
        var (id, name) = GetElementIdAndName(shape.NonVisualShapeProperties);
        var text = GetShapeText(shape);
        
        return new SlideElement
        {
            Type = "Text",
            Id = id,
            Name = name,
            Location = $"slide:{slideIndex}:shape:{name}",
            Text = text
        };
    }

    private SlideElement AnalyzePicture(P.Picture picture, int slideIndex)
    {
        var (id, name) = GetElementIdAndName(picture.NonVisualPictureProperties);
        
        return new SlideElement
        {
            Type = "Image",
            Id = id,
            Name = name,
            Location = $"slide:{slideIndex}:image:{name}"
        };
    }

    private SlideElement? AnalyzeGraphicFrame(P.GraphicFrame graphicFrame, int slideIndex)
    {
        var (id, name) = GetElementIdAndName(graphicFrame.NonVisualGraphicFrameProperties);
        
        var graphicData = graphicFrame.Graphic?.GraphicData;
        if (graphicData == null) return null;

        // Check if it's a table
        var table = graphicData.Elements<Drawing.Table>().FirstOrDefault();
        if (table != null)
        {
            var tableData = ExtractTableData(table);
            return new SlideElement
            {
                Type = "Table",
                Id = id,
                Name = name,
                Location = $"slide:{slideIndex}:table:{name}",
                TableData = tableData
            };
        }

        // For now, treat other graphic frames (charts, etc.) as generic elements
        return new SlideElement
        {
            Type = "GraphicFrame",
            Id = id,
            Name = name,
            Location = $"slide:{slideIndex}:graphic:{name}"
        };
    }

    private SlideElement? AnalyzeGroupShape(P.GroupShape groupShape, int slideIndex)
    {
        var (id, name) = GetElementIdAndName(groupShape.NonVisualGroupShapeProperties);
        
        // For groups, we list them but don't recurse into children for now
        // Children will be detected at the top level if the group is unwrapped
        return new SlideElement
        {
            Type = "Group",
            Id = id,
            Name = name,
            Location = $"slide:{slideIndex}:group:{name}"
        };
    }

    private (uint Id, string Name) GetElementIdAndName(OpenXmlElement? nvProperties)
    {
        if (nvProperties == null) return (0, "Unknown");

        // The cNvPr element might be parsed as OpenXmlUnknownElement due to namespace issues
        // in OpenXML SDK v3.5.1 when Drawing elements are nested inside Presentation elements
        var cnvPr = nvProperties.ChildElements.FirstOrDefault(e => e.LocalName == "cNvPr");
        if (cnvPr == null) return (0, "Unknown");

        uint id = 0;
        var idAttr = cnvPr.GetAttribute("id", "");
        if (!string.IsNullOrEmpty(idAttr.Value) && uint.TryParse(idAttr.Value, out var parsedId))
        {
            id = parsedId;
        }

        var nameAttr = cnvPr.GetAttribute("name", "");
        var name = !string.IsNullOrEmpty(nameAttr.Value) ? nameAttr.Value : "Unknown";

        return (id, name);
    }

    private string GetShapeText(P.Shape shape)
    {
        var textBody = shape.TextBody;
        if (textBody == null) return string.Empty;

        var text = new System.Text.StringBuilder();
        foreach (var paragraph in textBody.Elements<Drawing.Paragraph>())
        {
            foreach (var run in paragraph.Elements<Drawing.Run>())
            {
                if (run.Text?.Text != null)
                {
                    text.Append(run.Text.Text);
                }
            }
            text.Append(" ");
        }

        return text.ToString().Trim();
    }

    private List<List<string>> ExtractTableData(Drawing.Table table)
    {
        var data = new List<List<string>>();
        
        foreach (var row in table.Elements<Drawing.TableRow>())
        {
            var rowData = new List<string>();
            foreach (var cell in row.Elements<Drawing.TableCell>())
            {
                var cellText = new System.Text.StringBuilder();
                var textBody = cell.TextBody;
                if (textBody != null)
                {
                    foreach (var paragraph in textBody.Elements<Drawing.Paragraph>())
                    {
                        foreach (var run in paragraph.Elements<Drawing.Run>())
                        {
                            if (run.Text?.Text != null)
                            {
                                cellText.Append(run.Text.Text);
                            }
                        }
                    }
                }
                rowData.Add(cellText.ToString().Trim());
            }
            data.Add(rowData);
        }
        
        return data;
    }
}
