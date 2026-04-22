using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using OfficeEditor.Core.Models;

namespace PptxEditor.Core.Variables;

public class PptxVariableDetector
{
    private static readonly Regex VariablePattern = new(
        @"\{\{([^}|]+)(?:\|([^}]*))?\}\}",
        RegexOptions.Compiled);

    public List<VariableInfo> Scan(PresentationDocument document)
    {
        var variables = new List<VariableInfo>();
        var presentation = document.PresentationPart!.Presentation;
        var slideIdList = presentation.SlideIdList;
        
        if (slideIdList == null) return variables;

        int slideIndex = 1;
        foreach (var slideId in slideIdList.ChildElements.OfType<SlideId>())
        {
            var slidePart = (SlidePart)document.PresentationPart.GetPartById(slideId.RelationshipId!);
            var slide = slidePart.Slide;
            
            if (slide?.CommonSlideData?.ShapeTree != null)
            {
                foreach (var shape in slide.CommonSlideData.ShapeTree.ChildElements)
                {
                    if (shape is P.Shape textShape)
                    {
                        var shapeName = textShape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Unknown";
                        var text = GetShapeText(textShape);
                        var location = $"slide:{slideIndex}:shape:{shapeName}";
                        
                        var matches = VariablePattern.Matches(text);
                        foreach (Match match in matches)
                        {
                            var variableName = match.Groups[1].Value.Trim();
                            var defaultValue = match.Groups[2].Success ? match.Groups[2].Value : null;
                            
                            if (!variables.Any(v => v.Name == variableName && v.Location == location))
                            {
                                variables.Add(new VariableInfo
                                {
                                    Name = variableName,
                                    FullMatch = match.Value,
                                    Location = location,
                                    DefaultValue = defaultValue
                                });
                            }
                        }
                    }
                }
            }
            
            slideIndex++;
        }

        return variables;
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
}
