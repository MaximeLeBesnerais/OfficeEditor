using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxEditor.Core.Variables;

public class PptxVariableReplacer
{
    private static readonly Regex VariablePattern = new(
        @"\{\{([^}|]+)(?:\|([^}]*))?\}\}",
        RegexOptions.Compiled);

    public void Replace(PresentationDocument document, Dictionary<string, string> data)
    {
        var presentation = document.PresentationPart!.Presentation!;
        var slideIdList = presentation.SlideIdList;
        
        if (slideIdList == null) return;

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
                        ReplaceInShape(textShape, data);
                    }
                }
            }
        }
    }

    private void ReplaceInShape(P.Shape shape, Dictionary<string, string> data)
    {
        var textBody = shape.TextBody;
        if (textBody == null) return;

        foreach (var paragraph in textBody.Elements<Drawing.Paragraph>())
        {
            foreach (var run in paragraph.Elements<Drawing.Run>())
            {
                if (run.Text?.Text != null && run.Text.Text.Contains("{{"))
                {
                    run.Text.Text = ReplaceVariablesInText(run.Text.Text, data);
                }
            }

            ReplaceVariablesAcrossRuns(paragraph, data);
        }
    }

    private void ReplaceVariablesAcrossRuns(Drawing.Paragraph paragraph, Dictionary<string, string> data)
    {
        var texts = paragraph.Elements<Drawing.Run>()
            .Where(r => r.Text?.Text != null)
            .Select(r => r.Text!)
            .ToList();
        if (texts.Count <= 1)
        {
            return;
        }

        var combinedText = string.Concat(texts.Select(t => t.Text));
        if (!combinedText.Contains("{{"))
        {
            return;
        }

        var replacedText = ReplaceVariablesInText(combinedText, data);
        if (replacedText == combinedText)
        {
            return;
        }

        texts[0].Text = replacedText;
        foreach (var text in texts.Skip(1))
        {
            text.Text = string.Empty;
        }
    }

    private string ReplaceVariablesInText(string text, Dictionary<string, string> data)
    {
        return VariablePattern.Replace(text, match =>
        {
            var variableName = match.Groups[1].Value.Trim();
            var defaultValue = match.Groups[2].Success ? match.Groups[2].Value : null;

            if (data.TryGetValue(variableName, out var value))
            {
                return value;
            }
            
            if (defaultValue != null)
            {
                return defaultValue;
            }
            
            return match.Value;
        });
    }
}
