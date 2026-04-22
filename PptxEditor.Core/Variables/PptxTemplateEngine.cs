using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxEditor.Core.Variables;

public class PptxTemplateEngine
{
    private static readonly Regex IfPattern = new(
        @"\{\{#if\s+([^}]+)\}\}(.*?)\{\{/if\}\}",
        RegexOptions.Compiled | RegexOptions.Singleline);

    private static readonly Regex IfNotPattern = new(
        @"\{\{#ifnot\s+([^}]+)\}\}(.*?)\{\{/ifnot\}\}",
        RegexOptions.Compiled | RegexOptions.Singleline);

    private static readonly Regex EachPattern = new(
        @"\{\{#each\s+([^}]+)\}\}(.*?)\{\{/each\}\}",
        RegexOptions.Compiled | RegexOptions.Singleline);

    private static readonly Regex ComparisonPattern = new(
        @"^(\w+)\s*([\u003e\u003c=!]+)\s*(.+)$",
        RegexOptions.Compiled);

    public void Process(PresentationDocument document, Dictionary<string, object> data)
    {
        var presentation = document.PresentationPart!.Presentation;
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
                        ProcessShape(textShape, data);
                    }
                }
            }
        }
    }

    private void ProcessShape(P.Shape shape, Dictionary<string, object> data)
    {
        var textBody = shape.TextBody;
        if (textBody == null) return;

        foreach (var paragraph in textBody.Elements<Drawing.Paragraph>().ToList())
        {
            var text = GetParagraphText(paragraph);
            
            text = ProcessConditionals(text, data);
            text = ProcessLoops(text, data);
            
            if (text != GetParagraphText(paragraph))
            {
                paragraph.RemoveAllChildren<Drawing.Run>();
                paragraph.Append(new Drawing.Run(new Drawing.Text(text)));
            }
        }
    }

    private string GetParagraphText(Drawing.Paragraph paragraph)
    {
        var text = new System.Text.StringBuilder();
        foreach (var run in paragraph.Elements<Drawing.Run>())
        {
            if (run.Text?.Text != null)
            {
                text.Append(run.Text.Text);
            }
        }
        return text.ToString();
    }

    private string ProcessConditionals(string text, Dictionary<string, object> data)
    {
        text = IfPattern.Replace(text, match =>
        {
            var condition = match.Groups[1].Value.Trim();
            var content = match.Groups[2].Value;
            
            return EvaluateCondition(condition, data) ? content : string.Empty;
        });

        text = IfNotPattern.Replace(text, match =>
        {
            var condition = match.Groups[1].Value.Trim();
            var content = match.Groups[2].Value;
            
            return !EvaluateCondition(condition, data) ? content : string.Empty;
        });

        return text;
    }

    private string ProcessLoops(string text, Dictionary<string, object> data)
    {
        return EachPattern.Replace(text, match =>
        {
            var arrayName = match.Groups[1].Value.Trim();
            var template = match.Groups[2].Value;
            
            if (data.TryGetValue(arrayName, out var arrayValue) && arrayValue is List<Dictionary<string, object>> array)
            {
                var results = new List<string>();
                foreach (var item in array)
                {
                    var itemText = template;
                    foreach (var kvp in item)
                    {
                        itemText = itemText.Replace($"{{{kvp.Key}}}", kvp.Value?.ToString() ?? "");
                    }
                    results.Add(itemText);
                }
                return string.Join("\n", results);
            }
            
            return string.Empty;
        });
    }

    private bool EvaluateCondition(string condition, Dictionary<string, object> data)
    {
        var comparisonMatch = ComparisonPattern.Match(condition);
        if (comparisonMatch.Success)
        {
            var left = comparisonMatch.Groups[1].Value.Trim();
            var op = comparisonMatch.Groups[2].Value.Trim();
            var right = comparisonMatch.Groups[3].Value.Trim().Trim('\'', '"');

            if (data.TryGetValue(left, out var leftValue))
            {
                var leftStr = leftValue?.ToString() ?? "";
                
                return op switch
                {
                    "==" => leftStr == right,
                    "!=" => leftStr != right,
                    ">" => CompareValues(leftStr, right) > 0,
                    "<" => CompareValues(leftStr, right) < 0,
                    ">=" => CompareValues(leftStr, right) >= 0,
                    "<=" => CompareValues(leftStr, right) <= 0,
                    _ => false
                };
            }
            
            return false;
        }

        if (data.TryGetValue(condition, out var value))
        {
            if (value is bool boolValue)
                return boolValue;
            
            var strValue = value?.ToString() ?? "";
            return !string.IsNullOrEmpty(strValue) && 
                   !strValue.Equals("false", StringComparison.OrdinalIgnoreCase) &&
                   !strValue.Equals("0", StringComparison.OrdinalIgnoreCase);
        }

        return false;
    }

    private int CompareValues(string left, string right)
    {
        if (double.TryParse(left, out var leftNum) && double.TryParse(right, out var rightNum))
        {
            return leftNum.CompareTo(rightNum);
        }
        
        return string.Compare(left, right, StringComparison.Ordinal);
    }
}
