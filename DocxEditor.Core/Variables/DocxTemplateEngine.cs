using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxEditor.Core.Variables;

public class DocxTemplateEngine
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

    public void Process(WordprocessingDocument document, Dictionary<string, object> data)
    {
        var body = document.MainDocumentPart?.Document?.Body;
        if (body != null)
        {
            ProcessElement(body, data);
        }
    }

    private void ProcessElement(OpenXmlElement element, Dictionary<string, object> data)
    {
        var paragraphs = element.Descendants<Paragraph>().ToList();
        
        foreach (var paragraph in paragraphs)
        {
            var text = paragraph.InnerText;
            
            // Process conditionals
            text = ProcessConditionals(text, data);
            
            // Process loops
            text = ProcessLoops(text, data);
            
            // Update paragraph text
            if (text != paragraph.InnerText)
            {
                paragraph.RemoveAllChildren<Run>();
                paragraph.Append(new Run(new Text(text)));
            }
        }
    }

    private string ProcessConditionals(string text, Dictionary<string, object> data)
    {
        // Process {{#if}} blocks
        text = IfPattern.Replace(text, match =>
        {
            var condition = match.Groups[1].Value.Trim();
            var content = match.Groups[2].Value;
            
            return EvaluateCondition(condition, data) ? content : string.Empty;
        });

        // Process {{#ifnot}} blocks
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
        // Check for comparison operators
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

        // Simple boolean check
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
        // Try numeric comparison first
        if (double.TryParse(left, out var leftNum) && double.TryParse(right, out var rightNum))
        {
            return leftNum.CompareTo(rightNum);
        }
        
        // Fall back to string comparison
        return string.Compare(left, right, StringComparison.Ordinal);
    }
}
