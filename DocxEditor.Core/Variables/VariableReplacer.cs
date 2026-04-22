using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxEditor.Core.Variables;

public class VariableReplacer
{
    private static readonly Regex VariablePattern = new(
        @"\{\{([^}|]+)(?:\|([^}]*))?\}\}",
        RegexOptions.Compiled);

    public void Replace(WordprocessingDocument document, Dictionary<string, string> data)
    {
        // Replace in body
        var body = document.MainDocumentPart?.Document.Body;
        if (body != null)
        {
            ReplaceInElement(body, data);
        }

        // Replace in headers
        var headers = document.MainDocumentPart?.HeaderParts;
        if (headers != null)
        {
            foreach (var header in headers)
            {
                ReplaceInElement(header.Header, data);
            }
        }

        // Replace in footers
        var footers = document.MainDocumentPart?.FooterParts;
        if (footers != null)
        {
            foreach (var footer in footers)
            {
                ReplaceInElement(footer.Footer, data);
            }
        }
    }

    private void ReplaceInElement(OpenXmlElement element, Dictionary<string, string> data)
    {
        var paragraphs = element.Descendants<Paragraph>();
        
        foreach (var paragraph in paragraphs)
        {
            var runs = paragraph.Elements<Run>().ToList();
            
            foreach (var run in runs)
            {
                var texts = run.Elements<Text>().ToList();
                
                foreach (var text in texts)
                {
                    if (text.Text.Contains("{{"))
                    {
                        text.Text = ReplaceVariablesInText(text.Text, data);
                    }
                }
            }
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
            
            // If no data provided, use default value if available
            if (!string.IsNullOrEmpty(defaultValue))
            {
                return defaultValue;
            }
            
            return match.Value; // Keep original if no data and no default
        });
    }
}
