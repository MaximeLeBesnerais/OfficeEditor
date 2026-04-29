using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeEditor.Core.Models;

namespace DocxEditor.Core.Variables;

public class DocxVariableDetector
{
    private static readonly Regex VariablePattern = new(
        @"\{\{([^}|]+)(?:\|([^}]*))?\}\}",
        RegexOptions.Compiled);

    public List<VariableInfo> Scan(WordprocessingDocument document)
    {
        var variables = new List<VariableInfo>();
        
        // Scan body
        var body = document.MainDocumentPart?.Document?.Body;
        if (body != null)
        {
            variables.AddRange(ScanElement(body, "body"));
        }

        // Scan headers
        var headers = document.MainDocumentPart?.HeaderParts;
        if (headers != null)
        {
            foreach (var header in headers)
            {
                if (header.Header != null)
                {
                    variables.AddRange(ScanElement(header.Header, $"header:{header.Uri}"));
                }
            }
        }

        // Scan footers
        var footers = document.MainDocumentPart?.FooterParts;
        if (footers != null)
        {
            foreach (var footer in footers)
            {
                if (footer.Footer != null)
                {
                    variables.AddRange(ScanElement(footer.Footer, $"footer:{footer.Uri}"));
                }
            }
        }

        return variables;
    }

    private List<VariableInfo> ScanElement(OpenXmlElement element, string location)
    {
        var variables = new List<VariableInfo>();
        var text = element.InnerText;
        
        var matches = VariablePattern.Matches(text);
        foreach (Match match in matches)
        {
            var variableName = match.Groups[1].Value.Trim();
            var defaultValue = match.Groups[2].Success ? match.Groups[2].Value : null;
            
            // Check if we already have this variable
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

        return variables;
    }
}
