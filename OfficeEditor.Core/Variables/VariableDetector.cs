using System.Text.RegularExpressions;
using OfficeEditor.Core.Models;

namespace OfficeEditor.Core.Variables;

public class VariableDetector
{
    private static readonly Regex VariablePattern = new(
        @"\{\{([^}|]+)(?:\|([^}]*))?\}\}",
        RegexOptions.Compiled);

    public List<VariableInfo> Scan(string text, string location)
    {
        var variables = new List<VariableInfo>();
        
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

        return variables;
    }

    public List<VariableInfo> ScanElements(IEnumerable<(string Text, string Location)> elements)
    {
        var variables = new List<VariableInfo>();
        
        foreach (var (text, location) in elements)
        {
            var elementVariables = Scan(text, location);
            foreach (var variable in elementVariables)
            {
                if (!variables.Any(v => v.Name == variable.Name && v.Location == variable.Location))
                {
                    variables.Add(variable);
                }
            }
        }

        return variables;
    }
}
