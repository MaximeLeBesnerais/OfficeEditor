using System.Text.RegularExpressions;

namespace OfficeEditor.Core.Variables;

public class VariableReplacer
{
    private static readonly Regex VariablePattern = new(
        @"\{\{([^}|]+)(?:\|([^}]*))?\}\}",
        RegexOptions.Compiled);

    public string Replace(string text, Dictionary<string, string> data)
    {
        return VariablePattern.Replace(text, match =>
        {
            var variableName = match.Groups[1].Value.Trim();
            var defaultValue = match.Groups[2].Success ? match.Groups[2].Value : null;

            if (data.TryGetValue(variableName, out var value))
            {
                return value;
            }
            
            if (!string.IsNullOrEmpty(defaultValue))
            {
                return defaultValue;
            }
            
            return match.Value;
        });
    }
}
