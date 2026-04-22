namespace DocxEditor.Core.Models;

public class StyleMapping
{
    public Dictionary<string, string> StyleMap { get; init; } = new();

    public string? GetStyle(string markdownElement)
    {
        return StyleMap.TryGetValue(markdownElement, out var style) ? style : null;
    }

    public static StyleMapping Default { get; } = new()
    {
        StyleMap = new Dictionary<string, string>
        {
            ["heading1"] = "Heading1",
            ["heading2"] = "Heading2",
            ["heading3"] = "Heading3",
            ["heading4"] = "Heading4",
            ["heading5"] = "Heading5",
            ["heading6"] = "Heading6",
            ["paragraph"] = "Normal",
            ["blockquote"] = "Quote",
            ["code"] = "Code",
            ["tip"] = "Tip",
            ["warning"] = "Warning",
            ["note"] = "Note"
        }
    };
}
