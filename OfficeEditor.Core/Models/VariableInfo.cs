namespace OfficeEditor.Core.Models;

public record VariableInfo
{
    public required string Name { get; init; }
    public required string FullMatch { get; init; }
    public required string Location { get; init; }
    public string? DefaultValue { get; init; }
    public int? ParagraphIndex { get; init; }
}
