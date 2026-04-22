namespace PptxEditor.Core.Models;

public record SlideElement
{
    public required string Type { get; init; }
    public required uint Id { get; init; }
    public required string Name { get; init; }
    public required string Location { get; init; }
    public string? Text { get; init; }
    public List<List<string>>? TableData { get; init; }
    public string? ImagePath { get; init; }
}

public record SlideAnatomy
{
    public int SlideIndex { get; init; }
    public List<SlideElement> Elements { get; init; } = new();
}
