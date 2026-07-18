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

    // Position/extent in EMU (English Metric Units) from the element's a:xfrm.
    // Null when the element carries no explicit transform — placeholder shapes
    // commonly omit a:xfrm and inherit their position from the slide layout.
    public long? X { get; init; }
    public long? Y { get; init; }
    public long? Cx { get; init; }
    public long? Cy { get; init; }
}

public record SlideAnatomy
{
    public int SlideIndex { get; init; }
    public List<SlideElement> Elements { get; init; } = new();
}
