namespace OfficeEditor.Api.Components.Shared;

public sealed class TimingStats
{
    public int Slides { get; init; }
    public double TotalMs { get; init; }
    public double PerSlideMs { get; init; }
    public string? Label { get; init; }
    public TimingExtra? Extra { get; init; }
}

public sealed class TimingExtra
{
    public required string Label { get; init; }
    public double Ms { get; init; }
}
