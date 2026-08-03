namespace PptxEditor.Core.Generation.Model;

/// <summary>
/// Design tokens: the palette, fonts, shape defaults and metrics every
/// element references. Content colors resolve against <see cref="Palette"/> by name;
/// raw hex is accepted but warned (off-token drift).
/// </summary>
public sealed record DesignTokens
{
    /// <summary>Named colors, token name → #RRGGBB. Required (may be empty).</summary>
    public required IReadOnlyDictionary<string, string> Palette { get; init; }

    /// <summary>Font slot tokens (display / body). Optional.</summary>
    public FontTokens Fonts { get; init; } = new();

    /// <summary>Shape defaults (corner radius, card style). Optional.</summary>
    public ShapeTokens Shape { get; init; } = new();

    /// <summary>Metric defaults in points. Optional; defaults from the design tokens.</summary>
    public MetricTokens Metrics { get; init; } = new();
}

/// <summary>Font slot tokens referenced by text elements ("display" | "body").</summary>
public sealed record FontTokens
{
    /// <summary>Display/heading typeface family name.</summary>
    public string? Display { get; init; }

    /// <summary>Body typeface family name.</summary>
    public string? Body { get; init; }
}

/// <summary>Shape defaults applied by components and styled containers.</summary>
public sealed record ShapeTokens
{
    /// <summary>Default corner radius in points (0 = square).</summary>
    public double CornerRadius { get; init; }

    /// <summary>Default card treatment (three styles are enough for v1).</summary>
    public CardStyle CardStyle { get; init; } = CardStyle.Flat;
}

/// <summary>Metric defaults in points; all values are pt — never percentages.</summary>
public sealed record MetricTokens
{
    /// <summary>Default slide margin.</summary>
    public double MarginPt { get; init; } = 43;

    /// <summary>Default inter-element gutter.</summary>
    public double GutterPt { get; init; } = 18;

    /// <summary>Default title font size.</summary>
    public double TitleSizePt { get; init; } = 30;

    /// <summary>Default body font size.</summary>
    public double BodySizePt { get; init; } = 14;
}

/// <summary>Card surface treatment (shape.cardStyle).</summary>
public enum CardStyle
{
    /// <summary>Flat solid fill, no border or shadow.</summary>
    Flat,

    /// <summary>Outlined card (stroke, no fill emphasis).</summary>
    Outline,

    /// <summary>Card with a drop shadow.</summary>
    Shadow
}
