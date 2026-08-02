namespace DocxEditor.Core.Generation.Design;

/// <summary>
/// Fully-resolved, immutable shape treatment for the positioned tier (text box, rectangle, callout,
/// line). Fills/strokes are normalized to #RRGGBB hex and already merged with the design
/// <c>shapes</c> defaults, so the positioned emitter consumes plain values and never resolves tokens.
/// </summary>
public sealed record ResolvedShapeStyle
{
    /// <summary>Fill as #RRGGBB. Null = no fill.</summary>
    public string? FillHex { get; init; }

    /// <summary>Stroke color as #RRGGBB. Null = no stroke.</summary>
    public string? StrokeColorHex { get; init; }

    /// <summary>Stroke width in points. Null = no meaningful width (no stroke).</summary>
    public double? StrokeWidthPt { get; init; }

    /// <summary>Corner radius in points (0 = square).</summary>
    public double CornerRadiusPt { get; init; }

    /// <summary>True when a fill is present.</summary>
    public bool HasFill => FillHex is not null;

    /// <summary>True when a stroke is present.</summary>
    public bool HasStroke => StrokeColorHex is not null;
}
