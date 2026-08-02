namespace DocxEditor.Core.Generation.Model;

/// <summary>
/// Image content shared by inline images (<see cref="ImageElement"/> in flow) and floating
/// images (<see cref="PositionedImage"/> in the positioned tier). Sizes are explicit box
/// constraints in points; when omitted the emitter uses the image's natural size.
/// </summary>
public sealed record ImageElement : FlowBlock
{
    /// <summary>Image source (path, URL or base64 payload — interpreted by emitters). Required.</summary>
    public required string Source { get; init; }

    /// <summary>Fit mode. Defaults to <see cref="ImageFitMode.Fill"/>.</summary>
    public ImageFitMode Fit { get; init; } = ImageFitMode.Fill;

    /// <summary>Optional caller-supplied source crop (0..1 fractions of each edge).</summary>
    public ImageCrop? Crop { get; init; }

    /// <summary>Accessibility alt text.</summary>
    public string? Alt { get; init; }

    /// <summary>Explicit display width in points (&gt; 0). Optional: natural size when absent.</summary>
    public double? WidthPt { get; init; }

    /// <summary>Explicit display height in points (&gt; 0). Optional: natural size when absent.</summary>
    public double? HeightPt { get; init; }
}

/// <summary>Caller-supplied source crop as 0..1 fractions of each edge from the source
/// rectangle (not OOXML units — emitters scale to the target representation).</summary>
public sealed record ImageCrop
{
    /// <summary>Fraction cropped from the left edge.</summary>
    public double Left { get; init; }

    /// <summary>Fraction cropped from the top edge.</summary>
    public double Top { get; init; }

    /// <summary>Fraction cropped from the right edge.</summary>
    public double Right { get; init; }

    /// <summary>Fraction cropped from the bottom edge.</summary>
    public double Bottom { get; init; }
}
