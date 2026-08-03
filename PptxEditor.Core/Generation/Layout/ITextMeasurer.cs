namespace PptxEditor.Core.Generation.Layout;

/// <summary>
/// Text measurement seam for the overflow pass. P3 implements this in
/// <c>Generation/Layout/TextMeasure.cs</c> on top of TextFitService's font metrics;
/// the resolver consumes it without any OOXML dependency of its own.
/// </summary>
public interface ITextMeasurer
{
    /// <summary>
    /// Returns the largest scale in (0, 1] at which <paramref name="request"/>'s runs fit
    /// the given box; 1 when they fit unscaled. Never returns 0 — below
    /// <see cref="TextMeasureRequest.MinScale"/> the value bottoms out at MinScale.
    /// </summary>
    double FitScale(TextMeasureRequest request);
}

/// <summary>One text-fit measurement request: fully resolved runs and the box they must fit.</summary>
public sealed record TextMeasureRequest
{
    /// <summary>Resolved runs (element defaults already applied).</summary>
    public required IReadOnlyList<ResolvedTextRun> Runs { get; init; }

    /// <summary>Box width available to text in points (element width minus horizontal insets).</summary>
    public required double BoxWidthPt { get; init; }

    /// <summary>Box height available to text in points (element height minus vertical insets).</summary>
    public required double BoxHeightPt { get; init; }

    /// <summary>Smallest acceptable scale (fontScale ≥ MinScale). Defaults to 0.5.</summary>
    public double MinScale { get; init; } = 0.5;
}
