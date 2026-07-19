using PptxEditor.Core.Generation.Model;

namespace PptxEditor.Core.Generation.Fixtures;

/// <summary>
/// One parity fixture (plan.md §5, §2 rule 2): a generated deck exercising a single
/// primitive family, diffed as PowerPoint-render (ground truth) vs Typst-render (the
/// spec of record). <see cref="ThresholdRmse"/> is the per-primitive normalized-RMSE
/// ceiling applied per page and per document average.
/// </summary>
public sealed record ParityFixture
{
    /// <summary>Stable slug used for directory names, reports and threshold keys (e.g. "rect-radii").</summary>
    public required string Name { get; init; }

    /// <summary>What the fixture exercises, for humans reading the diff report.</summary>
    public required string Description { get; init; }

    /// <summary>
    /// Per-primitive normalized RMSE ceiling (0–1). Provisional values are calibrated
    /// against PowerPoint ground truth on a render-capable machine; see
    /// tools/visual-diff/baselines/gen/README.md.
    /// </summary>
    public required double ThresholdRmse { get; init; }

    /// <summary>
    /// Builds the fixture document. The argument is the image source path to use for
    /// image elements: callers pass an absolute path when emitting OOXML (the emitter
    /// reads and embeds the file) and a path relative to the emitted .typ file when
    /// emitting Typst (Typst resolves it against the source file's directory). Fixtures
    /// without images ignore it.
    /// </summary>
    public required Func<string, GenerationDocument> BuildDocument { get; init; }
}
