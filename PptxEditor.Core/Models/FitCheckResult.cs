namespace PptxEditor.Core.Models;

/// <summary>
/// Outcome of a text-fit measurement performed by
/// <see cref="PptxEditor.Core.Services.TextFitService"/>.
/// All measurements are points (1/72 inch) at a 100% font-scale baseline:
/// any <c>fontScale</c> already stored on the shape reflects previously
/// entered text and is deliberately ignored (PowerPoint re-baselines on edit).
/// </summary>
public sealed record FitCheckResult
{
    /// <summary>
    /// True when the measured content at 100% scale exceeds the box content area.
    /// When <see cref="AppliedFontScale"/> or <see cref="AppliedLineSpacingReduction"/>
    /// is set, the autofit mechanism resolves the overflow at that scale.
    /// </summary>
    public bool Overflow { get; init; }

    /// <summary>Widest measured content line, in points (at the reported scale baseline).</summary>
    public double ContentWidthPt { get; init; }

    /// <summary>Total measured content height (lines + paragraph spacing), in points.</summary>
    public double ContentHeightPt { get; init; }

    /// <summary>Box width available to text (box width minus left/right insets), in points.</summary>
    public double BoxContentWidthPt { get; init; }

    /// <summary>Box height available to text (box height minus top/bottom insets), in points.</summary>
    public double BoxContentHeightPt { get; init; }

    /// <summary>Font scale (0.5–1.0) required to fit, when auto-shrink applied. Null when no shrink was needed/applied.</summary>
    public double? AppliedFontScale { get; init; }

    /// <summary>Line-spacing reduction (0.1/0.2) applied before/along the font scale. Null when none was applied.</summary>
    public double? AppliedLineSpacingReduction { get; init; }

    /// <summary>True when the replacement part of the operation succeeded (always true for check-only calls).</summary>
    public bool Replaced { get; init; } = true;

    /// <summary>Replacement error, when <see cref="Replaced"/> is false.</summary>
    public string? ReplaceError { get; init; }

    /// <summary>
    /// Non-fatal accuracy notes (font substitution, unparseable TTC, missing metrics,
    /// complex-script shaping limits). Never silently empty when measurement degraded.
    /// </summary>
    public IReadOnlyList<string> Warnings { get; init; } = [];
}
