namespace PptxEditor.Core.Models;

/// <summary>
/// Options for <see cref="PptxEditor.Core.Services.TextFitService.ReplaceTextWithFitCheck"/>.
/// </summary>
public sealed record FitCheckOptions
{
    /// <summary>
    /// When true, auto-shrink is attempted even on shapes without a
    /// <c>&lt;a:normAutofit&gt;</c> body property. Shapes that already carry
    /// <c>normAutofit</c> are always shrink candidates.
    /// </summary>
    public bool AutoShrink { get; init; }

    /// <summary>Lowest font scale the shrink search may reach (default 0.5, mirroring PowerPoint's practical floor).</summary>
    public double MinScale { get; init; } = 0.5;

    /// <summary>
    /// When true (default), line-spacing reductions of 10% and 20% are tried at
    /// full font scale before font shrinking begins, mirroring PowerPoint's order.
    /// </summary>
    public bool AllowLineSpacingReduction { get; init; } = true;

    /// <summary>
    /// When true (default), a successful shrink persists
    /// <c>&lt;a:normAutofit fontScale="…" lnSpcReduction="…"/&gt;</c> on the
    /// <b>slide shape's</b> bodyPr only — never on layout/master parts and
    /// never as a global setting (AGENTS.pptx.md rule 5).
    /// </summary>
    public bool PersistAutoFit { get; init; } = true;
}
