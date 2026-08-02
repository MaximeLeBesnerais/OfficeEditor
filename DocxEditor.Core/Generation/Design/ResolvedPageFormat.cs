using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Design;

/// <summary>
/// Fully-resolved, immutable page geometry for one section. <see cref="WidthPt"/> and
/// <see cref="HeightPt"/> are the effective (post-orientation) dimensions, so the OOXML emitter
/// writes them straight to <c>w:pgSz</c> and never performs its own orientation swap or token
/// resolution. Section fields that are null fall back to design <c>page</c> defaults and then to
/// the built-in A4 / portrait / 1-inch defaults.
/// </summary>
public sealed record ResolvedPageFormat
{
    /// <summary>Effective page width in points (already orientation-aware).</summary>
    public required double WidthPt { get; init; }

    /// <summary>Effective page height in points (already orientation-aware).</summary>
    public required double HeightPt { get; init; }

    /// <summary>Orientation that produced the effective dimensions.</summary>
    public PageOrientation Orientation { get; init; } = PageOrientation.Portrait;

    /// <summary>Resolved page margins in points.</summary>
    public Margins Margins { get; init; } = Margins.Defaults;

    /// <summary>Optional column layout for the section body.</summary>
    public PageColumns? Columns { get; init; }

    /// <summary>Section break applied before this section (null = none/first section).</summary>
    public SectionBreakType? BreakType { get; init; }

    /// <summary>Width of the text area (page width minus left/right margins).</summary>
    public double TextWidthPt => WidthPt - Margins.LeftPt - Margins.RightPt;

    /// <summary>Height of the text area (page height minus top/bottom margins).</summary>
    public double TextHeightPt => HeightPt - Margins.TopPt - Margins.BottomPt;
}
