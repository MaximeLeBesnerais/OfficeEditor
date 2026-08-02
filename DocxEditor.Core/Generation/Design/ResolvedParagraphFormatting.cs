using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Design;

/// <summary>
/// Fully-resolved, immutable paragraph formatting (alignment, space before/after in points, line
/// spacing multiple). The emitters convert these values directly to <c>w:jc</c>/<c>w:spacing</c>;
/// no token interpretation happens at emit time.
/// </summary>
public sealed record ResolvedParagraphFormatting
{
    /// <summary>The identity/empty formatting — nothing set.</summary>
    public static ResolvedParagraphFormatting Empty { get; } = new();

    /// <summary>Horizontal alignment. Null = inherit.</summary>
    public TextAlignment? Alignment { get; init; }

    /// <summary>Space before the paragraph in points. Null = inherit.</summary>
    public double? SpaceBeforePt { get; init; }

    /// <summary>Space after the paragraph in points. Null = inherit.</summary>
    public double? SpaceAfterPt { get; init; }

    /// <summary>Line spacing multiple (1 = single, 1.5 = one-and-a-half). Null = inherit.</summary>
    public double? LineSpacingMultiple { get; init; }

    /// <summary>True when no field is set; such a value needs no properties emitted.</summary>
    public bool IsEmpty =>
        Alignment is null && SpaceBeforePt is null && SpaceAfterPt is null && LineSpacingMultiple is null;
}
