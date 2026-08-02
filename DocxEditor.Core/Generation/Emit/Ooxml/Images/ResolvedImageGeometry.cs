using DocxEditor.Core.Generation.Assets;
using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Images;

/// <summary>
/// Resolved drawing geometry for a placed picture: the EMU extents and centering offset of
/// the DrawingML <c>a:xfrm</c>, the source crop (<c>a:srcRect</c>) in spcPct units, and the
/// display size in points. Produced by <see cref="Resolve"/> from the pure
/// <see cref="ImageFitCalculator"/>; consumed by the inline and positioned picture factories.
/// </summary>
public sealed record ResolvedImageGeometry
{
    /// <summary>Display width in EMU (≥ 1).</summary>
    public required long ExtentsCxEmu { get; init; }

    /// <summary>Display height in EMU (≥ 1).</summary>
    public required long ExtentsCyEmu { get; init; }

    /// <summary>Horizontal offset in EMU within the reserved box (contain centering).</summary>
    public long OffsetXEmu { get; init; }

    /// <summary>Vertical offset in EMU within the reserved box.</summary>
    public long OffsetYEmu { get; init; }

    /// <summary>Source crop in spcPct units, or null when nothing is cropped.</summary>
    public ImageSourceRect? SourceRect { get; init; }

    /// <summary>Display width in points.</summary>
    public required double WidthPt { get; init; }

    /// <summary>Display height in points.</summary>
    public required double HeightPt { get; init; }

    /// <summary>
    /// Computes the placement geometry for an asset from the vocabulary fit/crop and an
    /// optional explicit box (points). When the box is omitted the natural size is used.
    /// </summary>
    public static ResolvedImageGeometry Resolve(
        ImageAsset asset, ImageFitMode fit, ImageCrop? crop, double? widthPt, double? heightPt)
    {
        ArgumentNullException.ThrowIfNull(asset);
        var result = ImageFitCalculator.Compute(asset, fit, crop, widthPt, heightPt);
        return new ResolvedImageGeometry
        {
            ExtentsCxEmu = Math.Max(1, DrawingUnits.ToEmu(result.WidthPt)),
            ExtentsCyEmu = Math.Max(1, DrawingUnits.ToEmu(result.HeightPt)),
            OffsetXEmu = result.OffsetXEmu,
            OffsetYEmu = result.OffsetYEmu,
            SourceRect = result.SourceRect,
            WidthPt = result.WidthPt,
            HeightPt = result.HeightPt
        };
    }
}
