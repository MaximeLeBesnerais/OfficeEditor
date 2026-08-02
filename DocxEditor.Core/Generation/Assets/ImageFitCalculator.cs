using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Assets;

/// <summary>
/// A source-rectangle crop in DrawingML spcPct units (1/1000ths of a percent; 100000 = 100%).
/// Values are the fractions of each edge cropped away from the source, matching the
/// <c>a:srcRect</c> l/t/r/b attributes.
/// </summary>
public readonly record struct ImageSourceRect(int Left, int Top, int Right, int Bottom);

/// <summary>
/// Result of computing the display geometry for an image: the final display box in points
/// (and EMU), a centering offset in EMU (contain within a reserved box), and the source
/// rectangle to crop, when any.
/// </summary>
public sealed record ImageFitResult
{
    /// <summary>Final display width in points.</summary>
    public required double WidthPt { get; init; }

    /// <summary>Final display height in points.</summary>
    public required double HeightPt { get; init; }

    /// <summary>Horizontal centering offset in EMU (nonzero only for contain in a reserved box).</summary>
    public long OffsetXEmu { get; init; }

    /// <summary>Vertical centering offset in EMU.</summary>
    public long OffsetYEmu { get; init; }

    /// <summary>Source crop in spcPct units, or null when nothing is cropped.</summary>
    public ImageSourceRect? SourceRect { get; init; }
}

/// <summary>
/// Pure fit/crop geometry for the generation <see cref="ImageFitMode"/> family. The author
/// crop (0..1 fractions per edge) always applies first and defines the effective source
/// region; the fit mode then maps that region into the display box:
///
/// <list type="bullet">
/// <item><c>Fill</c> — covers the box (center-crops the effective region; distortion-free).</item>
/// <item><c>Contain</c> — fits the effective region inside the box, centered, reserving the box.</item>
/// <item><c>Stretch</c> — fills the box exactly (may distort the effective region).</item>
/// <item><c>Crop</c> — honors the author crop rectangle verbatim; no crop means Fill.</item>
/// </list>
/// </summary>
public static class ImageFitCalculator
{
    private const int SrcRectScale = 100000; // 100000 = 100%

    /// <summary>
    /// Computes the display geometry for an asset. When neither width nor height is
    /// requested the natural (DPI-derived) size is used, reduced by the crop.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="asset"/> is null.</exception>
    /// <exception cref="ArgumentException">The crop removes the whole image, or the box is non-positive.</exception>
    /// <exception cref="NotSupportedException">The fit mode is not one of the vocabulary's four.</exception>
    public static ImageFitResult Compute(
        ImageAsset asset,
        ImageFitMode fit,
        ImageCrop? crop,
        double? requestedWidthPt,
        double? requestedHeightPt)
    {
        ArgumentNullException.ThrowIfNull(asset);

        var cropLeft = crop?.Left ?? 0;
        var cropTop = crop?.Top ?? 0;
        var cropRight = crop?.Right ?? 0;
        var cropBottom = crop?.Bottom ?? 0;

        var visibleWidthFraction = 1.0 - cropLeft - cropRight;
        var visibleHeightFraction = 1.0 - cropTop - cropBottom;
        if (visibleWidthFraction <= 0 || visibleHeightFraction <= 0)
        {
            throw new ArgumentException("The crop rectangle removes the whole image.", nameof(crop));
        }

        var effectiveWidth = asset.Width * visibleWidthFraction;
        var effectiveHeight = asset.Height * visibleHeightFraction;

        var boxWidthPt = requestedWidthPt;
        var boxHeightPt = requestedHeightPt;
        if (boxWidthPt is null && boxHeightPt is null)
        {
            boxWidthPt = asset.NaturalWidthPt * visibleWidthFraction;
            boxHeightPt = asset.NaturalHeightPt * visibleHeightFraction;
        }
        else if (boxWidthPt is null)
        {
            boxWidthPt = boxHeightPt!.Value * effectiveWidth / effectiveHeight;
        }
        else if (boxHeightPt is null)
        {
            boxHeightPt = boxWidthPt.Value * effectiveHeight / effectiveWidth;
        }

        if (boxWidthPt.Value <= 0 || boxHeightPt.Value <= 0)
        {
            throw new ArgumentException("The display box must have positive dimensions.", nameof(requestedWidthPt));
        }

        var baseSourceRect = ToSourceRect(cropLeft, cropTop, cropRight, cropBottom);
        var effectiveFit = fit == ImageFitMode.Crop && crop is null ? ImageFitMode.Fill : fit;

        var widthPt = boxWidthPt.Value;
        var heightPt = boxHeightPt.Value;
        long offsetXEmu = 0;
        long offsetYEmu = 0;
        ImageSourceRect? sourceRect = baseSourceRect;

        switch (effectiveFit)
        {
            case ImageFitMode.Stretch:
                break;

            case ImageFitMode.Contain:
                var (containedWidthPt, containedHeightPt, offsetXPt, offsetYPt) =
                    ContainBox(boxWidthPt.Value, boxHeightPt.Value, effectiveWidth, effectiveHeight);
                widthPt = containedWidthPt;
                heightPt = containedHeightPt;
                offsetXEmu = DrawingUnits.ToEmu(offsetXPt);
                offsetYEmu = DrawingUnits.ToEmu(offsetYPt);
                break;

            case ImageFitMode.Crop:
                sourceRect = baseSourceRect;
                break;

            case ImageFitMode.Fill:
                var cover = ComputeCoverCrop(effectiveWidth, effectiveHeight, boxWidthPt.Value, boxHeightPt.Value);
                sourceRect = ComposeCrop(baseSourceRect, cover, visibleWidthFraction, visibleHeightFraction);
                break;

            default:
                throw new NotSupportedException($"Image fit mode '{fit}' is not supported.");
        }

        return new ImageFitResult
        {
            WidthPt = widthPt,
            HeightPt = heightPt,
            OffsetXEmu = offsetXEmu,
            OffsetYEmu = offsetYEmu,
            SourceRect = sourceRect
        };
    }

    /// <summary>
    /// The largest sub-box of <paramref name="boxWidthPt"/>×<paramref name="boxHeightPt"/> with
    /// the source aspect, centered. Returns the contained box plus the centering deltas.
    /// </summary>
    private static (double WidthPt, double HeightPt, double OffsetXpt, double OffsetYPt) ContainBox(
        double boxWidthPt, double boxHeightPt, double sourceWidth, double sourceHeight)
    {
        var boxAspect = boxWidthPt / boxHeightPt;
        var sourceAspect = sourceWidth / sourceHeight;
        if (Math.Abs(boxAspect - sourceAspect) < 1e-9)
        {
            return (boxWidthPt, boxHeightPt, 0, 0);
        }
        if (sourceAspect > boxAspect)
        {
            var heightPt = boxWidthPt / sourceAspect;
            return (boxWidthPt, heightPt, 0, (boxHeightPt - heightPt) / 2);
        }
        var widthPt = boxHeightPt * sourceAspect;
        return (widthPt, boxHeightPt, (boxWidthPt - widthPt) / 2, 0);
    }

    /// <summary>Center-crop fractions (of the effective region) needed to cover the box, or all-zero.</summary>
    private static (double Left, double Top, double Right, double Bottom) ComputeCoverCrop(
        double sourceWidth, double sourceHeight, double boxWidthPt, double boxHeightPt)
    {
        var boxAspect = boxWidthPt / boxHeightPt;
        var sourceAspect = sourceWidth / sourceHeight;
        if (sourceAspect > boxAspect)
        {
            var visible = boxAspect / sourceAspect;
            var each = (1.0 - visible) / 2.0;
            return (each, 0, each, 0);
        }
        if (sourceAspect < boxAspect)
        {
            var visible = sourceAspect / boxAspect;
            var each = (1.0 - visible) / 2.0;
            return (0, each, 0, each);
        }
        return (0, 0, 0, 0);
    }

    /// <summary>
    /// Composes an author crop (fractions of the full source) with a cover crop (fractions of
    /// the effective, already-cropped region), yielding the final source rectangle.
    /// </summary>
    private static ImageSourceRect? ComposeCrop(
        ImageSourceRect? baseCrop,
        (double Left, double Top, double Right, double Bottom) cover,
        double visibleWidthFraction,
        double visibleHeightFraction)
    {
        var baseLeft = baseCrop?.Left / (double)SrcRectScale ?? 0;
        var baseTop = baseCrop?.Top / (double)SrcRectScale ?? 0;
        var baseRight = baseCrop?.Right / (double)SrcRectScale ?? 0;
        var baseBottom = baseCrop?.Bottom / (double)SrcRectScale ?? 0;

        var left = baseLeft + cover.Left * visibleWidthFraction;
        var top = baseTop + cover.Top * visibleHeightFraction;
        var right = baseRight + cover.Right * visibleWidthFraction;
        var bottom = baseBottom + cover.Bottom * visibleHeightFraction;

        return ToSourceRect(left, top, right, bottom);
    }

    private static ImageSourceRect? ToSourceRect(double left, double top, double right, double bottom)
    {
        var l = ToSrcRectUnit(left);
        var t = ToSrcRectUnit(top);
        var r = ToSrcRectUnit(right);
        var b = ToSrcRectUnit(bottom);
        if (l == 0 && t == 0 && r == 0 && b == 0)
        {
            return null;
        }
        return new ImageSourceRect(l, t, r, b);
    }

    private static int ToSrcRectUnit(double fraction) =>
        (int)Math.Round(fraction * SrcRectScale, MidpointRounding.AwayFromZero);
}
