using DocumentFormat.OpenXml.Drawing;
using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Flow;

/// <summary>Result of a fit computation: the rendered box (points) plus an optional source crop.</summary>
internal readonly record struct ImageFitResult(double WidthPt, double HeightPt, SourceRectangle? SourceRect);

/// <summary>
/// Fit-mode geometry for inline pictures (fill/contain/crop/stretch), mirroring the math of
/// the PPTX emitter's <c>ImageFitGeometry</c> — this is the flow branch's local copy so the
/// design workstream can replace it during integration. srcRect components are 1/1000ths of
/// a percent (spcPct family): 100000 = 100% of the source edge.
/// </summary>
internal static class ImageFitGeometry
{
    private const int SrcRectScale = 100000;

    /// <summary>
    /// Computes the rendered box and source crop for a fit mode against an author crop.
    /// The author crop (0..1 fractions) is applied to the source first; crop/cover modes then
    /// center-crop the remaining region to the box aspect via a:srcRect.
    /// </summary>
    public static ImageFitResult ComputeFit(
        ImageFitMode fit,
        double naturalWidthPt,
        double naturalHeightPt,
        double boxWidthPt,
        double boxHeightPt,
        ImageCrop? crop)
    {
        if (naturalWidthPt <= 0 || naturalHeightPt <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(naturalWidthPt), "Natural image dimensions must be positive.");
        }
        if (boxWidthPt <= 0 || boxHeightPt <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(boxWidthPt), "Box dimensions must be positive.");
        }

        var left = crop?.Left ?? 0;
        var top = crop?.Top ?? 0;
        var right = crop?.Right ?? 0;
        var bottom = crop?.Bottom ?? 0;
        var croppedWidth = naturalWidthPt * (1 - left - right);
        var croppedHeight = naturalHeightPt * (1 - top - bottom);

        switch (fit)
        {
            case ImageFitMode.Stretch:
                return new ImageFitResult(boxWidthPt, boxHeightPt, AuthorCrop(left, top, right, bottom, null));
            case ImageFitMode.Contain:
            {
                var scale = Math.Min(boxWidthPt / croppedWidth, boxHeightPt / croppedHeight);
                var width = croppedWidth * scale;
                var height = croppedHeight * scale;
                return new ImageFitResult(width, height, AuthorCrop(left, top, right, bottom, null));
            }
            case ImageFitMode.Fill:
            case ImageFitMode.Crop:
            default:
            {
                var cover = ComputeFillCrop(croppedWidth, croppedHeight, boxWidthPt, boxHeightPt);
                return new ImageFitResult(boxWidthPt, boxHeightPt, AuthorCrop(left, top, right, bottom, cover));
            }
        }
    }

    /// <summary>
    /// Combines the author crop (fractions) with an optional cover crop (srcRect units) into
    /// one source rectangle. Returns null when no edge is cropped at all.
    /// </summary>
    private static SourceRectangle? AuthorCrop(double left, double top, double right, double bottom, SourceRectangle? cover)
    {
        var l = CropFractionToInt(left);
        var t = CropFractionToInt(top);
        var r = CropFractionToInt(right);
        var b = CropFractionToInt(bottom);

        if (cover is not null)
        {
            // Cover crop fractions are relative to the author-cropped region; rescale them back
            // to the full source so the resulting srcRect always refers to the original image.
            var horizontalSpan = 1 - left - right;
            var verticalSpan = 1 - top - bottom;
            l += (int)Math.Round((cover.Left?.Value ?? 0) * horizontalSpan, MidpointRounding.AwayFromZero);
            t += (int)Math.Round((cover.Top?.Value ?? 0) * verticalSpan, MidpointRounding.AwayFromZero);
            r += (int)Math.Round((cover.Right?.Value ?? 0) * horizontalSpan, MidpointRounding.AwayFromZero);
            b += (int)Math.Round((cover.Bottom?.Value ?? 0) * verticalSpan, MidpointRounding.AwayFromZero);
        }

        if (l == 0 && t == 0 && r == 0 && b == 0)
        {
            return null;
        }
        return new SourceRectangle { Left = l, Top = t, Right = r, Bottom = b };
    }

    /// <summary>
    /// Center-crop (cover) the source to the box aspect. Returns null when the aspects already
    /// match. The source here is the author-cropped region, so this mirrors the PPTX emitter.
    /// </summary>
    private static SourceRectangle? ComputeFillCrop(double imageWidth, double imageHeight, double boxWidth, double boxHeight)
    {
        if (Math.Abs(imageWidth * boxHeight - imageHeight * boxWidth) < 1e-9)
        {
            return null;
        }

        int left = 0, top = 0, right = 0, bottom = 0;
        if (imageWidth * boxHeight > imageHeight * boxWidth)
        {
            // Source wider than the box → crop left/right.
            var visibleFraction = (boxWidth * imageHeight) / (imageWidth * boxHeight);
            left = right = CropFractionToInt((1 - visibleFraction) / 2.0);
        }
        else
        {
            // Source taller than the box → crop top/bottom.
            var visibleFraction = (imageWidth * boxHeight) / (boxWidth * imageHeight);
            top = bottom = CropFractionToInt((1 - visibleFraction) / 2.0);
        }

        if (left == 0 && top == 0 && right == 0 && bottom == 0)
        {
            return null;
        }
        return new SourceRectangle { Left = left, Top = top, Right = right, Bottom = bottom };
    }

    private static int CropFractionToInt(double fraction) =>
        (int)Math.Round(fraction * SrcRectScale, MidpointRounding.AwayFromZero);
}
