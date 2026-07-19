using PptxEditor.Core.Models;

namespace PptxEditor.Core.Generation.Emit.Ooxml;

/// <summary>
/// Fit-mode geometry for from-scratch pictures (F7: fill/crop/contain). Mirrors the math
/// of <see cref="PptxEditor.Core.Services.PptxElementReplacer"/>'s private ApplyFit — that
/// service is owned by another workstream and operates on existing pictures, so the pure
/// number-crunching lives here for the emitter (and is unit-tested directly).
/// srcRect components are 1/1000ths of a percent (spcPct family, AGENTS.pptx.md rule 2).
/// </summary>
public static class ImageFitGeometry
{
    private const int SrcRectScale = 100000; // 100000 = 100%

    /// <summary>
    /// Fill (cover): center-crop the image via a:srcRect so the visible region matches the
    /// frame aspect. Returns null when the aspects already match (no srcRect needed).
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">Non-positive image or frame dimension.</exception>
    public static SourceRect? ComputeFillCrop(int imageWidth, int imageHeight, long frameCx, long frameCy)
    {
        Validate(imageWidth, imageHeight, frameCx, frameCy);

        // Compare aspects in exact integer cross-products (imageW/imageH vs frameCx/frameCy).
        var imageCross = (long)imageWidth * frameCy;
        var frameCross = (long)imageHeight * frameCx;
        if (imageCross == frameCross)
        {
            return null; // aspects already match: no crop needed
        }

        int left = 0, top = 0, right = 0, bottom = 0;
        if (imageCross > frameCross)
        {
            // Image wider than the frame → crop left/right.
            var visibleFraction = (double)frameCross / imageCross;
            left = right = CropFractionToInt((1.0 - visibleFraction) / 2.0);
        }
        else
        {
            // Image taller than the frame → crop top/bottom.
            var visibleFraction = (double)imageCross / frameCross;
            top = bottom = CropFractionToInt((1.0 - visibleFraction) / 2.0);
        }

        if (left == 0 && top == 0 && right == 0 && bottom == 0)
        {
            return null;
        }
        return new SourceRect(left, top, right, bottom);
    }

    /// <summary>
    /// Contain (fit inside): shrink the frame around its center so the frame aspect matches
    /// the image aspect; the image then fills the smaller frame exactly (never a srcRect).
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">Non-positive image or frame dimension.</exception>
    public static (long X, long Y, long Cx, long Cy) ComputeContainFrame(
        int imageWidth, int imageHeight, long x, long y, long cx, long cy)
    {
        Validate(imageWidth, imageHeight, cx, cy);

        var imageCross = (long)imageWidth * cy;
        var frameCross = (long)imageHeight * cx;
        if (imageCross == frameCross)
        {
            return (x, y, cx, cy); // frame already matches the image aspect
        }

        long newCx, newCy;
        if (imageCross > frameCross)
        {
            // Image wider than the frame → width-limited: keep cx, shrink cy.
            newCx = cx;
            newCy = (long)Math.Round(cx * (double)imageHeight / imageWidth, MidpointRounding.AwayFromZero);
        }
        else
        {
            // Image taller than the frame → height-limited: keep cy, shrink cx.
            newCx = (long)Math.Round(cy * (double)imageWidth / imageHeight, MidpointRounding.AwayFromZero);
            newCy = cy;
        }

        // Shrink around the center: shift the offset by half of each delta.
        return (x + (cx - newCx) / 2, y + (cy - newCy) / 2, newCx, newCy);
    }

    private static void Validate(int imageWidth, int imageHeight, long frameCx, long frameCy)
    {
        if (imageWidth <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(imageWidth), imageWidth, "Image width must be positive.");
        }
        if (imageHeight <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(imageHeight), imageHeight, "Image height must be positive.");
        }
        if (frameCx <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(frameCx), frameCx, "Frame width must be positive.");
        }
        if (frameCy <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(frameCy), frameCy, "Frame height must be positive.");
        }
    }

    private static int CropFractionToInt(double fraction) =>
        (int)Math.Round(fraction * SrcRectScale, MidpointRounding.AwayFromZero);
}
