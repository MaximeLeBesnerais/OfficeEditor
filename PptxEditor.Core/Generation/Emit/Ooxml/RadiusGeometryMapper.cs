using DocumentFormat.OpenXml.Drawing;
using PptxEditor.Core.Generation.Model;

namespace PptxEditor.Core.Generation.Emit.Ooxml;

/// <summary>
/// Maps per-corner radii (points) to the OOXML preset-geometry family and its adj values
/// (plan.md §3.3 rect; AGENTS.pptx.md round1Rect/round2SameRect semantics). Pure module —
/// unit-tested directly. adj values live in the spcPct unit family (1/1000ths of a percent
/// of <c>min(w, h)</c>, clamped to the preset maximum 50000 = 50%).
/// </summary>
public static class RadiusGeometryMapper
{
    /// <summary>Maximum adj value for rounded-rect presets (radius = min(w,h) / 2).</summary>
    public const int MaxAdjust = 50000;

    /// <summary>A resolved preset-geometry choice plus its adj1/adj2 values.</summary>
    /// <param name="Preset">The OOXML preset geometry.</param>
    /// <param name="Adj1">Primary radius adjust value (roundRect: the only adj).</param>
    /// <param name="Adj2">Secondary radius adjust value (unused for roundRect).</param>
    public sealed record RadiusGeometry(ShapeTypeValues Preset, int Adj1, int Adj2);

    /// <summary>
    /// Resolves <paramref name="radii"/> against the box size. Returns null when the shape
    /// is square-cornered (caller emits plain <c>rect</c>). Supported native families:
    /// uniform → roundRect; one corner distinct (top-right carries adj1) → round1Rect;
    /// top pair vs bottom pair → round2SameRect; diagonals → round2DiagRect. Anything else
    /// is not expressible without custGeom (Tier 2, plan.md §3.3) and throws loudly.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">Non-positive box dimension.</exception>
    /// <exception cref="ArgumentException">Corner combination outside the native families.</exception>
    public static RadiusGeometry? Map(CornerRadii? radii, double widthPt, double heightPt)
    {
        if (radii is not { } r)
        {
            return null;
        }
        if (widthPt <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(widthPt), widthPt, "Box width must be positive to resolve corner radii.");
        }
        if (heightPt <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(heightPt), heightPt, "Box height must be positive to resolve corner radii.");
        }

        var min = Math.Min(widthPt, heightPt);
        int Adj(double value) => Math.Clamp(
            (int)Math.Round(value / min * 100000.0, MidpointRounding.AwayFromZero), 0, MaxAdjust);

        var tl = Adj(r.TopLeft);
        var tr = Adj(r.TopRight);
        var br = Adj(r.BottomRight);
        var bl = Adj(r.BottomLeft);

        if (tl == 0 && tr == 0 && br == 0 && bl == 0)
        {
            return null; // square corners → plain rect
        }
        if (tl == tr && tr == br && br == bl)
        {
            return new RadiusGeometry(ShapeTypeValues.RoundRectangle, tl, 0);
        }
        if (tl == br && br == bl)
        {
            // round1Rect: adj1 rounds the top-right corner, adj2 the other three.
            return new RadiusGeometry(ShapeTypeValues.Round1Rectangle, tr, tl);
        }
        if (tl == tr && bl == br)
        {
            // round2SameRect: adj1 rounds the two top corners, adj2 the two bottom corners.
            return new RadiusGeometry(ShapeTypeValues.Round2SameRectangle, tl, bl);
        }
        if (tr == bl && tl == br)
        {
            // round2DiagRect: adj1 rounds top-right + bottom-left, adj2 the other diagonal.
            return new RadiusGeometry(ShapeTypeValues.Round2DiagonalRectangle, tr, tl);
        }

        throw new ArgumentException(
            $"Per-corner radius (tl={r.TopLeft}, tr={r.TopRight}, br={r.BottomRight}, bl={r.BottomLeft}) " +
            "is not expressible with native OOXML presets. Supported families: uniform (roundRect), " +
            "top-right distinct from the other three (round1Rect), top pair vs bottom pair " +
            "(round2SameRect), diagonals (round2DiagRect). Arbitrary per-corner radii require " +
            "custGeom, which is Tier 2 (plan.md §3.3).");
    }
}
