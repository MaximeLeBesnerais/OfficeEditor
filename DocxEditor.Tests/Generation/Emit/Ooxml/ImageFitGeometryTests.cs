using PptxEditor.Core.Generation.Emit.Ooxml;
using PptxEditor.Core.Models;

namespace DocxEditor.Tests.Generation.Emit.Ooxml;

/// <summary>
/// F7 fit math for from-scratch pictures: fill center-crop srcRect, contain frame shrink.
/// Mirrors the numbers locked in by Unit/ImageFitModeTests.cs for the replace path.
/// </summary>
public sealed class ImageFitGeometryTests
{
    private const long Emu = 12700;

    [Fact]
    public void Fill_WiderImage_CropsLeftRight()
    {
        // 4000×2000 (2:1) into a square frame: visible fraction 0.5 → 25% off each side.
        var rect = ImageFitGeometry.ComputeFillCrop(4000, 2000, 100 * Emu, 100 * Emu);
        Assert.Equal(new SourceRect(25000, 0, 25000, 0), rect);
    }

    [Fact]
    public void Fill_TallerImage_CropsTopBottom()
    {
        var rect = ImageFitGeometry.ComputeFillCrop(2000, 4000, 100 * Emu, 100 * Emu);
        Assert.Equal(new SourceRect(0, 25000, 0, 25000), rect);
    }

    [Fact]
    public void Fill_MatchingAspect_NoCrop()
    {
        Assert.Null(ImageFitGeometry.ComputeFillCrop(2000, 1000, 200 * Emu, 100 * Emu));
    }

    [Fact]
    public void Fill_WideFrame_CropsTopBottom()
    {
        // 3000×2000 (3:2) into a 2:1 frame: visible height fraction 3/4 → 1/8 off top+bottom.
        var rect = ImageFitGeometry.ComputeFillCrop(3000, 2000, 200 * Emu, 100 * Emu);
        Assert.Equal(new SourceRect(0, 12500, 0, 12500), rect);
    }

    [Fact]
    public void Contain_WiderImage_ShrinksHeightAroundCenter()
    {
        // 4000×2000 into 100×100pt: width-limited → cy halves, y shifts by a quarter.
        var (x, y, cx, cy) = ImageFitGeometry.ComputeContainFrame(4000, 2000, 10 * Emu, 20 * Emu, 100 * Emu, 100 * Emu);
        Assert.Equal(10 * Emu, x);
        Assert.Equal(20 * Emu + 25 * Emu, y);
        Assert.Equal(100 * Emu, cx);
        Assert.Equal(50 * Emu, cy);
    }

    [Fact]
    public void Contain_TallerImage_ShrinksWidthAroundCenter()
    {
        var (x, y, cx, cy) = ImageFitGeometry.ComputeContainFrame(2000, 4000, 10 * Emu, 20 * Emu, 100 * Emu, 100 * Emu);
        Assert.Equal(10 * Emu + 25 * Emu, x);
        Assert.Equal(20 * Emu, y);
        Assert.Equal(50 * Emu, cx);
        Assert.Equal(100 * Emu, cy);
    }

    [Fact]
    public void Contain_MatchingAspect_FrameUnchanged()
    {
        var frame = ImageFitGeometry.ComputeContainFrame(2000, 1000, 5 * Emu, 6 * Emu, 200 * Emu, 100 * Emu);
        Assert.Equal((5 * Emu, 6 * Emu, 200 * Emu, 100 * Emu), frame);
    }

    [Theory]
    [InlineData(0, 100)]
    [InlineData(100, 0)]
    public void InvalidImageDimensions_Throw(int w, int h)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => ImageFitGeometry.ComputeFillCrop(w, h, 100, 100));
        Assert.Throws<ArgumentOutOfRangeException>(() => ImageFitGeometry.ComputeContainFrame(w, h, 0, 0, 100, 100));
    }

    [Theory]
    [InlineData(0, 100)]
    [InlineData(100, 0)]
    public void InvalidFrameDimensions_Throw(long cx, long cy)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => ImageFitGeometry.ComputeFillCrop(100, 100, cx, cy));
        Assert.Throws<ArgumentOutOfRangeException>(() => ImageFitGeometry.ComputeContainFrame(100, 100, 0, 0, cx, cy));
    }
}
