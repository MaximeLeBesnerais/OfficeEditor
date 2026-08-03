using DocumentFormat.OpenXml.Drawing;
using PptxEditor.Core.Generation.Emit.Ooxml;
using PptxEditor.Core.Generation.Model;

namespace DocxEditor.Tests.Generation.Emit.Ooxml;

/// <summary>
/// Per-corner radius → OOXML preset/adj mapping (P4 acceptance). adj values are
/// 1/100000ths of min(w, h), clamped at 50000 (the spcPct family).
/// </summary>
public sealed class RadiusGeometryMapperTests
{
    [Fact]
    public void Map_Null_ReturnsNull()
    {
        Assert.Null(RadiusGeometryMapper.Map(null, 100, 100));
    }

    [Fact]
    public void Map_AllZero_ReturnsNull()
    {
        Assert.Null(RadiusGeometryMapper.Map(CornerRadii.All(0), 100, 100));
    }

    [Fact]
    public void Map_Uniform_RoundRectWithSingleAdj()
    {
        // 10pt on min(200,100)=100 → 10/100 = 10000.
        var geometry = RadiusGeometryMapper.Map(CornerRadii.All(10), 200, 100);
        var expected = new RadiusGeometryMapper.RadiusGeometry(ShapeTypeValues.RoundRectangle, 10000, 0);
        Assert.Equal(expected, geometry);
    }

    [Fact]
    public void Map_Uniform_UsesSmallerDimension()
    {
        // 25pt on min(400,200)=200 → 25/200 = 12500.
        var geometry = RadiusGeometryMapper.Map(CornerRadii.All(25), 400, 200);
        Assert.Equal(12500, geometry!.Adj1);
    }

    [Fact]
    public void Map_OversizedRadius_ClampsAt50000()
    {
        // 80pt on min(100,100)=100 → 80000 → clamped to the preset maximum.
        var geometry = RadiusGeometryMapper.Map(CornerRadii.All(80), 100, 100);
        Assert.Equal(RadiusGeometryMapper.MaxAdjust, geometry!.Adj1);
    }

    [Fact]
    public void Map_TopRightOnly_Round1Rect()
    {
        // round1Rect: adj1 = the top-right corner, adj2 = the other three.
        var geometry = RadiusGeometryMapper.Map(new CornerRadii(0, 10, 0, 0), 200, 100);
        var expected = new RadiusGeometryMapper.RadiusGeometry(ShapeTypeValues.Round1Rectangle, 10000, 0);
        Assert.Equal(expected, geometry);
    }

    [Fact]
    public void Map_ThreeRoundedOneSquare_Round1Rect()
    {
        // All corners 10 except top-right 0: still the round1Rect family (adj1=0, adj2=10000).
        var geometry = RadiusGeometryMapper.Map(new CornerRadii(10, 0, 10, 10), 200, 100);
        var expected = new RadiusGeometryMapper.RadiusGeometry(ShapeTypeValues.Round1Rectangle, 0, 10000);
        Assert.Equal(expected, geometry);
    }

    [Fact]
    public void Map_TopPair_Round2SameRect()
    {
        // round2SameRect: adj1 = the two top corners, adj2 = the two bottom corners.
        var geometry = RadiusGeometryMapper.Map(new CornerRadii(10, 10, 0, 0), 200, 100);
        var expected = new RadiusGeometryMapper.RadiusGeometry(ShapeTypeValues.Round2SameRectangle, 10000, 0);
        Assert.Equal(expected, geometry);
    }

    [Fact]
    public void Map_TopAndBottomPairs_Round2SameRect()
    {
        var geometry = RadiusGeometryMapper.Map(new CornerRadii(20, 20, 5, 5), 200, 100);
        var expected = new RadiusGeometryMapper.RadiusGeometry(ShapeTypeValues.Round2SameRectangle, 20000, 5000);
        Assert.Equal(expected, geometry);
    }

    [Fact]
    public void Map_Diagonal_Round2DiagRect()
    {
        // round2DiagRect: adj1 = top-right + bottom-left, adj2 = the other diagonal.
        var geometry = RadiusGeometryMapper.Map(new CornerRadii(0, 10, 0, 10), 200, 100);
        var expected = new RadiusGeometryMapper.RadiusGeometry(ShapeTypeValues.Round2DiagonalRectangle, 10000, 0);
        Assert.Equal(expected, geometry);
    }

    [Theory]
    [InlineData(10, 0, 0, 0)]   // top-left alone: not a native family
    [InlineData(0, 0, 10, 0)]   // bottom-right alone
    [InlineData(0, 0, 0, 10)]   // bottom-left alone
    [InlineData(10, 0, 0, 10)]  // left pair (rotation would break fills/shadows)
    [InlineData(10, 20, 30, 40)]// four distinct corners → custGeom (Tier 2)
    public void Map_UnsupportedCombination_Throws(double tl, double tr, double br, double bl)
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            RadiusGeometryMapper.Map(new CornerRadii(tl, tr, br, bl), 200, 100));
        Assert.Contains("custGeom", ex.Message);
    }

    [Theory]
    [InlineData(0, 100)]
    [InlineData(100, 0)]
    [InlineData(-5, 100)]
    public void Map_NonPositiveBox_Throws(double w, double h)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            RadiusGeometryMapper.Map(CornerRadii.All(10), w, h));
    }
}
