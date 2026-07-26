using System.Globalization;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using PptxEditor.Core.Converters.SmartArt;
using PptxEditor.Core.Models;

namespace DocxEditor.Tests.Unit.SmartArt;

public sealed class SmartArtDrawingExtractorTests
{
    private const string DspNs = "http://schemas.microsoft.com/office/drawing/2008/diagram";
    private const string ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    [Fact]
    public void Extract_RoundRectShape_ReturnsShapeElementWithCornerRadius()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""1904255"" cy=""1142553""/>
    </a:xfrm>
    <a:prstGeom prst=""roundRect"">
      <a:avLst>
        <a:gd name=""adj"" fmla=""val 10000""/>
      </a:avLst>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""C00000""/>
    </a:solidFill>
    <a:ln w=""25400"">
      <a:solidFill>
        <a:srgbClr val=""FFFFFF""/>
      </a:solidFill>
    </a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 150, shapeH: 90);

        Assert.NotNull(result);
        Assert.Equal("Shape", result.Type);
        Assert.True(result.Width > 0);
        Assert.True(result.Height > 0);
        Assert.NotNull(result.Shape);
        Assert.Equal("rect", result.Shape.ShapeType);
        Assert.Equal("#C00000", result.Shape.FillColor);
        Assert.Equal("#FFFFFF", result.Shape.StrokeColor);
        Assert.True(result.Shape.StrokeWidth > 0);
        Assert.True(result.Shape.CornerRadius > 0);
    }

    [Fact]
    public void Extract_RoundRectWithSchemeColor_ReturnsShapeWithMappedFill()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""1904255"" cy=""1142553""/>
    </a:xfrm>
    <a:prstGeom prst=""roundRect"">
      <a:avLst><a:gd name=""adj"" fmla=""val 10000""/></a:avLst>
    </a:prstGeom>
    <a:solidFill>
      <a:schemeClr val=""accent1""/>
    </a:solidFill>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 150, shapeH: 90);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("rect", result.Shape.ShapeType);
        Assert.NotEmpty(result.Shape.FillColor);
    }

    [Theory]
    [InlineData("C00000", "tint", 60000, "#D96666")]   // SmartArt connector pattern: pale accent
    [InlineData("C00000", "shade", 50000, "#600000")]
    [InlineData("4472C4", "lumMod", 50000, "#223962")]
    [InlineData("000000", "lumOff", 20000, "#333333")]
    [InlineData("FFFFFF", "alpha", 15000, "#FFFFFF26")] // 15% opacity → 8-digit hex (Typst rgb accepts #RRGGBBAA)
    public void Extract_FillWithColorTransform_AppliesTransform(string baseColor, string op, int opVal, string expected)
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""1904255"" cy=""1142553""/>
    </a:xfrm>
    <a:prstGeom prst=""rect"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""{baseColor}""><a:{op} val=""{opVal}""/></a:srgbClr>
    </a:solidFill>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 150, shapeH: 90);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal(expected, result.Shape.FillColor);
    }

    [Theory]
    [InlineData("roundRect")]
    [InlineData("round1Rect")]
    [InlineData("round2SameRect")]
    public void Extract_RoundedRectEmptyAvLst_UsesEcmaDefaultCornerRadius(string prst)
    {
        // ECMA-376 default for the rounded-rectangle family is adj = 16667 (1/6 of
        // min(w,h)). Cached SmartArt drawings carry an empty avLst; the shape must
        // still render rounded instead of collapsing to a sharp rectangle.
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""1904255"" cy=""1142553""/>
    </a:xfrm>
    <a:prstGeom prst=""{prst}"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""C00000""/>
    </a:solidFill>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 150, shapeH: 90);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("rect", result.Shape.ShapeType);
        // 16667/100000 * min(150, 90) = 15.0003
        Assert.Equal(15.0, result.Shape.CornerRadius, precision: 1);
    }

    [Fact]
    public void Extract_PlainRectEmptyAvLst_KeepsSharpCorners()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""1904255"" cy=""1142553""/>
    </a:xfrm>
    <a:prstGeom prst=""rect"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""C00000""/>
    </a:solidFill>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 150, shapeH: 90);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal(0.0, result.Shape.CornerRadius);
    }

    [Fact]
    public void Extract_RightArrow_ReturnsPolygonShape()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""2072723"" y=""640544""/>
      <a:ext cx=""403702"" cy=""472255""/>
    </a:xfrm>
    <a:prstGeom prst=""rightArrow"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""ED7D31""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 50, shapeH: 40);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        Assert.NotEmpty(result.Shape.Points);
        Assert.Equal(7, result.Shape.Points.Count);
        Assert.Equal("#ED7D31", result.Shape.FillColor);
    }

    [Fact]
    public void Extract_RotatedShape_ReturnsShapeWithRotation()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm rot=""5400000"">
      <a:off x=""3417127"" y=""1581246""/>
      <a:ext cx=""403702"" cy=""472255""/>
    </a:xfrm>
    <a:prstGeom prst=""rightArrow"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""A5A5A5""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 50, shapeH: 40);

        Assert.NotNull(result);
        Assert.True(result.Rotation > 0);
    }

    [Fact]
    public void Extract_NoGeometry_ReturnsNull()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""1904255"" cy=""1142553""/>
    </a:xfrm>
  </dsp:spPr>
  <dsp:txBody>
    <a:bodyPr/>
    <a:p><a:r><a:t>Text only</a:t></a:r></a:p>
  </dsp:txBody>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 100, shapeH: 100);

        Assert.Null(result);
    }

    [Fact]
    public void Extract_UnsupportedGeometry_ReturnsNull()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""1904255"" cy=""1142553""/>
    </a:xfrm>
    <a:prstGeom prst=""star8"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""C00000""/>
    </a:solidFill>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 100, shapeH: 100);

        Assert.Null(result);
    }

    [Fact]
    public void Extract_WithFrameAndScaleOffset_ComputesPositionCorrectly()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""127000"" y=""254000""/>
      <a:ext cx=""2540000"" cy=""1270000""/>
    </a:xfrm>
    <a:prstGeom prst=""rect"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""4472C4""/>
    </a:solidFill>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element,
            offX: 50, offY: 30,
            scaleX: 0.5, scaleY: 0.5,
            frameX: 100, frameY: 200,
            shapeW: 200, shapeH: 100);

        Assert.NotNull(result);
        Assert.Equal(105, result.X, 0.5);
        Assert.Equal(140, result.Y, 0.5);
        Assert.Equal(100, result.Width, 0.5);
        Assert.Equal(50, result.Height, 0.5);
        Assert.NotNull(result.Shape);
        Assert.Equal("#4472C4", result.Shape.FillColor);
    }

    [Fact]
    public void Extract_NoSpPr_ReturnsNull()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:nvSpPr>
    <dsp:cNvPr id=""0"" name=""""/>
  </dsp:nvSpPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 100, shapeH: 100);

        Assert.Null(result);
    }

    [Fact]
    public void Extract_NoFillNoStroke_RectShapeStillReturns()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""0"" y=""0""/>
      <a:ext cx=""1270000"" cy=""1270000""/>
    </a:xfrm>
    <a:prstGeom prst=""rect"">
      <a:avLst/>
    </a:prstGeom>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 100, shapeH: 100);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("rect", result.Shape.ShapeType);
        Assert.True(string.IsNullOrEmpty(result.Shape.FillColor));
        Assert.True(string.IsNullOrEmpty(result.Shape.StrokeColor));
    }

    [Fact]
    public void Extract_GradientFill_PopulatesFillGradient()
    {
        // showeet corpus pattern: SmartArt colors live in a:gradFill on dsp:sp, not
        // a:solidFill. Two stops — scheme color (static fallback accent1) plus an
        // srgbClr stop with per-stop alpha.
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""1904255"" cy=""1142553""/>
    </a:xfrm>
    <a:prstGeom prst=""rect"">
      <a:avLst/>
    </a:prstGeom>
    <a:gradFill>
      <a:gsLst>
        <a:gs pos=""0""><a:schemeClr val=""accent1""/></a:gs>
        <a:gs pos=""100000""><a:srgbClr val=""0B1026""><a:alpha val=""50000""/></a:srgbClr></a:gs>
      </a:gsLst>
      <a:lin ang=""5400000"" scaled=""1""/>
    </a:gradFill>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 150, shapeH: 90);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.True(string.IsNullOrEmpty(result.Shape.FillColor));
        var gradient = Assert.IsType<TypstGradientFill>(result.Shape.FillGradient);
        Assert.Equal(90.0, gradient.Angle); // 5400000 / 60000
        Assert.Equal(2, gradient.Stops.Count);
        Assert.Equal(new TypstGradientStop("#4472C4", 0.0), gradient.Stops[0]);
        Assert.Equal(new TypstGradientStop("#0B10267F", 1.0), gradient.Stops[1]);
    }

    [Fact]
    public void Extract_GradientFill_ThemeSchemeColors_ResolvedFromMap()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""1904255"" cy=""1142553""/>
    </a:xfrm>
    <a:prstGeom prst=""ellipse"">
      <a:avLst/>
    </a:prstGeom>
    <a:gradFill>
      <a:gsLst>
        <a:gs pos=""0""><a:schemeClr val=""accent1""/></a:gs>
        <a:gs pos=""100000""><a:schemeClr val=""accent2""/></a:gs>
      </a:gsLst>
      <a:lin ang=""0"" scaled=""1""/>
    </a:gradFill>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);
        var schemeColors = new Dictionary<string, string>
        {
            ["accent1"] = "112233",
            ["accent2"] = "AABBCC"
        };

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 150, shapeH: 90,
            schemeColors: schemeColors);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        var gradient = Assert.IsType<TypstGradientFill>(result.Shape.FillGradient);
        Assert.Equal(new TypstGradientStop("#112233", 0.0), gradient.Stops[0]);
        Assert.Equal(new TypstGradientStop("#AABBCC", 1.0), gradient.Stops[1]);
    }

    [Fact]
    public void Extract_SolidFillTakesPrecedenceOverGradient_IgnoresGradient()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""1904255"" cy=""1142553""/>
    </a:xfrm>
    <a:prstGeom prst=""rect"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""C00000""/>
    </a:solidFill>
    <a:gradFill>
      <a:gsLst>
        <a:gs pos=""0""><a:srgbClr val=""000000""/></a:gs>
        <a:gs pos=""100000""><a:srgbClr val=""FFFFFF""/></a:gs>
      </a:gsLst>
      <a:lin ang=""0"" scaled=""1""/>
    </a:gradFill>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 150, shapeH: 90);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("#C00000", result.Shape.FillColor);
        Assert.Null(result.Shape.FillGradient);
    }

    [Fact]
    public void Extract_NoLineElement_SetsNoStrokeFlag()
    {
        // Text-container shapes in cached SmartArt drawings often have no a:ln at
        // all — PowerPoint renders them borderless. The flag makes the Typst emitter
        // write stroke: none instead of inheriting Typst's default 1pt black stroke.
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""1904255"" cy=""1142553""/>
    </a:xfrm>
    <a:prstGeom prst=""rect"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""4472C4""/>
    </a:solidFill>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 150, shapeH: 90);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.True(result.Shape.NoStroke);
    }

    [Fact]
    public void Extract_LineWithNoFill_SetsNoStrokeFlag()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""2072723"" y=""640544""/>
      <a:ext cx=""403702"" cy=""472255""/>
    </a:xfrm>
    <a:prstGeom prst=""rightArrow"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""ED7D31""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 50, shapeH: 40);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.True(result.Shape.NoStroke);
    }

    [Fact]
    public void Extract_ExplicitStroke_ClearsNoStrokeFlag()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""1904255"" cy=""1142553""/>
    </a:xfrm>
    <a:prstGeom prst=""rect"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""4472C4""/>
    </a:solidFill>
    <a:ln w=""25400"">
      <a:solidFill>
        <a:srgbClr val=""FFFFFF""/>
      </a:solidFill>
    </a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 150, shapeH: 90);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.False(result.Shape.NoStroke);
    }

    [Fact]
    public void ComputeFrameFit_SalesDeckCalibration_UniformScaleCentersContent()
    {
        // Slide 15 calibration numbers: drawing bbox 359.86x239.91pt inside a
        // 360x288pt frame at (300,126). PowerPoint renders the cached drawing 1:1
        // in frame space with symmetric ~24pt vertical margins.
        var bounds = (MinX: 0.070, MinY: 24.047, Width: 359.859, Height: 239.906);
        var frame = (X: 300.0, Y: 126.0, Width: 360.0, Height: 288.0);

        var fit = SmartArtDrawingExtractor.ComputeFrameFit(bounds, frame);

        // Uniform (aspect-preserving) scale — the old per-axis stretch gave
        // scaleY ≈ 1.20 and made boxes ~20% too tall.
        Assert.Equal(fit.ScaleX, fit.ScaleY, precision: 6);
        Assert.Equal(360.0 / 359.859, fit.ScaleX, precision: 6);

        // Mapped bounding box is centred in the frame: left == right margin,
        // top == bottom margin ≈ 24pt.
        var left = (fit.FrameX + bounds.MinX) * fit.ScaleX;
        var right = (fit.FrameX + bounds.MinX + bounds.Width) * fit.ScaleX;
        var top = (fit.FrameY + bounds.MinY) * fit.ScaleY;
        var bottom = (fit.FrameY + bounds.MinY + bounds.Height) * fit.ScaleY;

        Assert.Equal(left - frame.X, frame.X + frame.Width - right, precision: 6);
        Assert.Equal(top - frame.Y, frame.Y + frame.Height - bottom, precision: 6);
        Assert.Equal(24.0, top - frame.Y, precision: 1);
    }

    [Fact]
    public void ComputeFrameFit_AspectMatchingBounds_MatchesLegacyStretch()
    {
        // When the drawing bbox aspect matches the frame, uniform fit degenerates
        // to the previous per-axis stretch (no behaviour change).
        var bounds = (MinX: 10.0, MinY: 20.0, Width: 100.0, Height: 50.0);
        var frame = (X: 40.0, Y: 80.0, Width: 200.0, Height: 100.0);

        var fit = SmartArtDrawingExtractor.ComputeFrameFit(bounds, frame);

        Assert.Equal(2.0, fit.ScaleX, precision: 6);
        Assert.Equal(2.0, fit.ScaleY, precision: 6);
        // Legacy formula: shapeFrameX = frameX / scale - minX
        Assert.Equal(40.0 / 2.0 - 10.0, fit.FrameX, precision: 6);
        Assert.Equal(80.0 / 2.0 - 20.0, fit.FrameY, precision: 6);
    }

    [Fact]
    public void ComputeFrameFit_WideContentInTallFrame_KeepsNativeWidthAndCenters()
    {
        // Content spanning the full frame width must not be enlarged; the slack
        // is distributed as symmetric vertical margins (PowerPoint's internal
        // layout margin behaviour).
        var bounds = (MinX: 0.0, MinY: 0.0, Width: 400.0, Height: 200.0);
        var frame = (X: 0.0, Y: 0.0, Width: 400.0, Height: 400.0);

        var fit = SmartArtDrawingExtractor.ComputeFrameFit(bounds, frame);

        Assert.Equal(1.0, fit.ScaleX, precision: 6);
        Assert.Equal(1.0, fit.ScaleY, precision: 6);

        var top = (fit.FrameY + bounds.MinY) * fit.ScaleY;
        var bottom = (fit.FrameY + bounds.MinY + bounds.Height) * fit.ScaleY;
        Assert.Equal(100.0, top, precision: 6);
        Assert.Equal(300.0, bottom, precision: 6);
    }

    [Fact]
    public void Extract_LinePreset_ReturnsRectShape()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""1904255"" cy=""114255""/>
    </a:xfrm>
    <a:prstGeom prst=""line"">
      <a:avLst/>
    </a:prstGeom>
    <a:ln w=""25400"">
      <a:solidFill>
        <a:srgbClr val=""C00000""/>
      </a:solidFill>
    </a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 150, shapeH: 9);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("rect", result.Shape.ShapeType);
        Assert.Equal("#C00000", result.Shape.StrokeColor);
        Assert.True(result.Shape.StrokeWidth > 0);
    }

    [Fact]
    public void Extract_DownArrow_ReturnsPolygonShape()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""403702"" cy=""472255""/>
    </a:xfrm>
    <a:prstGeom prst=""downArrow"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""ED7D31""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 50, shapeH: 40);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        Assert.NotEmpty(result.Shape.Points);
        Assert.Equal(7, result.Shape.Points.Count);
        Assert.Equal("#ED7D31", result.Shape.FillColor);
    }

    [Fact]
    public void Extract_UpArrow_ReturnsPolygonShape()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""403702"" cy=""472255""/>
    </a:xfrm>
    <a:prstGeom prst=""upArrow"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""4472C4""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 50, shapeH: 40);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        Assert.NotEmpty(result.Shape.Points);
        Assert.Equal(7, result.Shape.Points.Count);
    }

    [Fact]
    public void Extract_LeftArrow_ReturnsPolygonShape()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""403702"" cy=""472255""/>
    </a:xfrm>
    <a:prstGeom prst=""leftArrow"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""5B9BD5""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 50, shapeH: 40);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        Assert.NotEmpty(result.Shape.Points);
        Assert.Equal(7, result.Shape.Points.Count);
    }

    [Fact]
    public void Extract_LeftRightArrow_ReturnsPolygonShape()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""403702"" cy=""472255""/>
    </a:xfrm>
    <a:prstGeom prst=""leftRightArrow"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""70AD47""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 60, shapeH: 40);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        Assert.NotEmpty(result.Shape.Points);
        Assert.Equal(10, result.Shape.Points.Count);
    }

    [Fact]
    public void Extract_UpDownArrow_ReturnsPolygonShape()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""403702"" cy=""472255""/>
    </a:xfrm>
    <a:prstGeom prst=""upDownArrow"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""FFC000""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 50, shapeH: 60);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        Assert.NotEmpty(result.Shape.Points);
        Assert.Equal(10, result.Shape.Points.Count);
    }

    [Fact]
    public void Extract_Trapezoid_ReturnsPolygonShape()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""403702"" cy=""472255""/>
    </a:xfrm>
    <a:prstGeom prst=""trapezoid"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""A5A5A5""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 50, shapeH: 40);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        Assert.NotEmpty(result.Shape.Points);
        Assert.Equal(4, result.Shape.Points.Count);
    }

    [Fact]
    public void Extract_CircularArrow_ReturnsPolygonShape()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""403702"" cy=""403702""/>
    </a:xfrm>
    <a:prstGeom prst=""circularArrow"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""70AD47""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 50, shapeH: 50);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        Assert.NotEmpty(result.Shape.Points);

        // ECMA-376 circularArrow (default adjustments): the outline is arc-based,
        // so the flattened polygon carries well over 100 points and spans the top
        // half of the shape box (start angle 180°, end angle ~341°).
        Assert.True(result.Shape.Points.Count > 100);
        Assert.True(result.Shape.Points.Min(p => p.X) < 0.1);
        Assert.True(result.Shape.Points.Max(p => p.X) > 0.9);
        Assert.True(result.Shape.Points.Max(p => p.Y) <= 0.51);
    }

    [Fact]
    public void Extract_CircularArrow_HonorsAdjustmentValues()
    {
        // adj4 is the start angle: 0 starts the arc at 3 o'clock (right edge,
        // first point x ≈ 0.94) instead of the default 180° (left edge, x ≈ 0.06).
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""403702"" cy=""403702""/>
    </a:xfrm>
    <a:prstGeom prst=""circularArrow"">
      <a:avLst>
        <a:gd name=""adj4"" fmla=""val 0""/>
      </a:avLst>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""70AD47""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 50, shapeH: 50);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        Assert.NotEmpty(result.Shape.Points);
        Assert.True(result.Shape.Points[0].X > 0.9);
        Assert.True(Math.Abs(result.Shape.Points[0].Y - 0.5) < 0.05);
    }

    [Fact]
    public void Extract_LeftCircularArrow_ReturnsPolygonShape()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""403702"" cy=""403702""/>
    </a:xfrm>
    <a:prstGeom prst=""leftCircularArrow"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""A5A5A5""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 50, shapeH: 50);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        Assert.NotEmpty(result.Shape.Points);

        // ECMA-376 leftCircularArrow (default adjustments) mirrors circularArrow:
        // the arc spans the bottom half of the shape box.
        Assert.True(result.Shape.Points.Count > 100);
        Assert.True(result.Shape.Points.Min(p => p.X) < 0.1);
        Assert.True(result.Shape.Points.Min(p => p.Y) >= 0.49);
        Assert.True(result.Shape.Points.Max(p => p.Y) > 0.9);
    }

    [Fact]
    public void Extract_Gear6_ReturnsPolygonShape()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""403702"" cy=""403702""/>
    </a:xfrm>
    <a:prstGeom prst=""gear6"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""4472C4""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 40, shapeH: 40);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        Assert.NotEmpty(result.Shape.Points);

        // ECMA-376 gear6 (default adjustments): the round body arcs are flattened,
        // so the outline carries well over 100 points; the top tooth nearly touches
        // the top edge and the side teeth reach x ≈ 0.05 / 0.95.
        Assert.True(result.Shape.Points.Count > 100);
        Assert.True(result.Shape.Points.Min(p => p.X) < 0.07);
        Assert.True(result.Shape.Points.Max(p => p.X) > 0.93);
        Assert.True(result.Shape.Points.Min(p => p.Y) < 0.03);
        Assert.True(result.Shape.Points.Max(p => p.Y) > 0.97);
    }

    [Fact]
    public void Extract_Gear6_HonorsAdjustmentValues()
    {
        // adj1 is the tooth depth: a smaller adj1 gives shallower teeth, i.e. the
        // valleys (minimum distance from center) sit further out.
        var xmlTemplate = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""403702"" cy=""403702""/>
    </a:xfrm>
    <a:prstGeom prst=""gear6"">
      <a:avLst>
        <a:gd name=""adj1"" fmla=""val {{0}}""/>
      </a:avLst>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""4472C4""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        double MinRadius(int adj1)
        {
            var element = ParseXml(string.Format(CultureInfo.InvariantCulture, xmlTemplate, adj1));
            var extracted = SmartArtDrawingExtractor.TryExtractShape(
                element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
                frameX: 0, frameY: 0, shapeW: 40, shapeH: 40);
            Assert.NotNull(extracted);
            Assert.NotNull(extracted.Shape);
            return extracted.Shape.Points.Min(p => Math.Sqrt(
                (p.X - 0.5) * (p.X - 0.5) + (p.Y - 0.5) * (p.Y - 0.5)));
        }

        var shallowTeeth = MinRadius(3000);
        var deepTeeth = MinRadius(20000);

        Assert.True(shallowTeeth > 0.45);
        Assert.True(deepTeeth < 0.32);
        Assert.True(shallowTeeth > deepTeeth + 0.1);
    }

    [Fact]
    public void Extract_Gear9_ReturnsPolygonShape()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""403702"" cy=""403702""/>
    </a:xfrm>
    <a:prstGeom prst=""gear9"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""ED7D31""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 40, shapeH: 40);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        Assert.NotEmpty(result.Shape.Points);

        // ECMA-376 gear9 (default adjustments): flattened body arcs give well over
        // 100 points; teeth nearly touch the left/right/top edges of the box.
        Assert.True(result.Shape.Points.Count > 120);
        Assert.True(result.Shape.Points.Min(p => p.X) < 0.02);
        Assert.True(result.Shape.Points.Max(p => p.X) > 0.98);
        Assert.True(result.Shape.Points.Min(p => p.Y) < 0.02);
    }

    private static OpenXmlElement ParseXml(string xml)
    {
        var xElement = XElement.Parse(xml);
        return ConvertXElementToOpenXml(xElement);
    }

    private static OpenXmlElement ConvertXElementToOpenXml(XElement xElement)
    {
        var prefix = xElement.GetPrefixOfNamespace(xElement.Name.Namespace);
        var localName = xElement.Name.LocalName;
        var nsUri = xElement.Name.NamespaceName;

        var result = new OpenXmlUnknownElement(
            prefix ?? "",
            localName,
            nsUri);

        foreach (var attr in xElement.Attributes())
        {
            if (attr.IsNamespaceDeclaration)
                continue;

            var attrNs = attr.Name.NamespaceName;
            result.SetAttribute(new OpenXmlAttribute(
                string.IsNullOrEmpty(attrNs) ? "" : "",
                attr.Name.LocalName,
                string.IsNullOrEmpty(attrNs) ? "" : attrNs,
                attr.Value));
        }

        foreach (var child in xElement.Elements())
        {
            result.Append(ConvertXElementToOpenXml(child));
        }

        return result;
    }
}
