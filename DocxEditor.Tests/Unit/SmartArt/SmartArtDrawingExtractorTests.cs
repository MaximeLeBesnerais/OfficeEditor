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
    [InlineData("C00000", "tint", 60000, "#DCA8A8")]   // SmartArt connector pattern: pale accent (gamma-linear blend)
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
    public void Extract_GradientFill_ShadeAndSatModStops_AppliedInDocumentOrder()
    {
        // SmartArt fixtures pattern (slides 25/58/131): gradient stops differ ONLY by
        // shade/satMod transforms. Without them all stops collapse to the raw scheme
        // color and the gradient renders flat.
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
        <a:gs pos=""0""><a:schemeClr val=""accent2""><a:shade val=""51000""/><a:satMod val=""130000""/></a:schemeClr></a:gs>
        <a:gs pos=""80000""><a:schemeClr val=""accent2""><a:shade val=""93000""/><a:satMod val=""130000""/></a:schemeClr></a:gs>
        <a:gs pos=""100000""><a:schemeClr val=""accent2""><a:shade val=""94000""/><a:satMod val=""135000""/></a:schemeClr></a:gs>
      </a:gsLst>
      <a:lin ang=""16200000"" scaled=""0""/>
    </a:gradFill>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);
        var schemeColors = new Dictionary<string, string> { ["accent2"] = "16A085" };

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 150, shapeH: 90,
            schemeColors: schemeColors);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        var gradient = Assert.IsType<TypstGradientFill>(result.Shape.FillGradient);
        Assert.Equal(3, gradient.Stops.Count);
        // shade (per-channel scale, exact in HSL for L<0.5) then satMod (HSL
        // saturation multiply) — all three stops must be DISTINCT dark teal tones.
        Assert.Equal(new TypstGradientStop("#005D4A", 0.0), gradient.Stops[0]);
        Assert.Equal(new TypstGradientStop("#01A888", 0.8), gradient.Stops[1]);
        Assert.Equal(new TypstGradientStop("#00AB8A", 1.0), gradient.Stops[2]);
    }

    [Theory]
    [InlineData("shade", 51000, "#0B5244")]  // per-channel ×0.51 (exact for shade when L<0.5)
    [InlineData("satMod", 130000, "#01B592")] // HSL saturation ×1.3, hue/lightness preserved
    public void Extract_GradientStop_SingleColorTransform_AppliesTransform(string op, int opVal, string expected)
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
    <a:gradFill>
      <a:gsLst>
        <a:gs pos=""0""><a:schemeClr val=""accent2""><a:{op} val=""{opVal}""/></a:schemeClr></a:gs>
        <a:gs pos=""100000""><a:schemeClr val=""accent2""/></a:gs>
      </a:gsLst>
      <a:lin ang=""0"" scaled=""0""/>
    </a:gradFill>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);
        var schemeColors = new Dictionary<string, string> { ["accent2"] = "16A085" };

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 150, shapeH: 90,
            schemeColors: schemeColors);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        var gradient = Assert.IsType<TypstGradientFill>(result.Shape.FillGradient);
        Assert.Equal(new TypstGradientStop(expected, 0.0), gradient.Stops[0]);
        Assert.Equal(new TypstGradientStop("#16A085", 1.0), gradient.Stops[1]);
    }

    [Fact]
    public void Extract_GradientFill_PopulatesFillGradient()
    {
        // SmartArt fixtures pattern: SmartArt colors live in a:gradFill on dsp:sp, not
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
    public void ComputeBoundingBox_UnsupportedPreset_DoesNotPolluteBounds()
    {
        // -033: an unrenderable connector shape (preset not in the map,
        // silently dropped at extraction) must not set the fit bbox — its cached
        // xfrm legitimately extends outside the node layout and would otherwise
        // shrink the whole diagram (slide 33: scale 0.63 instead of ~1.0).
        var rect = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""127000"" y=""254000""/>
      <a:ext cx=""2540000"" cy=""1270000""/>
    </a:xfrm>
    <a:prstGeom prst=""rect"">
      <a:avLst/>
    </a:prstGeom>
  </dsp:spPr>
</dsp:sp>";
        var star = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""-6350000"" y=""-1270000""/>
      <a:ext cx=""8890000"" cy=""7620000""/>
    </a:xfrm>
    <a:prstGeom prst=""star8"">
      <a:avLst/>
    </a:prstGeom>
  </dsp:spPr>
</dsp:sp>";

        var bounds = SmartArtDrawingExtractor.ComputeBoundingBox(
            new[] { ParseXml(rect), ParseXml(star) });

        Assert.NotNull(bounds);
        Assert.Equal(10.0, bounds.Value.MinX, precision: 6);
        Assert.Equal(20.0, bounds.Value.MinY, precision: 6);
        Assert.Equal(200.0, bounds.Value.Width, precision: 6);
        Assert.Equal(100.0, bounds.Value.Height, precision: 6);
    }

    [Fact]
    public void ComputeBoundingBox_MissingPrstGeom_DoesNotPolluteBounds()
    {
        var noGeom = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""-6350000"" y=""-1270000""/>
      <a:ext cx=""8890000"" cy=""7620000""/>
    </a:xfrm>
  </dsp:spPr>
</dsp:sp>";
        var rect = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""127000"" y=""254000""/>
      <a:ext cx=""2540000"" cy=""1270000""/>
    </a:xfrm>
    <a:prstGeom prst=""rect"">
      <a:avLst/>
    </a:prstGeom>
  </dsp:spPr>
</dsp:sp>";

        var bounds = SmartArtDrawingExtractor.ComputeBoundingBox(
            new[] { ParseXml(noGeom), ParseXml(rect) });

        Assert.NotNull(bounds);
        Assert.Equal(10.0, bounds.Value.MinX, precision: 6);
        Assert.Equal(20.0, bounds.Value.MinY, precision: 6);
        Assert.Equal(200.0, bounds.Value.Width, precision: 6);
        Assert.Equal(100.0, bounds.Value.Height, precision: 6);
    }

    [Fact]
    public void ComputeBoundingBox_RotatedShape_UnionsRotatedCorners()
    {
        // -130: quadrants stored as tall rects rotated ±90° must contribute
        // their ROTATED footprint. rot=90° about center (120,50) turns the 40×100pt
        // rect at (100,0) into a 100×40pt footprint at (70,30).
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm rot=""5400000"">
      <a:off x=""1270000"" y=""0""/>
      <a:ext cx=""508000"" cy=""1270000""/>
    </a:xfrm>
    <a:prstGeom prst=""rect"">
      <a:avLst/>
    </a:prstGeom>
  </dsp:spPr>
</dsp:sp>";

        var bounds = SmartArtDrawingExtractor.ComputeBoundingBox(new[] { ParseXml(xml) });

        Assert.NotNull(bounds);
        Assert.Equal(70.0, bounds.Value.MinX, precision: 6);
        Assert.Equal(30.0, bounds.Value.MinY, precision: 6);
        Assert.Equal(100.0, bounds.Value.Width, precision: 6);
        Assert.Equal(40.0, bounds.Value.Height, precision: 6);
    }

    [Fact]
    public void ComputeBoundingBox_RotatedShape270_UnionsRotatedCorners()
    {
        // Slide-30 pattern: rot=270°, wide-short rect becomes tall-wide footprint.
        // rect (0,100) 100×40pt, center (50,120) → footprint 40×100 at (30,70).
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm rot=""16200000"">
      <a:off x=""0"" y=""1270000""/>
      <a:ext cx=""1270000"" cy=""508000""/>
    </a:xfrm>
    <a:prstGeom prst=""rect"">
      <a:avLst/>
    </a:prstGeom>
  </dsp:spPr>
</dsp:sp>";

        var bounds = SmartArtDrawingExtractor.ComputeBoundingBox(new[] { ParseXml(xml) });

        Assert.NotNull(bounds);
        Assert.Equal(30.0, bounds.Value.MinX, precision: 6);
        Assert.Equal(70.0, bounds.Value.MinY, precision: 6);
        Assert.Equal(40.0, bounds.Value.Width, precision: 6);
        Assert.Equal(100.0, bounds.Value.Height, precision: 6);
    }

    [Fact]
    public void ComputeBoundingBox_UnrotatedShapes_UnchangedUnion()
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
  </dsp:spPr>
</dsp:sp>";

        var bounds = SmartArtDrawingExtractor.ComputeBoundingBox(new[] { ParseXml(xml) });

        Assert.NotNull(bounds);
        Assert.Equal(10.0, bounds.Value.MinX, precision: 6);
        Assert.Equal(20.0, bounds.Value.MinY, precision: 6);
        Assert.Equal(200.0, bounds.Value.Width, precision: 6);
        Assert.Equal(100.0, bounds.Value.Height, precision: 6);
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
    public void Extract_LinePreset_ReturnsLineShape()
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
      <a:tailEnd type=""triangle""/>
    </a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 150, shapeH: 9);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("line", result.Shape.ShapeType);
        Assert.Equal((0.0, 0.0), result.Shape.Points[0]);
        Assert.Equal((1.0, 1.0), result.Shape.Points[1]);
        Assert.Equal("#C00000", result.Shape.StrokeColor);
        Assert.True(result.Shape.StrokeWidth > 0);
        Assert.True(result.Shape.ArrowAtEnd);
    }

    [Theory]
    [InlineData(0, 1270000, 0.5, 0.0, 0.5, 1.0)]
    [InlineData(1270000, 0, 0.0, 0.5, 1.0, 0.5)]
    public void Extract_LinePreset_UsesZeroExtentAsAxis(
        long cx, long cy, double startX, double startY, double endX, double endY)
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}""><dsp:spPr>
  <a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""{cx}"" cy=""{cy}""/></a:xfrm>
  <a:prstGeom prst=""line""><a:avLst/></a:prstGeom>
  <a:ln><a:solidFill><a:srgbClr val=""000000""/></a:solidFill></a:ln>
</dsp:spPr></dsp:sp>";

        var result = SmartArtDrawingExtractor.TryExtractShape(
            ParseXml(xml), offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 100, shapeH: 100);

        Assert.NotNull(result?.Shape);
        Assert.Equal((startX, startY), result!.Shape!.Points[0]);
        Assert.Equal((endX, endY), result.Shape.Points[1]);
    }

    [Fact]
    public void Extract_CustomGeometry_UsesPathCoordinateSpaceAndKeepsContours()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}""><dsp:spPr>
  <a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""1270000"" cy=""1270000""/></a:xfrm>
  <a:custGeom>
    <a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l=""0"" t=""0"" r=""0"" b=""0""/>
    <a:pathLst><a:path w=""200"" h=""100""><a:moveTo><a:pt x=""0"" y=""0""/></a:moveTo>
      <a:lnTo><a:pt x=""200"" y=""100""/></a:lnTo><a:close/>
      <a:moveTo><a:pt x=""50"" y=""25""/></a:moveTo>
      <a:lnTo><a:pt x=""150"" y=""25""/></a:lnTo>
    </a:path></a:pathLst>
  </a:custGeom>
  <a:solidFill><a:srgbClr val=""123456""/></a:solidFill>
</dsp:spPr></dsp:sp>";

        var result = SmartArtDrawingExtractor.TryExtractShape(
            ParseXml(xml), offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 100, shapeH: 100);

        Assert.NotNull(result?.Shape);
        Assert.Equal("path", result!.Shape!.ShapeType);
        Assert.Equal(2, result.Shape.Subpaths.Count);
        Assert.Equal((0.0, 0.0), result.Shape.Subpaths[0][0]);
        Assert.Equal((1.0, 1.0), result.Shape.Subpaths[0][1]);
        Assert.Equal((0.25, 0.25), result.Shape.Subpaths[1][0]);
        Assert.Equal((0.75, 0.25), result.Shape.Subpaths[1][1]);
        Assert.Equal([true, false], result.Shape.ClosedSubpaths);
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

    [Fact]
    public void Extract_HomePlate_ReturnsAspectAwarePolygon()
    {
        // ECMA homePlate: point depth = adj * min(w,h) (default adj 50000), so on a
        // wide 80x10pt bar the point sits at x = 1 - 0.5*10/80 = 0.9375 — a static
        // 0.5-width table would be grossly wrong (-025).
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""1016000"" cy=""127000""/>
    </a:xfrm>
    <a:prstGeom prst=""homePlate"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""16A085""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 80, shapeH: 10);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        Assert.Equal(5, result.Shape.Points.Count);
        Assert.Equal((0, 0), result.Shape.Points[0]);
        Assert.Equal(0.9375, result.Shape.Points[1].X, precision: 6);
        Assert.Equal(0.0, result.Shape.Points[1].Y, precision: 6);
        Assert.Equal((1, 0.5), result.Shape.Points[2]);
        Assert.Equal(0.9375, result.Shape.Points[3].X, precision: 6);
        Assert.Equal(1.0, result.Shape.Points[3].Y, precision: 6);
        Assert.Equal((0, 1), result.Shape.Points[4]);
    }

    [Fact]
    public void Extract_HomePlate_HonorsAdjustmentValue()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""1016000"" cy=""127000""/>
    </a:xfrm>
    <a:prstGeom prst=""homePlate"">
      <a:avLst>
        <a:gd name=""adj"" fmla=""val 25000""/>
      </a:avLst>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""16A085""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 80, shapeH: 10);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal(5, result.Shape.Points.Count);
        // dx1 = 0.25 * 10pt → x1 = 77.5/80 = 0.96875
        Assert.Equal(0.96875, result.Shape.Points[1].X, precision: 6);
    }

    [Fact]
    public void Extract_FlowChartManualOperation_ReturnsTrapezoid()
    {
        // ECMA: full-width top edge, bottom edge inset by w/5 on both sides
        // (fixed 5x5 path space — the preset has no avLst).
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""635000"" cy=""127000""/>
    </a:xfrm>
    <a:prstGeom prst=""flowChartManualOperation"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""16A085""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 50, shapeH: 10);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        Assert.Equal(4, result.Shape.Points.Count);
        Assert.Equal((0, 0), result.Shape.Points[0]);
        Assert.Equal((1, 0), result.Shape.Points[1]);
        Assert.Equal(0.8, result.Shape.Points[2].X, precision: 6);
        Assert.Equal(1.0, result.Shape.Points[2].Y, precision: 6);
        Assert.Equal(0.2, result.Shape.Points[3].X, precision: 6);
        Assert.Equal(1.0, result.Shape.Points[3].Y, precision: 6);
    }

    [Fact]
    public void Extract_QuadArrow_HonorsAdjustmentValues()
    {
        // Slide-131 values: adj1 (shaft thickness) 2000, adj2 (head half-width)
        // 4000, adj3 (head length) 5000 — all fractions of min(w,h)=100pt.
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""1270000"" cy=""1270000""/>
    </a:xfrm>
    <a:prstGeom prst=""quadArrow"">
      <a:avLst>
        <a:gd name=""adj1"" fmla=""val 2000""/>
        <a:gd name=""adj2"" fmla=""val 4000""/>
        <a:gd name=""adj3"" fmla=""val 5000""/>
      </a:avLst>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""16A085""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 100, shapeH: 100);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        Assert.Equal(24, result.Shape.Points.Count);
        // Left tip, then head base: head half-width dx2 = 4pt, head length x1 = 5pt,
        // shaft half-width dx3 = 1pt.
        Assert.Equal((0, 0.5), result.Shape.Points[0]);
        Assert.Equal(0.05, result.Shape.Points[1].X, precision: 6);
        Assert.Equal(0.46, result.Shape.Points[1].Y, precision: 6);
        Assert.Equal(0.05, result.Shape.Points[2].X, precision: 6);
        Assert.Equal(0.49, result.Shape.Points[2].Y, precision: 6);
        // Up-arrow tip is the 7th point.
        Assert.Equal((0.5, 0), result.Shape.Points[6]);
        // Right/down tips land on the shape edges.
        Assert.Equal((1, 0.5), result.Shape.Points[12]);
        Assert.Equal((0.5, 1), result.Shape.Points[18]);
    }

    [Fact]
    public void Extract_BlockArc_ReturnsArcBand()
    {
        // ECMA blockArc defaults: 180° → 0° sweep (top-half annulus band),
        // thickness 25% of min(w,h).
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""1270000"" cy=""1270000""/>
    </a:xfrm>
    <a:prstGeom prst=""blockArc"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""16A085""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 100, shapeH: 100);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        // Two 180° arcs flattened at ~2° steps → well over 100 points.
        Assert.True(result.Shape.Points.Count > 100);
        // Band spans the full width and the TOP half of the box.
        Assert.Equal(0.0, result.Shape.Points[0].X, precision: 3);
        Assert.Equal(0.5, result.Shape.Points[0].Y, precision: 3);
        Assert.True(result.Shape.Points.Min(p => p.Y) < 0.02);
        Assert.True(result.Shape.Points.Max(p => p.Y) <= 0.51);
        Assert.True(result.Shape.Points.Max(p => p.X) > 0.98);
    }

    [Fact]
    public void Extract_BlockArc_HonorsAdjustmentValues()
    {
        // Slide-33 connector: start 315°, end 45°, thin band (adj3 376).
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""1270000"" cy=""1270000""/>
    </a:xfrm>
    <a:prstGeom prst=""blockArc"">
      <a:avLst>
        <a:gd name=""adj1"" fmla=""val 18900000""/>
        <a:gd name=""adj2"" fmla=""val 2700000""/>
        <a:gd name=""adj3"" fmla=""val 376""/>
      </a:avLst>
    </a:prstGeom>
    <a:noFill/>
    <a:ln w=""9525"">
      <a:solidFill><a:srgbClr val=""16A085""/></a:solidFill>
    </a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 100, shapeH: 100);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        // Outer-arc start at 315°: (0.5 + 0.5·cos315°, 0.5 + 0.5·sin315°).
        Assert.Equal(0.8536, result.Shape.Points[0].X, precision: 3);
        Assert.Equal(0.1464, result.Shape.Points[0].Y, precision: 3);
        // The 90° arc band occupies the right side of the box only.
        Assert.True(result.Shape.Points.Min(p => p.X) > 0.8);
        Assert.True(result.Shape.Points.Max(p => p.X) <= 1.0);
    }

    [Fact]
    public void ComputeBoundingBox_NoFillBlockArc_DoesNotPolluteBounds()
    {
        // -033: a stroke-only connector arc (noFill blockArc) is clipped to
        // the frame by PowerPoint, never fitted — its cached xfrm legitimately
        // extends outside the node layout, so it must not set the fit bbox.
        var rect = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""127000"" y=""254000""/>
      <a:ext cx=""2540000"" cy=""1270000""/>
    </a:xfrm>
    <a:prstGeom prst=""rect"">
      <a:avLst/>
    </a:prstGeom>
  </dsp:spPr>
</dsp:sp>";
        var connectorArc = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""-6350000"" y=""-1270000""/>
      <a:ext cx=""8890000"" cy=""7620000""/>
    </a:xfrm>
    <a:prstGeom prst=""blockArc"">
      <a:avLst>
        <a:gd name=""adj1"" fmla=""val 18900000""/>
        <a:gd name=""adj2"" fmla=""val 2700000""/>
        <a:gd name=""adj3"" fmla=""val 376""/>
      </a:avLst>
    </a:prstGeom>
    <a:noFill/>
    <a:ln w=""9525"">
      <a:solidFill><a:srgbClr val=""16A085""/></a:solidFill>
    </a:ln>
  </dsp:spPr>
</dsp:sp>";

        var bounds = SmartArtDrawingExtractor.ComputeBoundingBox(
            new[] { ParseXml(rect), ParseXml(connectorArc) });

        Assert.NotNull(bounds);
        Assert.Equal(10.0, bounds.Value.MinX, precision: 6);
        Assert.Equal(20.0, bounds.Value.MinY, precision: 6);
        Assert.Equal(200.0, bounds.Value.Width, precision: 6);
        Assert.Equal(100.0, bounds.Value.Height, precision: 6);
    }

    [Fact]
    public void ComputeBoundingBox_FilledBlockArc_ContributesToBounds()
    {
        // A FILLED blockArc is real content (e.g. donut-chart segments) and must
        // still set the fit bbox.
        var arc = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""-6350000"" y=""-1270000""/>
      <a:ext cx=""8890000"" cy=""7620000""/>
    </a:xfrm>
    <a:prstGeom prst=""blockArc"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""16A085""/>
    </a:solidFill>
  </dsp:spPr>
</dsp:sp>";

        var bounds = SmartArtDrawingExtractor.ComputeBoundingBox(new[] { ParseXml(arc) });

        Assert.NotNull(bounds);
        Assert.Equal(-500.0, bounds.Value.MinX, precision: 6);
        Assert.Equal(700.0, bounds.Value.Width, precision: 6);
    }

    [Fact]
    public void Extract_LeftRightRibbon_ReturnsRibbonOutline()
    {
        // ECMA leftRightRibbon (default adjustments, slide-79 aspect 2.5:1):
        // dy1 = h·a1/200000 = 0.25h, dy2 = -h·a3/200000 = -h/12 — the ribbon's
        // top fold starts at ly1 = vc+dy2-dy1 ≈ 0.1667h and the left tip at
        // ly2 ≈ 0.4167h.
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""2540000"" cy=""1016000""/>
    </a:xfrm>
    <a:prstGeom prst=""leftRightRibbon"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""16A085""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 200, shapeH: 80);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        // Two 180° fold arcs + one 90° arc flattened at ~2° steps → >100 points.
        Assert.True(result.Shape.Points.Count > 100);
        // Left tip and top-edge start (adj2 = 50000 → x1 = ss/2 = 40/200 = 0.2).
        Assert.Equal(0.0, result.Shape.Points[0].X, precision: 9);
        Assert.Equal(0.416665, result.Shape.Points[0].Y, precision: 6);
        Assert.Equal(0.2, result.Shape.Points[1].X, precision: 6);
        Assert.Equal(0.0, result.Shape.Points[1].Y, precision: 6);
        Assert.Equal(0.2, result.Shape.Points[2].X, precision: 6);
        Assert.Equal(0.166665, result.Shape.Points[2].Y, precision: 6);
        Assert.Equal(0.5, result.Shape.Points[3].X, precision: 6);
        Assert.Equal(0.166665, result.Shape.Points[3].Y, precision: 6);
        // Right tip mirrored at x = 1: ry3 = 1 - ly2.
        Assert.Contains(result.Shape.Points, p =>
            Math.Abs(p.X - 1) < 1e-9 && Math.Abs(p.Y - (1 - 0.416665)) < 1e-6);
    }

    [Fact]
    public void Extract_LeftRightRibbon_HonorsAdjustmentValues()
    {
        // adj2 = 25000 → the ribbon's top edge starts at x1 = ss·0.25 instead of ss·0.5.
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""2540000"" cy=""1016000""/>
    </a:xfrm>
    <a:prstGeom prst=""leftRightRibbon"">
      <a:avLst>
        <a:gd name=""adj1"" fmla=""val 50000""/>
        <a:gd name=""adj2"" fmla=""val 25000""/>
        <a:gd name=""adj3"" fmla=""val 16667""/>
      </a:avLst>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""16A085""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 200, shapeH: 80);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        // x1 = 80·0.25 = 20pt → 20/200 = 0.1
        Assert.Equal(0.1, result.Shape.Points[1].X, precision: 6);
        Assert.Equal(0.0, result.Shape.Points[1].Y, precision: 6);
    }

    [Fact]
    public void Extract_UpArrowCallout_ReturnsArrowPolygon()
    {
        // ECMA upArrowCallout (default adjustments) on a 200x100pt box:
        // ss = 100; dx1 = ss·a2/100000 = 25, dx2 = ss·a1/200000 = 12.5,
        // y1 = ss·a3/100000 = 25, y2 = h·(1 - a4/100000) = 35.023.
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""2540000"" cy=""1270000""/>
    </a:xfrm>
    <a:prstGeom prst=""upArrowCallout"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""16A085""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 200, shapeH: 100);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        Assert.Equal(11, result.Shape.Points.Count);
        // Callout body bottom edge, then the arrow shaft/head on top.
        Assert.Equal(0.0, result.Shape.Points[0].X, precision: 9);
        Assert.Equal(0.35023, result.Shape.Points[0].Y, precision: 5);
        Assert.Equal(0.4375, result.Shape.Points[1].X, precision: 6);
        Assert.Equal(0.375, result.Shape.Points[3].X, precision: 6);
        Assert.Equal(0.25, result.Shape.Points[3].Y, precision: 6);
        Assert.Equal(0.5, result.Shape.Points[4].X, precision: 9);
        Assert.Equal(0.0, result.Shape.Points[4].Y, precision: 9);
        Assert.Equal(0.625, result.Shape.Points[5].X, precision: 6);
        Assert.Equal(1.0, result.Shape.Points[8].X, precision: 9);
        Assert.Equal(0.35023, result.Shape.Points[8].Y, precision: 5);
        Assert.Equal(1.0, result.Shape.Points[9].X, precision: 9);
        Assert.Equal(1.0, result.Shape.Points[9].Y, precision: 9);
    }

    [Fact]
    public void Extract_UpArrowCallout_HonorsAdjustmentValues()
    {
        // adj3 = 50000 → the arrow head is twice as tall (y1 = 0.5·ss).
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""2540000"" cy=""1270000""/>
    </a:xfrm>
    <a:prstGeom prst=""upArrowCallout"">
      <a:avLst>
        <a:gd name=""adj1"" fmla=""val 25000""/>
        <a:gd name=""adj2"" fmla=""val 25000""/>
        <a:gd name=""adj3"" fmla=""val 50000""/>
        <a:gd name=""adj4"" fmla=""val 64977""/>
      </a:avLst>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""16A085""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 200, shapeH: 100);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        // y1 = 100·0.5 = 50pt → 0.5
        Assert.Equal(0.5, result.Shape.Points[2].Y, precision: 6);
        Assert.Equal(0.5, result.Shape.Points[3].Y, precision: 6);
    }

    [Fact]
    public void Extract_Pie_ReturnsSectorPolygon()
    {
        // ECMA pie defaults: adj1 = 0 (start angle), adj2 = 16200000 (270°) →
        // a three-quarter sector from 3 o'clock sweeping clockwise to 12 o'clock.
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""1270000"" cy=""1270000""/>
    </a:xfrm>
    <a:prstGeom prst=""pie"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""16A085""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 100, shapeH: 100);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        // 270° arc flattened at ~2° steps → well over 100 points.
        Assert.True(result.Shape.Points.Count > 100);
        // Arc starts at angle 0 → right midpoint; closes back to the center.
        Assert.Equal(1.0, result.Shape.Points[0].X, precision: 3);
        Assert.Equal(0.5, result.Shape.Points[0].Y, precision: 3);
        Assert.Equal(0.5, result.Shape.Points[^1].X, precision: 6);
        Assert.Equal(0.5, result.Shape.Points[^1].Y, precision: 6);
        // The sweep covers the bottom and left quadrants: max Y at 90°.
        Assert.True(result.Shape.Points.Max(p => p.Y) > 0.98);
        Assert.True(result.Shape.Points.Min(p => p.X) < 0.02);
        Assert.True(result.Shape.Points.Min(p => p.Y) < 0.02);
    }

    [Fact]
    public void Extract_Pie_HonorsAdjustmentValues()
    {
        // Slide-87/88 wedge: adj1 = 16200000 (270°), adj2 = 3240000 (54°) —
        // end < start, so the sweep wraps through +360° (144° wedge).
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""1270000"" cy=""1270000""/>
    </a:xfrm>
    <a:prstGeom prst=""pie"">
      <a:avLst>
        <a:gd name=""adj1"" fmla=""val 16200000""/>
        <a:gd name=""adj2"" fmla=""val 3240000""/>
      </a:avLst>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""16A085""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 100, shapeH: 100);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        // Arc starts at 270° → top midpoint (0.5, 0).
        Assert.Equal(0.5, result.Shape.Points[0].X, precision: 3);
        Assert.Equal(0.0, result.Shape.Points[0].Y, precision: 3);
        // 144° sweep → the wedge never reaches the left half of the box.
        Assert.True(result.Shape.Points.Min(p => p.X) > 0.4);
        // Ends at 54° → (0.5 + 0.5·cos54°, 0.5 + 0.5·sin54°) ≈ (0.794, 0.905).
        Assert.Contains(result.Shape.Points, p =>
            Math.Abs(p.X - 0.7939) < 0.01 && Math.Abs(p.Y - 0.9045) < 0.01);
    }

    [Fact]
    public void Extract_PieWedge_ReturnsQuarterSector()
    {
        // ECMA pieWedge has no avLst: a fixed quarter-ellipse sector with the
        // arc centered on the bottom-right corner (180° → 270° sweep).
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""1270000"" cy=""1270000""/>
    </a:xfrm>
    <a:prstGeom prst=""pieWedge"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""16A085""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 100, shapeH: 100);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        // 90° arc flattened at ~2° steps → ~46 arc points + moveTo + lnTo.
        Assert.True(result.Shape.Points.Count > 40);
        // Starts at bottom-left, arcs to top-right, closes at bottom-right.
        Assert.Equal(0.0, result.Shape.Points[0].X, precision: 9);
        Assert.Equal(1.0, result.Shape.Points[0].Y, precision: 9);
        Assert.Equal(1.0, result.Shape.Points[^2].X, precision: 3);
        Assert.Equal(0.0, result.Shape.Points[^2].Y, precision: 3);
        Assert.Equal(1.0, result.Shape.Points[^1].X, precision: 9);
        Assert.Equal(1.0, result.Shape.Points[^1].Y, precision: 9);
        // The arc bows toward the top-left: near 225° the point approaches
        // (1 - √2/2, 1 - √2/2) ≈ (0.293, 0.293).
        Assert.Contains(result.Shape.Points, p =>
            Math.Abs(p.X - 0.2929) < 0.02 && Math.Abs(p.Y - 0.2929) < 0.02);
    }

    [Fact]
    public void Extract_Trapezoid_HonorsAdjustmentValues()
    {
        // -b1 §1: the cached trapezoid adj (64780 on slides
        // 134/135) must set the top-edge inset — ECMA-376 x1 = ss·adj/100000
        // with ss = min(w,h), NOT the static table's fixed 25% of the width.
        // 200x100pt shape, adj=64780: x1 = 100·0.6478 = 64.78pt → 0.3239 of w.
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""0"" y=""0""/>
      <a:ext cx=""2540000"" cy=""1270000""/>
    </a:xfrm>
    <a:prstGeom prst=""trapezoid"">
      <a:avLst><a:gd name=""adj"" fmla=""val 64780""/></a:avLst>
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
            frameX: 0, frameY: 0, shapeW: 200, shapeH: 100);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        Assert.Equal(4, result.Shape.Points.Count);
        Assert.Equal(0.3239, result.Shape.Points[0].X, precision: 4);
        Assert.Equal(0.0, result.Shape.Points[0].Y, precision: 9);
        Assert.Equal(0.6761, result.Shape.Points[1].X, precision: 4);
        Assert.Equal(1.0, result.Shape.Points[2].X, precision: 9);
        Assert.Equal(1.0, result.Shape.Points[2].Y, precision: 9);
        Assert.Equal(0.0, result.Shape.Points[3].X, precision: 9);
        Assert.Equal(1.0, result.Shape.Points[3].Y, precision: 9);
    }

    [Fact]
    public void Extract_Trapezoid_EmptyAvLst_UsesEcmaDefaultAdj()
    {
        // ECMA-376 default adj = 25000: on a square shape (ss = w) the inset is
        // 25% of the width — identical to the legacy static-table rendering.
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""0"" y=""0""/>
      <a:ext cx=""1270000"" cy=""1270000""/>
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
            frameX: 0, frameY: 0, shapeW: 100, shapeH: 100);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        Assert.Equal(4, result.Shape.Points.Count);
        Assert.Equal(0.25, result.Shape.Points[0].X, precision: 6);
        Assert.Equal(0.75, result.Shape.Points[1].X, precision: 6);
    }

    [Fact]
    public void Extract_NonIsoscelesTrapezoid_ReturnsPolygonShape()
    {
        // ECMA-376 defaults adj1 = adj2 = 20000: square shape → 20% insets.
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""0"" y=""0""/>
      <a:ext cx=""1270000"" cy=""1270000""/>
    </a:xfrm>
    <a:prstGeom prst=""nonIsoscelesTrapezoid"">
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
            frameX: 0, frameY: 0, shapeW: 100, shapeH: 100);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        Assert.Equal(4, result.Shape.Points.Count);
        Assert.Equal(0.2, result.Shape.Points[0].X, precision: 6);
        Assert.Equal(0.0, result.Shape.Points[0].Y, precision: 9);
        Assert.Equal(0.8, result.Shape.Points[1].X, precision: 6);
        Assert.Equal(0.0, result.Shape.Points[1].Y, precision: 9);
        Assert.Equal(1.0, result.Shape.Points[2].X, precision: 9);
        Assert.Equal(1.0, result.Shape.Points[2].Y, precision: 9);
        Assert.Equal(0.0, result.Shape.Points[3].X, precision: 9);
        Assert.Equal(1.0, result.Shape.Points[3].Y, precision: 9);
    }

    [Fact]
    public void Extract_NonIsoscelesTrapezoid_HonorsAdjustmentValues()
    {
        // Slide-134 label-box pattern: adj1 = 0, adj2 = 64780 on a 200x100pt
        // shape → x1 = 0, x2 = 100·0.6478 = 64.78pt → right top corner at
        // (200−64.78)/200 = 0.6761 of the width.
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""0"" y=""0""/>
      <a:ext cx=""2540000"" cy=""1270000""/>
    </a:xfrm>
    <a:prstGeom prst=""nonIsoscelesTrapezoid"">
      <a:avLst>
        <a:gd name=""adj1"" fmla=""val 0""/>
        <a:gd name=""adj2"" fmla=""val 64780""/>
      </a:avLst>
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
            frameX: 0, frameY: 0, shapeW: 200, shapeH: 100);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        Assert.Equal(4, result.Shape.Points.Count);
        Assert.Equal(0.0, result.Shape.Points[0].X, precision: 9);
        Assert.Equal(0.6761, result.Shape.Points[1].X, precision: 4);
    }

    [Fact]
    public void ComputeBoundingBox_NonIsoscelesTrapezoid_ContributesToBounds()
    {
        // -b1 §1: slides 134/135 regressed because the bbox
        // pollution skip dropped the nonIsoscelesTrapezoid label boxes
        // (x 217.6→640), shrinking the bbox to the 435.2pt-wide pyramid and
        // center-shifting the whole diagram +102pt. Now that the preset is
        // renderable, it must participate in the fit bbox again.
        var pyramid = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""0"" y=""0""/>
      <a:ext cx=""5527040"" cy=""4265930""/>
    </a:xfrm>
    <a:prstGeom prst=""trapezoid"">
      <a:avLst><a:gd name=""adj"" fmla=""val 64780""/></a:avLst>
    </a:prstGeom>
  </dsp:spPr>
</dsp:sp>";
        var labelBox = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""2763520"" y=""0""/>
      <a:ext cx=""5364480"" cy=""4265930""/>
    </a:xfrm>
    <a:prstGeom prst=""nonIsoscelesTrapezoid"">
      <a:avLst>
        <a:gd name=""adj1"" fmla=""val 0""/>
        <a:gd name=""adj2"" fmla=""val 64780""/>
      </a:avLst>
    </a:prstGeom>
  </dsp:spPr>
</dsp:sp>";

        var bounds = SmartArtDrawingExtractor.ComputeBoundingBox(
            new[] { ParseXml(pyramid), ParseXml(labelBox) });

        Assert.NotNull(bounds);
        Assert.Equal(0.0, bounds.Value.MinX, precision: 6);
        Assert.Equal(0.0, bounds.Value.MinY, precision: 6);
        // 2763520 + 5364480 = 8128000 EMU = 640pt — the identity-fit bbox.
        Assert.Equal(640.0, bounds.Value.Width, precision: 6);
        Assert.Equal(335.9, bounds.Value.Height, precision: 6);
    }

    [Fact]
    public void Extract_WedgeRectCallout_DegenerateAdjustments_RendersAsRect()
    {
        // -b1 §3: slide 49's first column body caches
        // adj1 = adj2 = 0 — the tip lands on the shape centre, a
        // self-intersecting bowtie — so the outline must degenerate to the
        // plain rect (the fifth point collapses onto the last rect vertex).
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""0"" y=""0""/>
      <a:ext cx=""2540000"" cy=""1270000""/>
    </a:xfrm>
    <a:prstGeom prst=""wedgeRectCallout"">
      <a:avLst>
        <a:gd name=""adj1"" fmla=""val 0""/>
        <a:gd name=""adj2"" fmla=""val 0""/>
      </a:avLst>
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
            frameX: 0, frameY: 0, shapeW: 200, shapeH: 100);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        Assert.Equal(5, result.Shape.Points.Count);
        Assert.Equal((0.0, 0.0), result.Shape.Points[0]);
        Assert.Equal((1.0, 0.0), result.Shape.Points[1]);
        Assert.Equal((1.0, 1.0), result.Shape.Points[2]);
        Assert.Equal((0.0, 1.0), result.Shape.Points[3]);
        // Degenerate tip: collapses onto (l,b) — no interior point, no bowtie.
        Assert.Equal((0.0, 1.0), result.Shape.Points[4]);
    }

    [Fact]
    public void Extract_WedgeRectCallout_HonorsAdjustmentValues()
    {
        // Slide-149 caption pattern: adj1 = 20250, adj2 = -60700 on a
        // 200x100pt shape → tip at (hc + 0.2025·w, vc − 0.607·h)
        // = (140.5, -10.7)pt → (0.7025, -0.107) normalized — the tip pokes
        // ABOVE the top edge (the reference's caption bump).
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""0"" y=""0""/>
      <a:ext cx=""2540000"" cy=""1270000""/>
    </a:xfrm>
    <a:prstGeom prst=""wedgeRectCallout"">
      <a:avLst>
        <a:gd name=""adj1"" fmla=""val 20250""/>
        <a:gd name=""adj2"" fmla=""val -60700""/>
      </a:avLst>
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
            frameX: 0, frameY: 0, shapeW: 200, shapeH: 100);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        Assert.Equal(5, result.Shape.Points.Count);
        Assert.Equal(0.7025, result.Shape.Points[4].X, precision: 4);
        Assert.Equal(-0.107, result.Shape.Points[4].Y, precision: 3);
    }

    [Fact]
    public void Extract_WedgeRectCallout_EmptyAvLst_UsesEcmaDefaults()
    {
        // ECMA-376 defaults adj1 = -20833, adj2 = 62500: tip below the
        // bottom edge, left of centre: (0.5 − 0.20833, 0.5 + 0.625).
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""0"" y=""0""/>
      <a:ext cx=""1270000"" cy=""1270000""/>
    </a:xfrm>
    <a:prstGeom prst=""wedgeRectCallout"">
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
            frameX: 0, frameY: 0, shapeW: 100, shapeH: 100);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        Assert.Equal(5, result.Shape.Points.Count);
        Assert.Equal(0.29167, result.Shape.Points[4].X, precision: 4);
        Assert.Equal(1.125, result.Shape.Points[4].Y, precision: 3);
    }

    [Fact]
    public void ComputeBoundingBox_WedgeRectCallout_ContributesToBounds()
    {
        // -b1 §3: slide 49 regressed because the bbox
        // pollution skip dropped the wedgeRectCallout column bodies
        // (y 67.2→335.9), shrinking the bbox to the 67.2pt-tall header row
        // and fit-scaling the diagram ×1.21 with a +127pt downshift. Now
        // that the preset is renderable it must set the fit bbox again.
        var header = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""853198"" y=""0""/>
      <a:ext cx=""1343775"" cy=""853198""/>
    </a:xfrm>
    <a:prstGeom prst=""rect"">
      <a:avLst/>
    </a:prstGeom>
  </dsp:spPr>
</dsp:sp>";
        var body = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""853198"" y=""853198""/>
      <a:ext cx=""1343775"" cy=""3412792""/>
    </a:xfrm>
    <a:prstGeom prst=""wedgeRectCallout"">
      <a:avLst>
        <a:gd name=""adj1"" fmla=""val 0""/>
        <a:gd name=""adj2"" fmla=""val 0""/>
      </a:avLst>
    </a:prstGeom>
  </dsp:spPr>
</dsp:sp>";

        var bounds = SmartArtDrawingExtractor.ComputeBoundingBox(
            new[] { ParseXml(header), ParseXml(body) });

        Assert.NotNull(bounds);
        Assert.Equal(0.0, bounds.Value.MinY, precision: 6);
        // (853198 + 3412792) EMU = 4265990 EMU — the full column height.
        Assert.Equal(4265990.0 / 12700.0, bounds.Value.Height, precision: 6);
    }

    [Fact]
    public void ComputeBoundingBoxes_RotatedShape_ReturnsBlindAndAwareBounds()
    {
        // rot=90° about center (120,50) turns the 40x100pt rect at (100,0)
        // into a 100x40pt footprint at (70,30): the blind union keeps the
        // raw off/ext rect, the aware union the rotated footprint.
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm rot=""5400000"">
      <a:off x=""1270000"" y=""0""/>
      <a:ext cx=""508000"" cy=""1270000""/>
    </a:xfrm>
    <a:prstGeom prst=""rect"">
      <a:avLst/>
    </a:prstGeom>
  </dsp:spPr>
</dsp:sp>";

        var bounds = SmartArtDrawingExtractor.ComputeBoundingBoxes(new[] { ParseXml(xml) });

        Assert.NotNull(bounds.Blind);
        Assert.Equal(100.0, bounds.Blind.Value.MinX, precision: 6);
        Assert.Equal(0.0, bounds.Blind.Value.MinY, precision: 6);
        Assert.Equal(40.0, bounds.Blind.Value.Width, precision: 6);
        Assert.Equal(100.0, bounds.Blind.Value.Height, precision: 6);

        Assert.NotNull(bounds.Aware);
        Assert.Equal(70.0, bounds.Aware.Value.MinX, precision: 6);
        Assert.Equal(30.0, bounds.Aware.Value.MinY, precision: 6);
        Assert.Equal(100.0, bounds.Aware.Value.Width, precision: 6);
        Assert.Equal(40.0, bounds.Aware.Value.Height, precision: 6);
    }

    [Fact]
    public void ComputeBoundingBoxes_UnrotatedShapes_BlindEqualsAware()
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
  </dsp:spPr>
</dsp:sp>";

        var bounds = SmartArtDrawingExtractor.ComputeBoundingBoxes(new[] { ParseXml(xml) });

        Assert.NotNull(bounds.Blind);
        Assert.NotNull(bounds.Aware);
        Assert.Equal(bounds.Blind.Value, bounds.Aware.Value);
    }

    [Fact]
    public void ComputeFrameFit_DualFit_PicksBlindWhenCloserToIdentity()
    {
        // -b1 §4 (slide 152, drawing149): the rotation-blind
        // bbox (14.3,16.1,625.7,319.9) fits the 640x335.9 frame at scale
        // 1.0229; the rotation-aware bbox (53.4,16.1,586.6,319.9) at 1.05.
        // A cached dsp:drawing is authored in frame coordinates, so the fit
        // closer to identity (1.0229) is the better approximation — the
        // dual-fit rule must pick it (restoring the pre-regression render).
        var blind = (MinX: 14.3, MinY: 16.1, Width: 625.7, Height: 319.9);
        var aware = (MinX: 53.4, MinY: 16.1, Width: 586.6, Height: 319.9);
        var frame = (X: 235.7, Y: 122.6, Width: 640.0, Height: 335.9);

        var fit = SmartArtDrawingExtractor.ComputeFrameFit(blind, aware, frame);

        var blindOnly = SmartArtDrawingExtractor.ComputeFrameFit(blind, frame);
        Assert.Equal(640.0 / 625.7, fit.ScaleX, precision: 6);
        Assert.Equal(blindOnly.ScaleX, fit.ScaleX, precision: 9);
        Assert.Equal(blindOnly.FrameX, fit.FrameX, precision: 9);
        Assert.Equal(blindOnly.FrameY, fit.FrameY, precision: 9);
    }

    [Fact]
    public void ComputeFrameFit_DualFit_PicksAwareWhenBlindInflates()
    {
        // -130: the rotation-blind bbox inflated the content height
        // by 45% (fit scale 0.688) while the rotation-aware bbox fits at
        // ~1.0 — the batch-1 win must be kept: dual-fit picks the aware fit.
        var frame = (X: 235.7, Y: 122.6, Width: 640.0, Height: 335.9);
        var aware = (MinX: 0.0, MinY: 0.0, Width: 640.0, Height: 335.9);
        var blind = (MinX: 0.0, MinY: 0.0, Width: 640.0, Height: 335.9 / 0.688);

        var fit = SmartArtDrawingExtractor.ComputeFrameFit(blind, aware, frame);

        Assert.Equal(1.0, fit.ScaleX, precision: 6);
        Assert.Equal(frame.X, fit.FrameX, precision: 6);
        Assert.Equal(frame.Y, fit.FrameY, precision: 6);
    }

    [Fact]
    public void ComputeFrameFit_DualFit_ScaleTie_PicksSmallerIdentityOffset()
    {
        // Slide-71 pattern (drawing74): the rot-90/270 labels' raw rects add
        // phantom width their rotated ink never occupies, so the blind and
        // aware bboxes tie on scale (both height-limited to 335.9/335.7).
        // The aware fit is the identity transform (centerOffset ≈ minX·scale);
        // the blind fit would shift the real ink +13.5pt. The tie-break must
        // compare the identity-translation error, not the raw centering
        // offset (which misleadingly favours the blind fit, 120 < 133.7).
        var frame = (X: 235.7, Y: 122.6, Width: 640.0, Height: 335.9);
        var blind = (MinX: 106.5, MinY: 0.1, Width: 399.7, Height: 335.7);
        var aware = (MinX: 133.8, MinY: 0.1, Width: 372.4, Height: 335.7);

        var fit = SmartArtDrawingExtractor.ComputeFrameFit(blind, aware, frame);

        var awareOnly = SmartArtDrawingExtractor.ComputeFrameFit(aware, frame);
        Assert.Equal(awareOnly.ScaleX, fit.ScaleX, precision: 9);
        Assert.Equal(awareOnly.FrameX, fit.FrameX, precision: 9);
        Assert.Equal(awareOnly.FrameY, fit.FrameY, precision: 9);
    }

    [Fact]
    public void ComputeFrameFit_DualFit_IdenticalBounds_MatchesSingleBoundsFit()
    {
        // Drawings without rotated shapes have blind == aware: the dual-fit
        // rule is a no-op and must reproduce the single-bounds fit exactly.
        var bounds = (MinX: 10.0, MinY: 20.0, Width: 100.0, Height: 50.0);
        var frame = (X: 40.0, Y: 80.0, Width: 200.0, Height: 100.0);

        var dual = SmartArtDrawingExtractor.ComputeFrameFit(bounds, bounds, frame);
        var single = SmartArtDrawingExtractor.ComputeFrameFit(bounds, frame);

        Assert.Equal(single.ScaleX, dual.ScaleX, precision: 9);
        Assert.Equal(single.ScaleY, dual.ScaleY, precision: 9);
        Assert.Equal(single.FrameX, dual.FrameX, precision: 9);
        Assert.Equal(single.FrameY, dual.FrameY, precision: 9);
    }

    [Fact]
    public void Extract_Chevron_WideShape_NotchDepthIsFractionOfMinSide()
    {
        // Slide 56 (drawing59.xml): 141.6x56.6pt chevrons with an empty avLst.
        // ECMA-376 chevron: notch depth dx1 = ss*adj/100000 (default adj 50000)
        // = 0.5*56.6 = 28.31pt => f = 0.20 of the WIDTH — the static polygon's
        // hardcoded 0.5-of-width notch is 2.5x too deep (-025 §4).
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""1797843"" cy=""719137""/>
    </a:xfrm>
    <a:prstGeom prst=""chevron"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""16A085""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 141.6, shapeH: 56.6);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        Assert.Equal(6, result.Shape.Points.Count);
        Assert.Equal((0, 0), result.Shape.Points[0]);
        Assert.Equal(0.8, result.Shape.Points[1].X, precision: 3);
        Assert.Equal(0.0, result.Shape.Points[1].Y, precision: 6);
        Assert.Equal((1, 0.5), result.Shape.Points[2]);
        Assert.Equal(0.8, result.Shape.Points[3].X, precision: 3);
        Assert.Equal(1.0, result.Shape.Points[3].Y, precision: 6);
        Assert.Equal((0, 1), result.Shape.Points[4]);
        Assert.Equal(0.2, result.Shape.Points[5].X, precision: 3);
        Assert.Equal(0.5, result.Shape.Points[5].Y, precision: 6);
    }

    [Fact]
    public void Extract_Chevron_HonorsAdjustmentValue()
    {
        // Square 100x100pt chevron with adj = 25000 (slide 58's cached value):
        // notch depth = 0.25*100 = 25pt => f = 0.25.
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""1270000"" cy=""1270000""/>
    </a:xfrm>
    <a:prstGeom prst=""chevron"">
      <a:avLst>
        <a:gd name=""adj"" fmla=""val 25000""/>
      </a:avLst>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""16A085""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 100, shapeH: 100);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal(6, result.Shape.Points.Count);
        Assert.Equal(0.75, result.Shape.Points[1].X, precision: 6);
        Assert.Equal(0.25, result.Shape.Points[5].X, precision: 6);
    }

    [Fact]
    public void Extract_Chevron_SquareShapeDefaultAdj_MatchesLegacyStaticOutline()
    {
        // On a square shape the ECMA default (adj = 50000 => dx1 = 0.5*ss)
        // reproduces the legacy static 0.5-of-width table exactly — the
        // per-shape computation is a strict generalisation.
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""1270000"" cy=""1270000""/>
    </a:xfrm>
    <a:prstGeom prst=""chevron"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""16A085""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 100, shapeH: 100);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal(6, result.Shape.Points.Count);
        Assert.Equal((0, 0), result.Shape.Points[0]);
        Assert.Equal((0.5, 0), result.Shape.Points[1]);
        Assert.Equal((1, 0.5), result.Shape.Points[2]);
        Assert.Equal((0.5, 1), result.Shape.Points[3]);
        Assert.Equal((0, 1), result.Shape.Points[4]);
        Assert.Equal((0.5, 0.5), result.Shape.Points[5]);
    }

    [Fact]
    public void Extract_Round2DiagRect_Slide113Adjustments_RoundsTopRightAndBottomLeftOnly()
    {
        // Slide 113 (drawing113.xml): round2DiagRect with adj1=0, adj2=16670 on a
        // 412.3x221.7pt shape. adj1 rounds top-left + bottom-right (radius 0 here
        // => sharp), adj2 rounds top-right + bottom-left with radius
        // a = ss*16670/100000 = 36.957pt (a/w = 0.0896, a/h = 0.1667).
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""5235649"" cy=""2815553""/>
    </a:xfrm>
    <a:prstGeom prst=""round2DiagRect"">
      <a:avLst>
        <a:gd name=""adj1"" fmla=""val 0""/>
        <a:gd name=""adj2"" fmla=""val 16670""/>
      </a:avLst>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""16A085""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 412.3, shapeH: 221.7);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        Assert.Equal("polygon", result.Shape.ShapeType);
        var points = result.Shape.Points;
        Assert.True(points.Count > 50); // two 90-degree corner arcs flattened in 2-degree steps

        // Sharp diagonal pair (adj1 = 0): the outline starts at the top-left
        // corner and reaches the bottom-right corner exactly.
        Assert.Equal((0, 0), points[0]);
        Assert.Contains(points, p => Math.Abs(p.X - 1) < 1e-6 && Math.Abs(p.Y - 1) < 1e-6);

        // Rounded top-right (adj2): the top edge ends at x2/w = 1 - a/w and the
        // corner arc lands on the right edge at y = a/h.
        Assert.Equal(0.9104, points[1].X, precision: 3);
        Assert.Equal(0.0, points[1].Y, precision: 6);
        Assert.Contains(points, p => Math.Abs(p.X - 1) < 1e-6 && Math.Abs(p.Y - 0.1667) < 1e-3);

        // Rounded bottom-left (adj2): the bottom edge starts at x = a/w and the
        // corner arc lands on the left edge at y = 1 - a/h.
        Assert.Contains(points, p => Math.Abs(p.X - 0.0896) < 1e-3 && Math.Abs(p.Y - 1) < 1e-6);
        Assert.Contains(points, p => Math.Abs(p.X) < 1e-6 && Math.Abs(p.Y - 0.8333) < 1e-3);
    }

    [Fact]
    public void Extract_Round2DiagRect_DefaultAdjustments_RoundsTopLeftAndBottomRightOnly()
    {
        // ECMA-376 defaults (adj1 = 16667, adj2 = 0): on a square 100x100pt shape
        // the top-left and bottom-right corners are rounded with radius ss/6;
        // the top-right and bottom-left corners stay sharp.
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""1270000"" cy=""1270000""/>
    </a:xfrm>
    <a:prstGeom prst=""round2DiagRect"">
      <a:avLst/>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val=""16A085""/>
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </dsp:spPr>
</dsp:sp>";

        var element = ParseXml(xml);

        var result = SmartArtDrawingExtractor.TryExtractShape(
            element, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 100, shapeH: 100);

        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        var points = result.Shape.Points;

        // moveTo (x1, t) with x1 = 100*16667/100000 = 16.667pt.
        Assert.Equal(0.16667, points[0].X, precision: 4);
        Assert.Equal(0.0, points[0].Y, precision: 4);

        // Sharp diagonal pair (adj2 = 0): top-right and bottom-left corners.
        Assert.Contains(points, p => Math.Abs(p.X - 1) < 1e-6 && Math.Abs(p.Y) < 1e-6);
        Assert.Contains(points, p => Math.Abs(p.X) < 1e-6 && Math.Abs(p.Y - 1) < 1e-6);

        // Rounded top-left / bottom-right (adj1): the right edge runs from the
        // sharp top-right corner down to y1 = 1 - 0.16667, then curves onto the
        // bottom edge at x = 1 - 0.16667.
        Assert.Contains(points, p => Math.Abs(p.X - 1) < 1e-6 && Math.Abs(p.Y - 0.83333) < 1e-4);
        Assert.Contains(points, p => Math.Abs(p.X - 0.83333) < 1e-4 && Math.Abs(p.Y - 1) < 1e-6);
    }

    [Fact]
    public void Extract_Donut_PreservesTransparentInnerContour()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}""><dsp:spPr>
  <a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""1270000"" cy=""1270000""/></a:xfrm>
  <a:prstGeom prst=""donut""><a:avLst><a:gd name=""adj"" fmla=""val 10000""/></a:avLst></a:prstGeom>
  <a:solidFill><a:srgbClr val=""16A085""/></a:solidFill>
</dsp:spPr></dsp:sp>";
        var result = SmartArtDrawingExtractor.TryExtractShape(ParseXml(xml), 0, 0, 1, 1, 0, 0, 100, 100);
        Assert.NotNull(result?.Shape);
        Assert.Equal("polygon", result.Shape!.ShapeType);
        Assert.Equal(2, result.Shape.Subpaths.Count);
        Assert.Equal(64, result.Shape.Subpaths[0].Count);
        Assert.Equal(64, result.Shape.Subpaths[1].Count);
        Assert.Equal(0.9, result.Shape.Subpaths[1].Max(p => p.X), precision: 6);
    }

    [Fact]
    public void Extract_Funnel_ReturnsTaperedSilhouette()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}""><dsp:spPr>
  <a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""2540000"" cy=""3810000""/></a:xfrm>
  <a:prstGeom prst=""funnel""><a:avLst/></a:prstGeom>
  <a:solidFill><a:srgbClr val=""D9E2F3""/></a:solidFill>
</dsp:spPr></dsp:sp>";
        var result = SmartArtDrawingExtractor.TryExtractShape(ParseXml(xml), 0, 0, 1, 1, 0, 0, 200, 300);
        Assert.NotNull(result?.Shape);
        Assert.Equal("polygon", result.Shape!.ShapeType);
        Assert.True(result.Shape.Points.Count > 50);
        Assert.Equal(0.0, result.Shape.Points.Min(p => p.X), precision: 6);
        Assert.Equal(1.0, result.Shape.Points.Max(p => p.X), precision: 6);
        Assert.Contains(result.Shape.Points, p => p.X is > 0.3 and < 0.7 && p.Y > 0.85);
    }

    [Fact]
    public void Extract_Arc_PreservesOpenCurvedConnectorPath()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}""><dsp:spPr>
  <a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""1270000"" cy=""1270000""/></a:xfrm>
  <a:prstGeom prst=""arc""><a:avLst><a:gd name=""adj1"" fmla=""val 0""/><a:gd name=""adj2"" fmla=""val 9000000""/></a:avLst></a:prstGeom>
  <a:noFill/><a:ln w=""9525""><a:solidFill><a:srgbClr val=""16A085""/></a:solidFill></a:ln>
</dsp:spPr></dsp:sp>";
        var result = SmartArtDrawingExtractor.TryExtractShape(ParseXml(xml), 0, 0, 1, 1, 0, 0, 100, 100);
        Assert.NotNull(result?.Shape);
        Assert.Equal("path", result.Shape!.ShapeType);
        Assert.Single(result.Shape.Subpaths);
        Assert.True(result.Shape.Subpaths[0].Count > 15);
        Assert.DoesNotContain(true, result.Shape.ClosedSubpaths);
    }

    [Fact]
    public void Extract_ExplicitNoFill_WinsOverSiblingGradient()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}""><dsp:spPr>
  <a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""1270000"" cy=""1270000""/></a:xfrm>
  <a:prstGeom prst=""rect""><a:avLst/></a:prstGeom><a:noFill/>
  <a:gradFill><a:gsLst><a:gs pos=""0""><a:srgbClr val=""000000""/></a:gs><a:gs pos=""100000""><a:srgbClr val=""FFFFFF""/></a:gs></a:gsLst><a:lin ang=""0""/></a:gradFill>
</dsp:spPr></dsp:sp>";
        var result = SmartArtDrawingExtractor.TryExtractShape(ParseXml(xml), 0, 0, 1, 1, 0, 0, 100, 100);
        Assert.NotNull(result?.Shape);
        Assert.Null(result.Shape!.FillGradient);
        Assert.True(string.IsNullOrEmpty(result.Shape.FillColor));
    }

    [Fact]
    public void Extract_AlphaModifiers_AreAppliedInDocumentOrder()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}""><dsp:spPr>
  <a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""1270000"" cy=""1270000""/></a:xfrm>
  <a:prstGeom prst=""rect""><a:avLst/></a:prstGeom>
  <a:solidFill><a:srgbClr val=""FFFFFF""><a:alpha val=""50000""/><a:alphaMod val=""50000""/><a:alphaOff val=""10000""/></a:srgbClr></a:solidFill>
</dsp:spPr></dsp:sp>";
        var result = SmartArtDrawingExtractor.TryExtractShape(ParseXml(xml), 0, 0, 1, 1, 0, 0, 100, 100);
        Assert.Equal("#FFFFFF59", result?.Shape?.FillColor);
    }

    [Fact]
    public void Extract_SystemColorLastClr_UsesPortableSixDigitFallback()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}""><dsp:spPr>
  <a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""1270000"" cy=""1270000""/></a:xfrm>
  <a:prstGeom prst=""rect""><a:avLst/></a:prstGeom>
  <a:solidFill><a:sysClr val=""window"" lastClr=""FFFFFF""/></a:solidFill>
</dsp:spPr></dsp:sp>";
        var result = SmartArtDrawingExtractor.TryExtractShape(ParseXml(xml), 0, 0, 1, 1, 0, 0, 100, 100);
        Assert.Equal("#FFFFFF", result?.Shape?.FillColor);
    }

    [Theory]
    [InlineData("<a:scrgbClr r=\"100000\" g=\"50000\" b=\"0\"/>", "#FF8000")]
    [InlineData("<a:hslClr hue=\"0\" sat=\"100000\" lum=\"50000\"/>", "#FF0000")]
    [InlineData("<a:prstClr val=\"tomato\"/>", "#FF6347")]
    public void Extract_AdditionalDrawingMlColorForms_ResolveSolidFill(string colorXml, string expected)
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}""><dsp:spPr>
  <a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""1270000"" cy=""1270000""/></a:xfrm>
  <a:prstGeom prst=""rect""><a:avLst/></a:prstGeom><a:solidFill>{colorXml}</a:solidFill>
</dsp:spPr></dsp:sp>";
        var result = SmartArtDrawingExtractor.TryExtractShape(ParseXml(xml), 0, 0, 1, 1, 0, 0, 100, 100);
        Assert.Equal(expected, result?.Shape?.FillColor);
    }

    [Fact]
    public void Extract_GradientWithScrgbAndHslStops_PreservesBothStops()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}""><dsp:spPr>
  <a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""1270000"" cy=""1270000""/></a:xfrm>
  <a:prstGeom prst=""rect""><a:avLst/></a:prstGeom>
  <a:gradFill><a:gsLst>
    <a:gs pos=""0""><a:scrgbClr r=""100000"" g=""0"" b=""0""/></a:gs>
    <a:gs pos=""100000""><a:hslClr hue=""7200000"" sat=""100000"" lum=""50000""/></a:gs>
  </a:gsLst><a:lin ang=""0""/></a:gradFill>
</dsp:spPr></dsp:sp>";
        var result = SmartArtDrawingExtractor.TryExtractShape(ParseXml(xml), 0, 0, 1, 1, 0, 0, 100, 100);
        var gradient = Assert.IsType<TypstGradientFill>(result?.Shape?.FillGradient);
        Assert.Equal("#FF0000", gradient.Stops[0].Color);
        Assert.Equal("#00FF00", gradient.Stops[1].Color);
    }

    [Fact]
    public void Extract_OuterShadowAndBevel_PreservesEffectValues()
    {
        var xml = $@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}""><dsp:spPr>
  <a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""1270000"" cy=""1270000""/></a:xfrm>
  <a:prstGeom prst=""rect""><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val=""16A085""/></a:solidFill>
  <a:ln w=""12700""><a:solidFill><a:srgbClr val=""FFFFFF""><a:alpha val=""50000""/></a:srgbClr></a:solidFill></a:ln>
  <a:effectLst><a:outerShdw blurRad=""40000"" dist=""23000"" dir=""5400000"" rotWithShape=""0""><a:srgbClr val=""000000""><a:alpha val=""35000""/></a:srgbClr></a:outerShdw></a:effectLst>
  <a:sp3d><a:bevelT w=""63500"" h=""25400""/></a:sp3d>
</dsp:spPr></dsp:sp>";
        var result = SmartArtDrawingExtractor.TryExtractShape(ParseXml(xml), 0, 0, 1, 1, 0, 0, 100, 100);
        Assert.NotNull(result?.Shape);
        Assert.Equal("#FFFFFF7F", result.Shape!.StrokeColor);
        Assert.Equal("#00000059", result.Shape.Shadow!.Color);
        Assert.False(result.Shape.Shadow.RotateWithShape);
        Assert.Equal(63500 / 12700.0, result.Shape.Bevel!.TopWidth, precision: 3);
    }

    [Fact]
    public void ComputeCachedDrawingFit_UnrelatedCoordinateSpaceUsesUniformFallback()
    {
        var blind = (MinX: -2500.0, MinY: -1500.0, Width: 5000.0, Height: 3000.0);
        var frame = (X: 100.0, Y: 50.0, Width: 640.0, Height: 335.9);
        var fit = SmartArtDrawingExtractor.ComputeCachedDrawingFit(blind, blind, frame);
        Assert.Equal(Math.Min(frame.Width / blind.Width, frame.Height / blind.Height), fit.ScaleX, precision: 9);
        Assert.Equal(fit.ScaleX, fit.ScaleY, precision: 9);
    }

    [Fact]
    public void ComputeCachedDrawingFit_FrameSpaceCachePreservesNativeTransform()
    {
        var bounds = (MinX: 17.98, MinY: 8.95, Width: 604.03, Height: 318.01);
        var frame = (X: 235.75, Y: 122.58, Width: 640.0, Height: 335.9);
        var fit = SmartArtDrawingExtractor.ComputeCachedDrawingFit(bounds, bounds, frame);
        Assert.Equal(1.0, fit.ScaleX, precision: 9);
        Assert.Equal(1.0, fit.ScaleY, precision: 9);
        Assert.Equal(frame.X, fit.FrameX, precision: 9);
        Assert.Equal(frame.Y, fit.FrameY, precision: 9);
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
