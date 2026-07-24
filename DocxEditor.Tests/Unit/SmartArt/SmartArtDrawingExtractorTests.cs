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
        Assert.Equal("C00000", result.Shape.FillColor);
        Assert.Equal("FFFFFF", result.Shape.StrokeColor);
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
        Assert.Equal("ED7D31", result.Shape.FillColor);
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
        Assert.Equal("4472C4", result.Shape.FillColor);
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
