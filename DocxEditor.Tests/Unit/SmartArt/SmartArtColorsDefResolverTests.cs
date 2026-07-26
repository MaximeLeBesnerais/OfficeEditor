using System.Xml.Linq;
using DocumentFormat.OpenXml;
using PptxEditor.Core.Converters.SmartArt;
using PptxEditor.Core.Models;

namespace DocxEditor.Tests.Unit.SmartArt;

/// <summary>
/// Tests for the colorsDef fill-precedence rule (INV-regressions-b1 §5-6):
/// when a cached dsp:sp gradFill conflicts with a flat dgm:colorsDef styleLbl
/// mapping (plain schemeClr/srgbClr, no transforms) and the quickStyle carries
/// no gradient, the cached gradient is truncated to its visible window
/// (cached pos 44-100% — what PowerPoint's relayout displays); in every other
/// case the cached spPr keeps precedence untouched.
/// </summary>
public sealed class SmartArtColorsDefResolverTests
{
    private const string DgmNs = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
    private const string DspNs = "http://schemas.microsoft.com/office/drawing/2008/diagram";
    private const string ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    // Presentation-point modelIds used by the synthetic data part.
    private const string FlatNode = "{00000000-0000-0000-0000-000000000001}";      // node1, idx 2/5
    private const string TransformedNode = "{00000000-0000-0000-0000-000000000002}"; // shadeNode (shade transform)
    private const string SpanNode = "{00000000-0000-0000-0000-000000000003}";        // spanNode (meth=span)
    private const string SrgbNode = "{00000000-0000-0000-0000-000000000004}";        // srgbNode (flat literal)
    private const string QuickStyleGradientNode = "{00000000-0000-0000-0000-000000000005}"; // gradLbl (qs gradFill)
    private const string DefaultLabelNode = "{00000000-0000-0000-0000-000000000006}";  // no presStyleLbl → node0
    private const string MissingLabelNode = "{00000000-0000-0000-0000-000000000007}";  // unknown styleLbl
    private const string UnknownModelId = "{00000000-0000-0000-0000-000000000099}";    // not in data part

    private const string ColorsDefXml = $@"
<dgm:colorsDef xmlns:dgm=""{DgmNs}"" xmlns:a=""{ANs}"" uniqueId=""test/colors"">
  <dgm:styleLbl name=""node0"">
    <dgm:fillClrLst meth=""repeat""><a:schemeClr val=""accent1""/></dgm:fillClrLst>
  </dgm:styleLbl>
  <dgm:styleLbl name=""node1"">
    <dgm:fillClrLst meth=""repeat""><a:schemeClr val=""accent2""/><a:schemeClr val=""accent3""/><a:schemeClr val=""accent4""/><a:schemeClr val=""accent5""/><a:schemeClr val=""accent6""/></dgm:fillClrLst>
  </dgm:styleLbl>
  <dgm:styleLbl name=""shadeNode"">
    <dgm:fillClrLst meth=""repeat""><a:schemeClr val=""accent2""><a:shade val=""75000""/></a:schemeClr></dgm:fillClrLst>
  </dgm:styleLbl>
  <dgm:styleLbl name=""spanNode"">
    <dgm:fillClrLst meth=""span""><a:schemeClr val=""accent1""/><a:schemeClr val=""accent2""/></dgm:fillClrLst>
  </dgm:styleLbl>
  <dgm:styleLbl name=""srgbNode"">
    <dgm:fillClrLst meth=""repeat""><a:srgbClr val=""123456""/></dgm:fillClrLst>
  </dgm:styleLbl>
  <dgm:styleLbl name=""gradLbl"">
    <dgm:fillClrLst meth=""repeat""><a:schemeClr val=""accent3""/></dgm:fillClrLst>
  </dgm:styleLbl>
</dgm:colorsDef>";

    private const string DataModelXml = $@"
<dgm:dataModel xmlns:dgm=""{DgmNs}"" xmlns:a=""{ANs}"">
  <dgm:ptLst>
    <dgm:pt modelId=""{FlatNode}"" type=""pres""><dgm:prSet presStyleLbl=""node1"" presStyleIdx=""2"" presStyleCnt=""5""/></dgm:pt>
    <dgm:pt modelId=""{TransformedNode}"" type=""pres""><dgm:prSet presStyleLbl=""shadeNode"" presStyleIdx=""0"" presStyleCnt=""1""/></dgm:pt>
    <dgm:pt modelId=""{SpanNode}"" type=""pres""><dgm:prSet presStyleLbl=""spanNode""/></dgm:pt>
    <dgm:pt modelId=""{SrgbNode}"" type=""pres""><dgm:prSet presStyleLbl=""srgbNode""/></dgm:pt>
    <dgm:pt modelId=""{QuickStyleGradientNode}"" type=""pres""><dgm:prSet presStyleLbl=""gradLbl""/></dgm:pt>
    <dgm:pt modelId=""{DefaultLabelNode}"" type=""pres""><dgm:prSet/></dgm:pt>
    <dgm:pt modelId=""{MissingLabelNode}"" type=""pres""><dgm:prSet presStyleLbl=""missingLbl""/></dgm:pt>
  </dgm:ptLst>
</dgm:dataModel>";

    private const string StyleDefXml = $@"
<dgm:styleDef xmlns:dgm=""{DgmNs}"" xmlns:a=""{ANs}"" uniqueId=""test/quickstyle"">
  <dgm:styleLbl name=""node0"">
    <dgm:style><a:fillRef idx=""3""><a:scrgbClr r=""0"" g=""0"" b=""0""/></a:fillRef><a:effectRef idx=""2""><a:scrgbClr r=""0"" g=""0"" b=""0""/></a:effectRef></dgm:style>
  </dgm:styleLbl>
  <dgm:styleLbl name=""node1"">
    <dgm:style><a:fillRef idx=""3""><a:scrgbClr r=""0"" g=""0"" b=""0""/></a:fillRef><a:effectRef idx=""2""><a:scrgbClr r=""0"" g=""0"" b=""0""/></a:effectRef></dgm:style>
  </dgm:styleLbl>
  <dgm:styleLbl name=""gradLbl"">
    <dgm:spPr><a:gradFill><a:gsLst><a:gs pos=""0""><a:schemeClr val=""phClr""/></a:gs><a:gs pos=""100000""><a:schemeClr val=""phClr""><a:shade val=""50000""/></a:schemeClr></a:gs></a:gsLst><a:lin ang=""16200000"" scaled=""0""/></a:gradFill></dgm:spPr>
    <dgm:style><a:fillRef idx=""3""><a:scrgbClr r=""0"" g=""0"" b=""0""/></a:fillRef></dgm:style>
  </dgm:styleLbl>
</dgm:styleDef>";

    private static SmartArtColorsDefResolver.ColorsDefContext CreateContext()
    {
        var context = SmartArtColorsDefResolver.CreateContext(
            ParseXml(ColorsDefXml), ParseXml(DataModelXml), ParseXml(StyleDefXml));
        Assert.NotNull(context);
        return context;
    }

    private static OpenXmlElement GradientShape(string modelId)
    {
        // Corpus pattern (drawing35): cached linear gradient with heavy
        // shade/satMod stops — conflicts with the flat colorsDef mapping.
        return ParseXml($@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"" modelId=""{modelId}"">
  <dsp:spPr>
    <a:xfrm>
      <a:off x=""892"" y=""305395""/>
      <a:ext cx=""1904255"" cy=""1142553""/>
    </a:xfrm>
    <a:prstGeom prst=""rect"">
      <a:avLst/>
    </a:prstGeom>
    <a:gradFill rotWithShape=""0"">
      <a:gsLst>
        <a:gs pos=""0""><a:schemeClr val=""accent4""><a:shade val=""51000""/><a:satMod val=""130000""/></a:schemeClr></a:gs>
        <a:gs pos=""80000""><a:schemeClr val=""accent4""><a:shade val=""93000""/><a:satMod val=""130000""/></a:schemeClr></a:gs>
        <a:gs pos=""100000""><a:schemeClr val=""accent4""><a:shade val=""94000""/><a:satMod val=""135000""/></a:schemeClr></a:gs>
      </a:gsLst>
      <a:lin ang=""16200000"" scaled=""0""/>
    </a:gradFill>
  </dsp:spPr>
</dsp:sp>");
    }

    private static TypstElement ExtractWithContext(OpenXmlElement shape,
        SmartArtColorsDefResolver.ColorsDefContext? context, string? modelId)
    {
        var result = SmartArtDrawingExtractor.TryExtractShape(
            shape, offX: 0, offY: 0, scaleX: 1.0, scaleY: 1.0,
            frameX: 0, frameY: 0, shapeW: 150, shapeH: 90,
            schemeColors: null, modelId: modelId, colorsDefContext: context);
        Assert.NotNull(result);
        Assert.NotNull(result.Shape);
        return result;
    }

    // ---------- resolver-level tests ----------

    [Fact]
    public void Resolver_FlatColorsDefEntry_ResolvesRepeatIndex()
    {
        var context = CreateContext();

        // node1 repeat list accent2..6, presStyleIdx=2 → accent4 (static fallback).
        Assert.Equal("#FFC000", context.TryResolveFlatFill(FlatNode, schemeColors: null));
    }

    [Fact]
    public void Resolver_FlatColorsDefEntry_UsesThemeSchemeMap()
    {
        var context = CreateContext();
        var schemeColors = new Dictionary<string, string> { ["accent4"] = "16A085" };

        Assert.Equal("#16A085", context.TryResolveFlatFill(FlatNode, schemeColors));
    }

    [Fact]
    public void Resolver_RepeatIndexWrapsAroundList()
    {
        var context = CreateContext();

        // presStyleIdx=2 with a 5-entry list → entry 2; DefaultLabelNode has no
        // presStyleIdx → index 0 → node0's single entry (accent1 fallback).
        Assert.Equal("#4472C4", context.TryResolveFlatFill(DefaultLabelNode, schemeColors: null));
    }

    [Fact]
    public void Resolver_FlatSrgbEntry_ReturnsLiteralHex()
    {
        var context = CreateContext();

        Assert.Equal("#123456", context.TryResolveFlatFill(SrgbNode, schemeColors: null));
    }

    [Fact]
    public void Resolver_ColorsDefWithTransforms_KeepsCachedPrecedence()
    {
        var context = CreateContext();

        Assert.Null(context.TryResolveFlatFill(TransformedNode, schemeColors: null));
    }

    [Fact]
    public void Resolver_SpanMethod_KeepsCachedPrecedence()
    {
        var context = CreateContext();

        Assert.Null(context.TryResolveFlatFill(SpanNode, schemeColors: null));
    }

    [Fact]
    public void Resolver_QuickStyleGradient_KeepsCachedPrecedence()
    {
        var context = CreateContext();

        Assert.Null(context.TryResolveFlatFill(QuickStyleGradientNode, schemeColors: null));
    }

    [Fact]
    public void Resolver_UnknownStyleLabelOrModelId_KeepsCachedPrecedence()
    {
        var context = CreateContext();

        Assert.Null(context.TryResolveFlatFill(MissingLabelNode, schemeColors: null));
        Assert.Null(context.TryResolveFlatFill(UnknownModelId, schemeColors: null));
        Assert.Null(context.TryResolveFlatFill(modelId: null, schemeColors: null));
    }

    [Fact]
    public void CreateContext_MissingParts_ReturnsNull()
    {
        Assert.Null(SmartArtColorsDefResolver.CreateContext(colorsDefRoot: null, ParseXml(DataModelXml), styleDefRoot: null));
        Assert.Null(SmartArtColorsDefResolver.CreateContext(ParseXml(ColorsDefXml), dataModelRoot: null, styleDefRoot: null));
    }

    // ---------- end-to-end precedence through TryExtractShape ----------

    [Fact]
    public void Extract_FlatColorsDefConflict_GradientTruncatedToVisibleWindow()
    {
        var context = CreateContext();

        var result = ExtractWithContext(GradientShape(FlatNode), context, FlatNode);

        // Clear-conflict case: the cached gradient is truncated to its visible
        // window — cached pos 100% at offset 1, cached pos ~44% at offset 0
        // (PowerPoint relayout shows only the top slice of the scaled="0"
        // gradient vector). The middle stop (pos 80%) lands at (80-44)/56.
        Assert.True(string.IsNullOrEmpty(result.Shape!.FillColor));
        var gradient = Assert.IsType<TypstGradientFill>(result.Shape.FillGradient);
        Assert.Equal(270.0, gradient.Angle);
        Assert.Equal(3, gradient.Stops.Count);
        Assert.Equal(0.0, gradient.Stops[0].Offset);
        Assert.Equal((80.0 - 44.0) / 56.0, gradient.Stops[1].Offset, precision: 6);
        Assert.Equal(1.0, gradient.Stops[2].Offset);

        // Window edges: the far edge keeps the cached pos-100 colour; the near
        // edge is the cached pos-44 colour (interpolated 55% from stop 0 to
        // stop 80), not the cached pos-0 colour.
        var untruncated = ExtractWithContext(GradientShape(TransformedNode), context, TransformedNode)
            .Shape!.FillGradient!;
        Assert.Equal(untruncated.Stops[2].Color, gradient.Stops[2].Color);
        Assert.NotEqual(untruncated.Stops[0].Color, gradient.Stops[0].Color);
        var stop0 = ParseHex(untruncated.Stops[0].Color);
        var stop80 = ParseHex(untruncated.Stops[1].Color);
        var expectedNear = LerpHex(stop0, stop80, 44.0 / 80.0);
        Assert.Equal(expectedNear, gradient.Stops[0].Color);
    }

    private static (int R, int G, int B) ParseHex(string hex)
    {
        var h = hex.TrimStart('#');
        return (Convert.ToInt32(h[..2], 16), Convert.ToInt32(h[2..4], 16), Convert.ToInt32(h[4..6], 16));
    }

    private static string LerpHex((int R, int G, int B) a, (int R, int G, int B) b, double t)
    {
        return $"#{(int)Math.Round(a.R + (b.R - a.R) * t):X2}{(int)Math.Round(a.G + (b.G - a.G) * t):X2}{(int)Math.Round(a.B + (b.B - a.B) * t):X2}";
    }

    [Fact]
    public void Extract_ColorsDefWithTransforms_CachedGradientWins()
    {
        var context = CreateContext();

        var result = ExtractWithContext(GradientShape(TransformedNode), context, TransformedNode);

        Assert.True(string.IsNullOrEmpty(result.Shape!.FillColor));
        Assert.NotNull(result.Shape.FillGradient);
    }

    [Fact]
    public void Extract_QuickStyleGradient_CachedGradientWins()
    {
        var context = CreateContext();

        var result = ExtractWithContext(GradientShape(QuickStyleGradientNode), context, QuickStyleGradientNode);

        Assert.True(string.IsNullOrEmpty(result.Shape!.FillColor));
        Assert.NotNull(result.Shape.FillGradient);
    }

    [Fact]
    public void Extract_NoContextAvailable_CachedGradientWins()
    {
        // A shape not loaded from a package (synthetic test XML) has no
        // drawing-part back-reference → no colorsDef context → status quo.
        var result = ExtractWithContext(GradientShape(FlatNode), context: null, FlatNode);

        Assert.True(string.IsNullOrEmpty(result.Shape!.FillColor));
        Assert.NotNull(result.Shape.FillGradient);
    }

    [Fact]
    public void Extract_ModelIdReadFromShapeWhenParameterOmitted()
    {
        var context = CreateContext();

        var result = ExtractWithContext(GradientShape(FlatNode), context, modelId: null);

        Assert.True(string.IsNullOrEmpty(result.Shape!.FillColor));
        var gradient = Assert.IsType<TypstGradientFill>(result.Shape.FillGradient);
        Assert.Equal(0.0, gradient.Stops[0].Offset);
        Assert.Equal(1.0, gradient.Stops[gradient.Stops.Count - 1].Offset);
    }

    [Fact]
    public void Extract_CachedSolidFill_NeverOverridden()
    {
        var context = CreateContext();
        var solidShape = ParseXml($@"
<dsp:sp xmlns:dsp=""{DspNs}"" xmlns:a=""{ANs}"" modelId=""{FlatNode}"">
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
</dsp:sp>");

        var result = ExtractWithContext(solidShape, context, FlatNode);

        Assert.Equal("#C00000", result.Shape!.FillColor);
        Assert.Null(result.Shape.FillGradient);
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
