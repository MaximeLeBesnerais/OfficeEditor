using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PptxEditor.Core.Converters;
using PptxEditor.Core.Models;
using Xunit;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Deck-fidelity batch 1 (branch fix/):
/// 1. a:grpFill — freeforms inside groups must inherit the parent group's fill
///    ( slides 2/3/4/9,  map groups 23/40/115).
/// 2. p:style a:fillRef/a:lnRef — shapes whose fill/stroke come only from the theme
///    format scheme must not be dropped ( slide 5, Groups 10/11).
/// 3. cap="small" — small-caps text must be emitted ( headers, master titleStyle).
/// </summary>
public class PptxToTypstConverterTests : IDisposable
{
    private const long EmusPerPoint = 12700;

    private readonly string _tempDir;

    public PptxToTypstConverterTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "PptxTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, true);
            }
        }
        catch
        {
            // Best effort cleanup
        }
    }

    #region Fix 1 — grpFill inheritance

    [Fact]
    public void Convert_GrpFillFreeformInsideGroup_InheritsParentGroupSolidFill()
    {
        var group = GroupShapeWithFill(
            10,
            hexFill: "68BC6C",
            transformGroup: TransformGroup(x: 100, y: 100, width: 200, height: 200, childX: 0, childY: 0, childWidth: 200, childHeight: 200),
            GrpFillFreeform(11, x: 0, y: 0, width: 100, height: 100));
        var path = CreateDeck("grpfill.pptx", group);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();

        var shape = Assert.Single(presentation.Slides[0].Elements, e => e.Type == "Shape");
        Assert.Equal("#68BC6C", shape.Shape?.FillColor);
    }

    [Fact]
    public void Convert_GrpFillFreeformInsideNestedGroup_InheritsOuterGroupFill()
    {
        var innerGroup = GroupShape(
            11,
            transformGroup: null,
            GrpFillFreeform(12, x: 0, y: 0, width: 100, height: 100));
        var outerGroup = GroupShapeWithFill(
            10,
            hexFill: "336699",
            transformGroup: null,
            innerGroup);
        var path = CreateDeck("grpfill-nested.pptx", outerGroup);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();

        var shape = Assert.Single(presentation.Slides[0].Elements, e => e.Type == "Shape");
        Assert.Equal("#336699", shape.Shape?.FillColor);
    }

    [Fact]
    public void Convert_GrpFillFreeform_GroupHasNoFill_StillNotInvented()
    {
        // A grpFill child whose group carries no fill at all must stay unfilled
        // (the gate may drop it — but no color may be fabricated).
        var group = GroupShape(
            10,
            transformGroup: null,
            GrpFillFreeform(11, x: 0, y: 0, width: 100, height: 100));
        var path = CreateDeck("grpfill-nofill.pptx", group);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();

        Assert.DoesNotContain(presentation.Slides[0].Elements,
            e => e.Type == "Shape" && !string.IsNullOrEmpty(e.Shape?.FillColor));
    }

    #endregion

    #region Fix 2 — p:style fillRef/lnRef resolution

    [Fact]
    public void Convert_StyleFillReference_ResolvesThemeSolidFillAndLine()
    {
        var group = GroupShape(10, transformGroup: null,
            StyledRoundRect(11, fillRefIdx: 1, lnRefIdx: 1, schemeColorName: "accent1"));
        var path = CreateDeck("styleref-solid.pptx", withTheme: true, group);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();

        var shape = Assert.Single(presentation.Slides[0].Elements, e => e.Type == "Shape");
        Assert.Equal("#FF0000", shape.Shape?.FillColor);
        Assert.Equal("#FF0000", shape.Shape?.StrokeColor);
        Assert.Equal(0.75, shape.Shape?.StrokeWidth ?? 0, 2);
    }

    [Fact]
    public void Convert_StyleFillReferenceGradient_ResolvesThemeGradientStops()
    {
        var group = GroupShape(10, transformGroup: null,
            StyledRoundRect(11, fillRefIdx: 3, lnRefIdx: 1, schemeColorName: "accent1"));
        var path = CreateDeck("styleref-gradient.pptx", withTheme: true, group);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();

        var shape = Assert.Single(presentation.Slides[0].Elements, e => e.Type == "Shape");
        Assert.NotNull(shape.Shape?.FillGradient);
        Assert.Equal(2, shape.Shape!.FillGradient!.Stops.Count);
    }

    [Fact]
    public void Convert_StyleReference_DoesNotOverrideExplicitFill()
    {
        var shapeWithFill = StyledRoundRect(11, fillRefIdx: 1, lnRefIdx: 1, schemeColorName: "accent1");
        shapeWithFill.ShapeProperties!.Append(new Drawing.SolidFill(
            new Drawing.RgbColorModelHex { Val = "0000FF" }));
        var path = CreateDeck("styleref-explicit.pptx", withTheme: true, shapeWithFill);

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();

        var shape = Assert.Single(presentation.Slides[0].Elements, e => e.Type == "Shape");
        Assert.Equal("#0000FF", shape.Shape?.FillColor);
    }

    #endregion

    #region Test helpers

    private string CreateDeck(string fileName, params OpenXmlElement[] slideElements)
        => CreateDeck(fileName, withTheme: false, slideElements);

    private string CreateDeck(string fileName, bool withTheme, params OpenXmlElement[] slideElements)
    {
        var path = Path.Combine(_tempDir, fileName);

        using (var document = PresentationDocument.Create(path, PresentationDocumentType.Presentation))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new Presentation
            {
                SlideMasterIdList = new SlideMasterIdList(),
                SlideIdList = new SlideIdList(),
                SlideSize = new SlideSize { Cx = (int)Pt(720), Cy = (int)Pt(540), Type = SlideSizeValues.Screen4x3 }
            };

            var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
            slideMasterPart.SlideMaster = new SlideMaster(
                new CommonSlideData(CreateShapeTree()),
                new ColorMap
                {
                    Background1 = Drawing.ColorSchemeIndexValues.Light1,
                    Text1 = Drawing.ColorSchemeIndexValues.Dark1,
                    Background2 = Drawing.ColorSchemeIndexValues.Light2,
                    Text2 = Drawing.ColorSchemeIndexValues.Dark2,
                    Accent1 = Drawing.ColorSchemeIndexValues.Accent1,
                    Accent2 = Drawing.ColorSchemeIndexValues.Accent2,
                    Accent3 = Drawing.ColorSchemeIndexValues.Accent3,
                    Accent4 = Drawing.ColorSchemeIndexValues.Accent4,
                    Accent5 = Drawing.ColorSchemeIndexValues.Accent5,
                    Accent6 = Drawing.ColorSchemeIndexValues.Accent6,
                    Hyperlink = Drawing.ColorSchemeIndexValues.Hyperlink,
                    FollowedHyperlink = Drawing.ColorSchemeIndexValues.FollowedHyperlink
                },
                new SlideLayoutIdList());

            if (withTheme)
            {
                AddTestTheme(slideMasterPart);
            }

            var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
            slideLayoutPart.SlideLayout = new P.SlideLayout(new CommonSlideData(CreateShapeTree()));
            slideLayoutPart.AddPart(slideMasterPart);
            slideMasterPart.SlideMaster.SlideLayoutIdList!.Append(new SlideLayoutId
            {
                Id = 2147483649,
                RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart)
            });

            presentationPart.Presentation.SlideMasterIdList.Append(new SlideMasterId
            {
                Id = 2147483648,
                RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
            });

            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.Slide = new Slide(new CommonSlideData(CreateShapeTree(slideElements)));
            slidePart.AddPart(slideLayoutPart);

            presentationPart.Presentation.SlideIdList.Append(new SlideId
            {
                Id = 256,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });
        }

        return path;
    }

    private static ShapeTree CreateShapeTree(params OpenXmlElement[] elements)
    {
        var shapeTree = new ShapeTree(
            new NonVisualGroupShapeProperties(
                new NonVisualDrawingProperties { Id = 0, Name = "" },
                new NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(
                new Drawing.TransformGroup(
                    new Drawing.Offset { X = 0, Y = 0 },
                    new Drawing.Extents { Cx = 0, Cy = 0 },
                    new Drawing.ChildOffset { X = 0, Y = 0 },
                    new Drawing.ChildExtents { Cx = 0, Cy = 0 })));

        foreach (var element in elements)
        {
            shapeTree.Append(element);
        }

        return shapeTree;
    }

    private static P.GroupShape GroupShape(uint id, Drawing.TransformGroup? transformGroup, params OpenXmlElement[] children)
    {
        var groupShape = new P.GroupShape(
            new NonVisualGroupShapeProperties(
                new NonVisualDrawingProperties { Id = id, Name = $"Group {id}" },
                new NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            transformGroup == null ? new GroupShapeProperties() : new GroupShapeProperties(transformGroup));

        foreach (var child in children)
        {
            groupShape.Append(child);
        }

        return groupShape;
    }

    private static P.GroupShape GroupShapeWithFill(uint id, string hexFill, Drawing.TransformGroup? transformGroup, params OpenXmlElement[] children)
    {
        var groupShapeProperties = new GroupShapeProperties();
        if (transformGroup != null)
        {
            groupShapeProperties.Append(transformGroup);
        }
        groupShapeProperties.Append(new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = hexFill }));

        var groupShape = new P.GroupShape(
            new NonVisualGroupShapeProperties(
                new NonVisualDrawingProperties { Id = id, Name = $"Group {id}" },
                new NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            groupShapeProperties);

        foreach (var child in children)
        {
            groupShape.Append(child);
        }

        return groupShape;
    }

    /// <summary>
    /// custGeom freeform with a:grpFill (inherit the group fill) and a:noFill outline —
    /// the / group-child pattern.
    /// </summary>
    private static P.Shape GrpFillFreeform(uint id, double x, double y, double width, double height)
    {
        return new P.Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = id, Name = $"Freeform {id}" },
                new NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(
                new Drawing.Transform2D(
                    new Drawing.Offset { X = Pt(x), Y = Pt(y) },
                    new Drawing.Extents { Cx = Pt(width), Cy = Pt(height) }),
                new Drawing.CustomGeometry(
                    new Drawing.AdjustValueList(),
                    new Drawing.Rectangle(),
                    new Drawing.PathList(
                        new Drawing.Path(
                            new Drawing.MoveTo(new Drawing.Point { X = "0", Y = "0" }),
                            new Drawing.LineTo(new Drawing.Point { X = "914400", Y = "0" }),
                            new Drawing.LineTo(new Drawing.Point { X = "914400", Y = "914400" }),
                            new Drawing.CloseShapePath()) { Width = 914400, Height = 914400 })),
                new Drawing.GroupFill(),
                new Drawing.Outline(new Drawing.NoFill())));
    }

    private static Drawing.TransformGroup TransformGroup(double x, double y, double width, double height, double childX, double childY, double childWidth, double childHeight)
    {
        return new Drawing.TransformGroup(
            new Drawing.Offset { X = Pt(x), Y = Pt(y) },
            new Drawing.Extents { Cx = Pt(width), Cy = Pt(height) },
            new Drawing.ChildOffset { X = Pt(childX), Y = Pt(childY) },
            new Drawing.ChildExtents { Cx = Pt(childWidth), Cy = Pt(childHeight) });
    }

    /// <summary>
    /// Theme with an explicit format scheme: fillStyleLst = [solid phClr, solid phClr
    /// tint 50%, gradient phClr shade 51%→94%], lnStyleLst = three solid-phClr outlines.
    /// accent1 is pure red so fillRef/lnRef resolution is easy to assert.
    /// </summary>
    private static void AddTestTheme(SlideMasterPart slideMasterPart)
    {
        var themePart = slideMasterPart.AddNewPart<ThemePart>();
        themePart.Theme = new Drawing.Theme(
            new Drawing.ThemeElements(
                new Drawing.ColorScheme(
                    new Drawing.Dark1Color(new Drawing.RgbColorModelHex { Val = "000000" }),
                    new Drawing.Light1Color(new Drawing.RgbColorModelHex { Val = "FFFFFF" }),
                    new Drawing.Dark2Color(new Drawing.RgbColorModelHex { Val = "44546A" }),
                    new Drawing.Light2Color(new Drawing.RgbColorModelHex { Val = "E7E6E6" }),
                    new Drawing.Accent1Color(new Drawing.RgbColorModelHex { Val = "FF0000" }),
                    new Drawing.Accent2Color(new Drawing.RgbColorModelHex { Val = "00FF00" }),
                    new Drawing.Accent3Color(new Drawing.RgbColorModelHex { Val = "0000FF" }),
                    new Drawing.Accent4Color(new Drawing.RgbColorModelHex { Val = "FFFF00" }),
                    new Drawing.Accent5Color(new Drawing.RgbColorModelHex { Val = "FF00FF" }),
                    new Drawing.Accent6Color(new Drawing.RgbColorModelHex { Val = "00FFFF" }),
                    new Drawing.Hyperlink(new Drawing.RgbColorModelHex { Val = "0563C1" }),
                    new Drawing.FollowedHyperlinkColor(new Drawing.RgbColorModelHex { Val = "954F72" }))
                { Name = "Test" },
                new Drawing.FontScheme(
                    new Drawing.MajorFont(new Drawing.LatinFont { Typeface = "Arial" }),
                    new Drawing.MinorFont(new Drawing.LatinFont { Typeface = "Arial" }))
                { Name = "Test" },
                new Drawing.FormatScheme(
                    new Drawing.FillStyleList(
                        new Drawing.SolidFill(PhClr()),
                        new Drawing.SolidFill(PhClr(new Drawing.Tint { Val = 50000 })),
                        new Drawing.GradientFill(
                            new Drawing.GradientStopList(
                                new Drawing.GradientStop(PhClr(new Drawing.Shade { Val = 51000 })) { Position = 0 },
                                new Drawing.GradientStop(PhClr(new Drawing.Shade { Val = 94000 })) { Position = 100000 }),
                            new Drawing.LinearGradientFill { Angle = 16200000, Scaled = true })),
                    new Drawing.LineStyleList(
                        new Drawing.Outline(new Drawing.SolidFill(PhClr())) { Width = 9525 },
                        new Drawing.Outline(new Drawing.SolidFill(PhClr())) { Width = 25400 },
                        new Drawing.Outline(new Drawing.SolidFill(PhClr())) { Width = 38100 }),
                    new Drawing.EffectStyleList(
                        new Drawing.EffectStyle(),
                        new Drawing.EffectStyle(),
                        new Drawing.EffectStyle()),
                    new Drawing.BackgroundFillStyleList(
                        new Drawing.SolidFill(PhClr()),
                        new Drawing.SolidFill(PhClr()),
                        new Drawing.SolidFill(PhClr())))
                { Name = "Test" }))
        { Name = "Test" };
    }

    private static Drawing.SchemeColor PhClr(params OpenXmlElement[] transforms)
    {
        var schemeColor = new Drawing.SchemeColor { Val = Drawing.SchemeColorValues.PhColor };
        foreach (var transform in transforms)
        {
            schemeColor.Append(transform);
        }
        return schemeColor;
    }

    private static Drawing.SchemeColor SchemeColorRef(string name)
    {
        var schemeColor = new Drawing.SchemeColor();
        schemeColor.SetAttribute(new OpenXmlAttribute("val", string.Empty, name));
        return schemeColor;
    }

    /// <summary>
    /// roundRect whose fill/line come ONLY from p:style (fillRef/lnRef into the theme
    /// format scheme) — the  slide 5 Group 10/11 child pattern: no fill marker
    /// and no a:ln in spPr.
    /// </summary>
    private static P.Shape StyledRoundRect(uint id, int fillRefIdx, int lnRefIdx, string schemeColorName)
    {
        var fontReference = new Drawing.FontReference(SchemeColorRef(schemeColorName));
        fontReference.SetAttribute(new OpenXmlAttribute("idx", string.Empty, "minor"));

        return new P.Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = id, Name = $"RoundRect {id}" },
                new NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(
                new Drawing.Transform2D(
                    new Drawing.Offset { X = Pt(10), Y = Pt(10) },
                    new Drawing.Extents { Cx = Pt(100), Cy = Pt(100) }),
                new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = Drawing.ShapeTypeValues.RoundRectangle }),
            new P.ShapeStyle(
                new Drawing.LineReference(SchemeColorRef(schemeColorName)) { Index = (uint)lnRefIdx },
                new Drawing.FillReference(SchemeColorRef(schemeColorName)) { Index = (uint)fillRefIdx },
                new Drawing.EffectReference(SchemeColorRef(schemeColorName)) { Index = 0U },
                fontReference));
    }

    private static long Pt(double points) => (long)Math.Round(points * EmusPerPoint);

    #endregion
}
