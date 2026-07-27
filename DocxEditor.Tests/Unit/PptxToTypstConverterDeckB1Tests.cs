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

    #region Test helpers

    private string CreateDeck(string fileName, params OpenXmlElement[] slideElements)
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

    private static long Pt(double points) => (long)Math.Round(points * EmusPerPoint);

    #endregion
}
