using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PptxEditor.Core.Converters;
using PptxEditor.Core.Models;
using Xunit;

namespace DocxEditor.Tests.Unit;

/// <summary>Regression contracts for the  and  freeform sweep.</summary>
public sealed class PptxToTypstConverterNightTests : IDisposable
{
    private readonly string _tempDir = Path.Combine(
        Path.GetTempPath(), nameof(PptxToTypstConverterNightTests), Guid.NewGuid().ToString("N"));

    public PptxToTypstConverterNightTests() => Directory.CreateDirectory(_tempDir);

    public void Dispose() => SlideOpsTestHelpers.BestEffortDelete(_tempDir);

    [Fact]
    public void CustomBezier_IsFlattenedWithEnoughPointsForIconCurves()
    {
        var shape = FreeformShape(2,
            new Drawing.MoveTo(new Drawing.Point { X = "0", Y = "0" }),
            new Drawing.CubicBezierCurveTo(
                new Drawing.Point { X = "0", Y = "1000000" },
                new Drawing.Point { X = "1000000", Y = "1000000" },
                new Drawing.Point { X = "1000000", Y = "0" }),
            new Drawing.CloseShapePath());

        var element = ConvertSingleShape(CreateDeck(shape));

        Assert.Equal("polygon", element.ShapeType);
        Assert.True(element.Points.Count >= 25, "Cubic freeforms need a smooth 24-step flattening.");
    }

    [Fact]
    public void CustomGeometry_AppliesHorizontalFlipToNormalizedPoints()
    {
        var shape = FreeformShape(2,
            new Drawing.MoveTo(new Drawing.Point { X = "0", Y = "0" }),
            new Drawing.LineTo(new Drawing.Point { X = "1000000", Y = "0" }),
            new Drawing.LineTo(new Drawing.Point { X = "0", Y = "1000000" }),
            new Drawing.CloseShapePath());
        shape.ShapeProperties!.Transform2D!.SetAttribute(new OpenXmlAttribute("flipH", string.Empty, "1"));

        var element = ConvertSingleShape(CreateDeck(shape));

        Assert.Equal((1.0, 0.0), element.Points[0]);
        Assert.Equal((0.0, 0.0), element.Points[1]);
        Assert.Equal((1.0, 1.0), element.Points[2]);
    }

    [Fact]
    public void SolidConnectorLineWithoutWidth_UsesDrawingMlDefaultWidth()
    {
        var shape = new P.Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = 2, Name = "Divider" },
                new NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(
                new Drawing.Transform2D(
                    new Drawing.Offset { X = 100000, Y = 100000 },
                    new Drawing.Extents { Cx = 0, Cy = 1000000 }),
                new Drawing.Outline(new Drawing.SolidFill(
                    new Drawing.RgbColorModelHex { Val = "FFFFFF" }))));

        var element = ConvertSingleShape(CreateDeck(shape));

        Assert.Equal(1.0, element.StrokeWidth, precision: 6);
        Assert.Equal("#FFFFFF", element.StrokeColor);
    }

    private static TypstShapeElement ConvertSingleShape(string path)
    {
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);
        var presentation = converter.Convert();
        return Assert.Single(presentation.Slides[0].Elements, e => e.Type == "Shape").Shape!;
    }

    private static P.Shape FreeformShape(uint id, params OpenXmlElement[] pathCommands)
    {
        return new P.Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = id, Name = "Freeform" },
                new NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(
                new Drawing.Transform2D(
                    new Drawing.Offset { X = 100000, Y = 100000 },
                    new Drawing.Extents { Cx = 1000000, Cy = 1000000 }),
                new Drawing.CustomGeometry(
                    new Drawing.AdjustValueList(),
                    new Drawing.Rectangle(),
                    new Drawing.PathList(new Drawing.Path(pathCommands) { Width = 1000000, Height = 1000000 })),
                new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = "123456" })));
    }

    private string CreateDeck(params OpenXmlElement[] slideElements)
    {
        var path = Path.Combine(_tempDir, $"{Guid.NewGuid():N}.pptx");
        using (var document = PresentationDocument.Create(path, PresentationDocumentType.Presentation))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new Presentation
            {
                SlideMasterIdList = new SlideMasterIdList(),
                SlideIdList = new SlideIdList(),
                SlideSize = new SlideSize { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 }
            };

            var masterPart = presentationPart.AddNewPart<SlideMasterPart>();
            masterPart.SlideMaster = new SlideMaster(
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

            var layoutPart = masterPart.AddNewPart<SlideLayoutPart>();
            layoutPart.SlideLayout = new P.SlideLayout(new CommonSlideData(CreateShapeTree()));
            layoutPart.AddPart(masterPart);
            masterPart.SlideMaster.SlideLayoutIdList!.Append(new SlideLayoutId
            {
                Id = 2147483649,
                RelationshipId = masterPart.GetIdOfPart(layoutPart)
            });

            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.Slide = new Slide(new CommonSlideData(CreateShapeTree(slideElements)));
            slidePart.AddPart(layoutPart);
            presentationPart.Presentation.SlideMasterIdList.Append(new SlideMasterId
            {
                Id = 2147483648,
                RelationshipId = presentationPart.GetIdOfPart(masterPart)
            });
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
        var tree = new ShapeTree(
            new NonVisualGroupShapeProperties(
                new NonVisualDrawingProperties { Id = 0, Name = string.Empty },
                new NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(
                new Drawing.TransformGroup(
                    new Drawing.Offset { X = 0, Y = 0 },
                    new Drawing.Extents { Cx = 0, Cy = 0 },
                    new Drawing.ChildOffset { X = 0, Y = 0 },
                    new Drawing.ChildExtents { Cx = 0, Cy = 0 })));
        foreach (var element in elements)
            tree.Append(element);
        return tree;
    }
}
