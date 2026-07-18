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
/// Tests for the OOXML bodyPr inset defaults applied by <see cref="PptxToTypstConverter"/>:
/// lIns/rIns default to 91440 EMU (7.2 pt) and tIns/bIns to 45720 EMU (3.6 pt) when the
/// attributes are absent (including when bodyPr itself is missing). Explicitly specified
/// values — including explicit zeros — must be preserved.
/// </summary>
public sealed class PptxToTypstConverterInsetsTests : IDisposable
{
    private const double DefaultHorizontalInsetPt = 7.2; // 91440 EMU / 12700
    private const double DefaultVerticalInsetPt = 3.6;   // 45720 EMU / 12700

    private readonly string _tempDir;

    public PptxToTypstConverterInsetsTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), nameof(PptxToTypstConverterInsetsTests), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        SlideOpsTestHelpers.BestEffortDelete(_tempDir);
    }

    [Fact]
    public void Convert_BodyPrWithoutInsetAttributes_UsesOoxmlDefaultInsets()
    {
        var path = CreateDeck(TextShape(2, new Drawing.BodyProperties(), "Hello"));

        var text = ConvertSingleTextElement(path);

        Assert.Equal(DefaultHorizontalInsetPt, text.PaddingLeft, precision: 3);
        Assert.Equal(DefaultVerticalInsetPt, text.PaddingTop, precision: 3);
        Assert.Equal(DefaultHorizontalInsetPt, text.PaddingRight, precision: 3);
        Assert.Equal(DefaultVerticalInsetPt, text.PaddingBottom, precision: 3);
    }

    [Fact]
    public void Convert_BodyPrWithExplicitZeroInsets_PreservesZero()
    {
        var path = CreateDeck(TextShape(2,
            new Drawing.BodyProperties { LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0 },
            "Hello"));

        var text = ConvertSingleTextElement(path);

        Assert.Equal(0, text.PaddingLeft);
        Assert.Equal(0, text.PaddingTop);
        Assert.Equal(0, text.PaddingRight);
        Assert.Equal(0, text.PaddingBottom);
    }

    [Fact]
    public void Convert_BodyPrWithPartialInsets_DefaultsOnlyMissingAttributes()
    {
        // Only lIns and bIns specified (9144 EMU = 0.72 pt, 25400 EMU = 2 pt);
        // tIns/rIns must fall back to the OOXML defaults.
        var path = CreateDeck(TextShape(2,
            new Drawing.BodyProperties { LeftInset = 9144, BottomInset = 25400 },
            "Hello"));

        var text = ConvertSingleTextElement(path);

        Assert.Equal(0.72, text.PaddingLeft, precision: 3);
        Assert.Equal(DefaultVerticalInsetPt, text.PaddingTop, precision: 3);
        Assert.Equal(DefaultHorizontalInsetPt, text.PaddingRight, precision: 3);
        Assert.Equal(2.0, text.PaddingBottom, precision: 3);
    }

    [Fact]
    public void Convert_TextBodyWithoutBodyPr_UsesOoxmlDefaultInsets()
    {
        var path = CreateDeck(TextShapeWithoutBodyPr(2, "Hello"));

        var text = ConvertSingleTextElement(path);

        Assert.Equal(DefaultHorizontalInsetPt, text.PaddingLeft, precision: 3);
        Assert.Equal(DefaultVerticalInsetPt, text.PaddingTop, precision: 3);
        Assert.Equal(DefaultHorizontalInsetPt, text.PaddingRight, precision: 3);
        Assert.Equal(DefaultVerticalInsetPt, text.PaddingBottom, precision: 3);
    }

    private static TypstTextElement ConvertSingleTextElement(string path)
    {
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();

        var element = Assert.Single(presentation.Slides[0].Elements, e => e.Type == "Text");
        return element.Text!;
    }

    private static P.Shape TextShape(uint id, Drawing.BodyProperties bodyPr, string text)
    {
        return new P.Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = id, Name = $"Text {id}" },
                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(
                new Drawing.Transform2D(
                    new Drawing.Offset { X = 1000000, Y = 500000 },
                    new Drawing.Extents { Cx = 4000000, Cy = 1000000 })),
            new TextBody(
                bodyPr,
                new Drawing.ListStyle(),
                new Drawing.Paragraph(new Drawing.Run(new Drawing.Text { Text = text }))));
    }

    private static P.Shape TextShapeWithoutBodyPr(uint id, string text)
    {
        return new P.Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = id, Name = $"Text {id}" },
                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(
                new Drawing.Transform2D(
                    new Drawing.Offset { X = 1000000, Y = 500000 },
                    new Drawing.Extents { Cx = 4000000, Cy = 1000000 })),
            new TextBody(
                new Drawing.ListStyle(),
                new Drawing.Paragraph(new Drawing.Run(new Drawing.Text { Text = text }))));
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
}
