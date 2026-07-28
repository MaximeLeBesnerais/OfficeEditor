using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PptxEditor.Core.Converters;
using Xunit;

namespace DocxEditor.Tests.Unit;

/// <summary>Small synthetic regressions for the remaining  converter edge cases.</summary>
public sealed class PptxToTypstConverterNightTests : IDisposable
{
    private const string PngBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";
    private readonly string _tempDir = SlideOpsTestHelpers.CreateTempDirectory(nameof(PptxToTypstConverterNightTests));

    public void Dispose() => SlideOpsTestHelpers.BestEffortDelete(_tempDir);

    [Fact]
    public void Convert_SlidePictureOverridesLayoutPictureAtSameFrame()
    {
        var path = Path.Combine(_tempDir, "picture-override.pptx");
        var layoutBytes = Convert.FromBase64String(PngBase64);
        var slideBytes = Convert.FromBase64String(PngBase64);

        using (var document = PresentationDocument.Create(path, PresentationDocumentType.Presentation))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new Presentation
            {
                SlideMasterIdList = new SlideMasterIdList(),
                SlideIdList = new SlideIdList(),
                SlideSize = new SlideSize { Cx = 12192000, Cy = 6858000 }
            };

            var masterPart = presentationPart.AddNewPart<SlideMasterPart>();
            masterPart.SlideMaster = new SlideMaster(
                new CommonSlideData(CreateShapeTree()),
                new ColorMap(),
                new SlideLayoutIdList());
            var layoutPart = masterPart.AddNewPart<SlideLayoutPart>();
            layoutPart.AddPart(masterPart);
            layoutPart.SlideLayout = new P.SlideLayout(new CommonSlideData(CreateShapeTree(LayoutPicture("rId7"))));
            masterPart.SlideMaster.SlideLayoutIdList!.Append(new SlideLayoutId
            {
                Id = 1,
                RelationshipId = masterPart.GetIdOfPart(layoutPart)
            });
            presentationPart.Presentation.SlideMasterIdList.Append(new SlideMasterId
            {
                Id = 1,
                RelationshipId = presentationPart.GetIdOfPart(masterPart)
            });

            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.AddPart(layoutPart);
            presentationPart.Presentation.SlideIdList.Append(new SlideId
            {
                Id = 1,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });

            var layoutImage = layoutPart.AddImagePart(ImagePartType.Png, "rId7");
            using (var stream = new MemoryStream(layoutBytes)) layoutImage.FeedData(stream);
            var slideImage = slidePart.AddImagePart(ImagePartType.Png, "rId7");
            using (var stream = new MemoryStream(slideBytes)) slideImage.FeedData(stream);
            slidePart.Slide = new Slide(new CommonSlideData(CreateShapeTree(SlidePicture("rId7"))));
        }

        using var opened = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(opened);
        var images = converter.Convert().Slides[0].Elements.Where(e => e.Type == "Image").ToList();
        Assert.Single(images);
    }

    [Fact]
    public void Convert_ZeroDimensionConnectorEmitsTypstLine()
    {
        var path = Path.Combine(_tempDir, "connector.pptx");
        using (var document = PresentationDocument.Create(path, PresentationDocumentType.Presentation))
        {
            var presentationPart = CreatePresentation(document);
            var masterPart = presentationPart.GetPartsOfType<SlideMasterPart>().Single();
            var layoutPart = masterPart.AddNewPart<SlideLayoutPart>();
            layoutPart.SlideLayout = new P.SlideLayout(new CommonSlideData(CreateShapeTree()));
            layoutPart.AddPart(masterPart);
            masterPart.SlideMaster!.SlideLayoutIdList!.Append(new SlideLayoutId
            {
                Id = 1,
                RelationshipId = masterPart.GetIdOfPart(layoutPart)
            });
            presentationPart.Presentation!.SlideMasterIdList!.Append(new SlideMasterId
            {
                Id = 1,
                RelationshipId = presentationPart.GetIdOfPart(masterPart)
            });
            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.AddPart(layoutPart);
            slidePart.Slide = new Slide(new CommonSlideData(CreateShapeTree(VerticalConnector())));
            presentationPart.Presentation.SlideIdList!.Append(new SlideId
            {
                Id = 1,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });
        }

        using var opened = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(opened);
        var presentation = converter.Convert();
        var shape = Assert.Single(presentation.Slides[0].Elements, e => e.Type == "Shape");
        Assert.Equal("line", shape.Shape!.ShapeType);
        Assert.Contains("#line(length:", converter.GenerateTypstSource(presentation));
    }

    [Fact]
    public void Convert_ParagraphHyperlinkAppliesLinkFormattingToRuns()
    {
        var path = Path.Combine(_tempDir, "paragraph-link.pptx");
        using (var document = PresentationDocument.Create(path, PresentationDocumentType.Presentation))
        {
            var presentationPart = CreatePresentation(document);
            var masterPart = presentationPart.GetPartsOfType<SlideMasterPart>().Single();
            var layoutPart = masterPart.AddNewPart<SlideLayoutPart>();
            layoutPart.SlideLayout = new P.SlideLayout(new CommonSlideData(CreateShapeTree()));
            layoutPart.AddPart(masterPart);
            masterPart.SlideMaster!.SlideLayoutIdList!.Append(new SlideLayoutId { Id = 1, RelationshipId = masterPart.GetIdOfPart(layoutPart) });
            presentationPart.Presentation!.SlideMasterIdList!.Append(new SlideMasterId { Id = 1, RelationshipId = presentationPart.GetIdOfPart(masterPart) });
            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.AddPart(layoutPart);
            slidePart.Slide = new Slide(new CommonSlideData(CreateShapeTree(HyperlinkShape())));
            presentationPart.Presentation.SlideIdList!.Append(new SlideId { Id = 1, RelationshipId = presentationPart.GetIdOfPart(slidePart) });
        }

        using var opened = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(opened);
        var text = Assert.Single(converter.Convert().Slides[0].Elements, e => e.Type == "Text").Text!;
        var run = Assert.Single(text.Paragraphs.SelectMany(p => p.Runs));
        Assert.True(run.Formatting.Underline);
    }

    private static PresentationPart CreatePresentation(PresentationDocument document)
    {
        var part = document.AddPresentationPart();
        part.Presentation = new Presentation
        {
            SlideMasterIdList = new SlideMasterIdList(),
            SlideIdList = new SlideIdList(),
            SlideSize = new SlideSize { Cx = 12192000, Cy = 6858000 }
        };
        var master = part.AddNewPart<SlideMasterPart>();
        master.SlideMaster = new SlideMaster(new CommonSlideData(CreateShapeTree()), new ColorMap(), new SlideLayoutIdList());
        return part;
    }

    private static P.Picture LayoutPicture(string relationshipId) => Picture(relationshipId, 20);
    private static P.Picture SlidePicture(string relationshipId) => Picture(relationshipId, 21);

    private static P.Picture Picture(string relationshipId, uint id) => new(
        new P.NonVisualPictureProperties(
            new NonVisualDrawingProperties { Id = id, Name = "Picture" },
            new P.NonVisualPictureDrawingProperties(),
            new ApplicationNonVisualDrawingProperties()),
        new P.BlipFill(new Drawing.Blip { Embed = relationshipId }, new Drawing.Stretch(new Drawing.FillRectangle())),
        new ShapeProperties(new Drawing.Transform2D(
            new Drawing.Offset { X = 1000000, Y = 1000000 },
            new Drawing.Extents { Cx = 3000000, Cy = 2000000 })));

    private static P.ConnectionShape VerticalConnector() => new(
        new P.NonVisualConnectionShapeProperties(
            new NonVisualDrawingProperties { Id = 30, Name = "Vertical connector" },
            new P.NonVisualConnectorShapeDrawingProperties(),
            new ApplicationNonVisualDrawingProperties()),
        new ShapeProperties(
            new Drawing.Transform2D(new Drawing.Offset { X = 1000000, Y = 1000000 }, new Drawing.Extents { Cx = 0, Cy = 2000000 }),
            new Drawing.Outline(new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = "000000" })) { Width = 12700 }));

    private static P.Shape HyperlinkShape() => new(
        new P.NonVisualShapeProperties(
            new NonVisualDrawingProperties { Id = 40, Name = "Paragraph hyperlink" },
            new P.NonVisualShapeDrawingProperties(),
            new ApplicationNonVisualDrawingProperties()),
        new ShapeProperties(new Drawing.Transform2D(new Drawing.Offset { X = 1000000, Y = 1000000 }, new Drawing.Extents { Cx = 4000000, Cy = 500000 })),
        new P.TextBody(new Drawing.BodyProperties(), new Drawing.ListStyle(),
            new Drawing.Paragraph(
                new Drawing.ParagraphProperties(new Drawing.HyperlinkOnClick { Id = "rId99" }),
                new Drawing.Run(new Drawing.RunProperties { FontSize = 1200 }, new Drawing.Text("paragraph link")))));

    private static ShapeTree CreateShapeTree(params OpenXmlElement[] elements)
    {
        var tree = new ShapeTree(
            new P.NonVisualGroupShapeProperties(new NonVisualDrawingProperties { Id = 0, Name = "" }, new P.NonVisualGroupShapeDrawingProperties(), new ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties(new Drawing.TransformGroup(
                new Drawing.Offset { X = 0, Y = 0 }, new Drawing.Extents { Cx = 0, Cy = 0 },
                new Drawing.ChildOffset { X = 0, Y = 0 }, new Drawing.ChildExtents { Cx = 0, Cy = 0 })));
        tree.Append(elements);
        return tree;
    }
}
