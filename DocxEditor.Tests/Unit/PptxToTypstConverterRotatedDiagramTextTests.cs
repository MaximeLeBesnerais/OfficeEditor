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
/// Rotated diagram text: a SmartArt <c>dsp:sp</c> whose text is rotated either by the
/// shape's own <c>a:xfrm@rot</c> (text rides the shape) or by a <c>dsp:txXfrm@rot</c>
/// (text-only rotation; the txXfrm box is already in post-rotation drawing space) must
/// emit the text rotated about the text-box centre in the Typst source.
/// </summary>
public class PptxToTypstConverterRotatedDiagramTextTests : IDisposable
{
    private readonly string _tempDir;

    public PptxToTypstConverterRotatedDiagramTextTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "PptxRotatedDiagramTextTests", Guid.NewGuid().ToString("N"));
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

    [Fact]
    public void Convert_DiagramShapeRotatedXfrm_TextCarriesShapeRotation()
    {
        // Shape rot=-5400000 (-90°) with txBody and no txXfrm: the text box is the
        // shape rect and rotates with the shape.
        var path = CreateDiagramPptx("shape-rot.pptx", @"<dsp:sp modelId=""{11111111-1111-1111-1111-111111111111}"">
      <dsp:spPr><a:xfrm rot=""-5400000""><a:off x=""0"" y=""0""/><a:ext cx=""1270000"" cy=""508000""/></a:xfrm><a:prstGeom prst=""rect""><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val=""DDDDDD""/></a:solidFill></dsp:spPr>
      <dsp:txBody>
        <a:bodyPr lIns=""12700"" tIns=""12700"" rIns=""12700"" bIns=""12700""><a:noAutofit/></a:bodyPr>
        <a:lstStyle/>
        <a:p><a:r><a:rPr lang=""en-US"" sz=""1500""/><a:t>Label</a:t></a:r></a:p>
      </dsp:txBody>
    </dsp:sp>");

        var (presentation, source) = Convert(path);

        var text = Assert.Single(presentation.Slides[0].Elements, e => e.Type == "Text");
        Assert.Equal(-90.0, text.Rotation, 3);

        // Rotated about the text-box centre (untrimmed box block inside #rotate).
        Assert.Contains("#rotate(-90.0deg, origin: center)[#block(", source);
    }

    [Fact]
    public void Convert_DiagramTxXfrmRot_TextCarriesTxXfrmRotation()
    {
        // Unrotated shape + txXfrm rot=16200000 (270°): text-only rotation about the
        // txXfrm box centre (slide-31 pattern).
        var path = CreateDiagramPptx("txxfrm-rot.pptx", @"<dsp:sp modelId=""{11111111-1111-1111-1111-111111111111}"">
      <dsp:spPr><a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""508000"" cy=""1270000""/></a:xfrm><a:prstGeom prst=""rect""><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val=""DDDDDD""/></a:solidFill></dsp:spPr>
      <dsp:txBody>
        <a:bodyPr lIns=""12700"" tIns=""12700"" rIns=""12700"" bIns=""12700""><a:noAutofit/></a:bodyPr>
        <a:lstStyle/>
        <a:p><a:r><a:rPr lang=""en-US"" sz=""1500""/><a:t>Label</a:t></a:r></a:p>
      </dsp:txBody>
      <dsp:txXfrm rot=""16200000""><a:off x=""0"" y=""0""/><a:ext cx=""1270000"" cy=""508000""/></dsp:txXfrm>
    </dsp:sp>");

        var (presentation, source) = Convert(path);

        var text = Assert.Single(presentation.Slides[0].Elements, e => e.Type == "Text");
        Assert.Equal(270.0, text.Rotation, 3);
        Assert.Contains("#rotate(270.0deg, origin: center)[#block(", source);
    }

    [Fact]
    public void Convert_DiagramShapeRotAndTxXfrmRot_TextUsesTxXfrmRotationOnly()
    {
        // Slide-30 pattern: shape rot=16200000 (270°) AND txXfrm rot=5400000 (90°).
        // The txXfrm box is already in post-rotation drawing space, so the shape
        // rotation must not be re-applied to the text — text rotation is 90° only.
        var path = CreateDiagramPptx("both-rot.pptx", @"<dsp:sp modelId=""{11111111-1111-1111-1111-111111111111}"">
      <dsp:spPr><a:xfrm rot=""16200000""><a:off x=""0"" y=""0""/><a:ext cx=""1270000"" cy=""508000""/></a:xfrm><a:prstGeom prst=""rect""><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val=""DDDDDD""/></a:solidFill></dsp:spPr>
      <dsp:txBody>
        <a:bodyPr lIns=""12700"" tIns=""12700"" rIns=""12700"" bIns=""12700""><a:noAutofit/></a:bodyPr>
        <a:lstStyle/>
        <a:p><a:r><a:rPr lang=""en-US"" sz=""1500""/><a:t>Label</a:t></a:r></a:p>
      </dsp:txBody>
      <dsp:txXfrm rot=""5400000""><a:off x=""0"" y=""0""/><a:ext cx=""508000"" cy=""1270000""/></dsp:txXfrm>
    </dsp:sp>");

        var (presentation, source) = Convert(path);

        var text = Assert.Single(presentation.Slides[0].Elements, e => e.Type == "Text");
        Assert.Equal(90.0, text.Rotation, 3);
        Assert.Contains("#rotate(90.0deg, origin: center)[#block(", source);
    }

    [Fact]
    public void Convert_DiagramShapeRotWithUnrotatedTxXfrm_TextRidesShapeRotation()
    {
        // Slide-152 pattern: shape rot=16200000 (270°) with a txXfrm that has NO rot
        // (cached in pre-rotation space) — the text rides the shape rotation.
        var path = CreateDiagramPptx("shape-rot-plain-txxfrm.pptx", @"<dsp:sp modelId=""{11111111-1111-1111-1111-111111111111}"">
      <dsp:spPr><a:xfrm rot=""16200000""><a:off x=""0"" y=""0""/><a:ext cx=""1270000"" cy=""508000""/></a:xfrm><a:prstGeom prst=""rect""><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val=""DDDDDD""/></a:solidFill></dsp:spPr>
      <dsp:txBody>
        <a:bodyPr lIns=""12700"" tIns=""12700"" rIns=""12700"" bIns=""12700""><a:noAutofit/></a:bodyPr>
        <a:lstStyle/>
        <a:p><a:r><a:rPr lang=""en-US"" sz=""1500""/><a:t>Label</a:t></a:r></a:p>
      </dsp:txBody>
      <dsp:txXfrm><a:off x=""0"" y=""0""/><a:ext cx=""1270000"" cy=""508000""/></dsp:txXfrm>
    </dsp:sp>");

        var (presentation, source) = Convert(path);

        var text = Assert.Single(presentation.Slides[0].Elements, e => e.Type == "Text");
        Assert.Equal(270.0, text.Rotation, 3);
        Assert.Contains("#rotate(270.0deg, origin: center)[#block(", source);
    }

    [Fact]
    public void Convert_DiagramUnrotatedText_NoRotationEmitted()
    {
        var path = CreateDiagramPptx("no-rot.pptx", @"<dsp:sp modelId=""{11111111-1111-1111-1111-111111111111}"">
      <dsp:spPr><a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""1270000"" cy=""508000""/></a:xfrm><a:prstGeom prst=""rect""><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val=""DDDDDD""/></a:solidFill></dsp:spPr>
      <dsp:txBody>
        <a:bodyPr lIns=""12700"" tIns=""12700"" rIns=""12700"" bIns=""12700""><a:noAutofit/></a:bodyPr>
        <a:lstStyle/>
        <a:p><a:r><a:rPr lang=""en-US"" sz=""1500""/><a:t>Label</a:t></a:r></a:p>
      </dsp:txBody>
      <dsp:txXfrm><a:off x=""0"" y=""0""/><a:ext cx=""1270000"" cy=""508000""/></dsp:txXfrm>
    </dsp:sp>");

        var (presentation, source) = Convert(path);

        var text = Assert.Single(presentation.Slides[0].Elements, e => e.Type == "Text");
        Assert.Equal(0.0, text.Rotation, 3);
        Assert.DoesNotContain("#rotate(", source);
    }

    private (TypstPresentation Presentation, string Source) Convert(string path)
    {
        var document = PresentationDocument.Open(path, false);
        var converter = new PptxToTypstConverter(document);
        var presentation = converter.Convert();
        var source = converter.GenerateTypstSource(presentation);
        converter.Dispose();
        document.Dispose();
        return (presentation, source);
    }

    private string CreateDiagramPptx(string fileName, string diagramShapeXml)
    {
        var path = Path.Combine(_tempDir, fileName);

        using (var document = PresentationDocument.Create(path, PresentationDocumentType.Presentation))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new Presentation
            {
                SlideMasterIdList = new SlideMasterIdList(),
                SlideIdList = new SlideIdList(),
                SlideSize = new SlideSize { Cx = (int)Pt(960), Cy = (int)Pt(540), Type = SlideSizeValues.Screen4x3 }
            };

            var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
            slideMasterPart.SlideMaster = new SlideMaster(
                new CommonSlideData(CreateShapeTree()),
                new ColorMap(),
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

            var dataPart = slidePart.AddNewPart<DiagramDataPart>();
            var dataRelId = slidePart.GetIdOfPart(dataPart);
            using (var stream = dataPart.GetStream(FileMode.Create))
            using (var writer = new StreamWriter(stream))
            {
                writer.Write(@"<dgm:dataModel xmlns:dgm=""http://schemas.openxmlformats.org/drawingml/2006/diagram"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><dgm:ptLst/></dgm:dataModel>");
            }

            var drawingPart = slidePart.AddNewPart<DiagramPersistLayoutPart>();
            using (var stream = drawingPart.GetStream(FileMode.Create))
            using (var writer = new StreamWriter(stream))
            {
                writer.Write(@"<dsp:drawing xmlns:dsp=""http://schemas.microsoft.com/office/drawing/2008/diagram"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">
  <dsp:spTree>
    " + diagramShapeXml + @"
  </dsp:spTree>
</dsp:drawing>");
            }

            var relIds = new OpenXmlUnknownElement("dgm", "relIds", "http://schemas.openxmlformats.org/drawingml/2006/diagram");
            relIds.SetAttribute(new OpenXmlAttribute("r", "dm", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", dataRelId));
            var frame = new P.GraphicFrame(
                new NonVisualGraphicFrameProperties(
                    new NonVisualDrawingProperties { Id = 4, Name = "Diagram" },
                    new NonVisualGraphicFrameDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new Transform(
                    new Drawing.Offset { X = Pt(0), Y = Pt(0) },
                    new Drawing.Extents { Cx = Pt(400), Cy = Pt(100) }),
                new Drawing.Graphic(new Drawing.GraphicData(relIds)
                {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
                }));

            slidePart.Slide = new Slide(new CommonSlideData(CreateShapeTree(frame)));
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

    private const long EmusPerPoint = 12700;
    private static long Pt(double points) => (long)Math.Round(points * EmusPerPoint);
}
