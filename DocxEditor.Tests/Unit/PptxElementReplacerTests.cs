using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PptxEditor.Core.Builders;
using PptxEditor.Core.Services;

namespace DocxEditor.Tests.Unit;

public class PptxElementReplacerTests : IDisposable
{
    private const string TestTableStyleId = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";

    private readonly string _testDir = Path.Combine(Path.GetTempPath(), $"pptx_replacer_tests_{Guid.NewGuid():N}");
    private readonly PptxElementReplacer _replacer = new();

    public PptxElementReplacerTests()
    {
        Directory.CreateDirectory(_testDir);
    }

    #region Helpers

    private string NewDeckPath() => Path.Combine(_testDir, $"deck_{Guid.NewGuid():N}.pptx");

    private string NewImagePath(string extension)
    {
        var path = Path.Combine(_testDir, $"image_{Guid.NewGuid():N}{extension}");
        if (extension.Equals(".png", StringComparison.OrdinalIgnoreCase))
        {
            CreateMinimalPng(path);
        }
        else
        {
            CreateMinimalJpeg(path);
        }
        return path;
    }

    private static SlidePart FirstSlidePart(PresentationDocument doc) =>
        doc.PresentationPart!.SlideParts.First();

    private static P.Shape FirstShape(PresentationDocument doc) =>
        FirstSlidePart(doc).Slide!.CommonSlideData!.ShapeTree!.Elements<P.Shape>().First();

    private static P.Picture FirstPicture(SlidePart slidePart) =>
        slidePart.Slide!.CommonSlideData!.ShapeTree!.Elements<P.Picture>().First();

    private static P.GraphicFrame FirstGraphicFrame(SlidePart slidePart) =>
        slidePart.Slide!.CommonSlideData!.ShapeTree!.Elements<P.GraphicFrame>().First();

    private static Drawing.Table FirstTable(P.GraphicFrame frame) =>
        frame.Graphic!.GraphicData!.Elements<Drawing.Table>().First();

    private static uint ReadElementId(OpenXmlElement nvProperties)
    {
        var cNvPr = nvProperties.ChildElements.First(e => e.LocalName == "cNvPr");
        var match = Regex.Match(cNvPr.OuterXml, @"\bid\s*=\s*""([^""]*)""");
        Assert.True(match.Success, "cNvPr id attribute not found");
        return uint.Parse(match.Groups[1].Value);
    }

    private string CreateDeckWithText(string text)
    {
        var path = NewDeckPath();
        using (var builder = PresentationBuilder.Create(path))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddText(text);
            builder.Save();
        }
        return path;
    }

    private string CreateDeckWithImage(out uint pictureId)
    {
        var path = NewDeckPath();
        var imagePath = NewImagePath(".jpg");
        using (var builder = PresentationBuilder.Create(path))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddImage(imagePath);
            builder.Save();
        }
        using (var doc = PresentationDocument.Open(path, false))
        {
            pictureId = ReadElementId(FirstPicture(FirstSlidePart(doc)).NonVisualPictureProperties!);
        }
        return path;
    }

    private string CreateDeckWithTable(int rows, int cols, out uint frameId)
    {
        var path = NewDeckPath();
        var data = Enumerable.Range(0, rows)
            .Select(r => Enumerable.Range(0, cols).Select(c => $"R{r}C{c}").ToList())
            .ToList();
        using (var builder = PresentationBuilder.Create(path))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTable(data);
            builder.Save();
        }
        using (var doc = PresentationDocument.Open(path, false))
        {
            frameId = ReadElementId(FirstGraphicFrame(FirstSlidePart(doc)).NonVisualGraphicFrameProperties!);
        }
        return path;
    }

    private static void CreateMinimalJpeg(string path)
    {
        // Minimal valid JPEG: 1x1 pixel, gray
        var jpegBytes = new byte[]
        {
            0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0x4A, 0x46, 0x49, 0x46, 0x00, 0x01,
            0x01, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00, 0xFF, 0xDB, 0x00, 0x43,
            0x00, 0x08, 0x06, 0x06, 0x07, 0x06, 0x05, 0x08, 0x07, 0x07, 0x07, 0x09,
            0x09, 0x08, 0x0A, 0x0C, 0x14, 0x0D, 0x0C, 0x0B, 0x0B, 0x0C, 0x19, 0x12,
            0x13, 0x0F, 0x14, 0x1D, 0x1A, 0x1F, 0x1E, 0x1D, 0x1A, 0x1C, 0x1C, 0x20,
            0x24, 0x2E, 0x27, 0x20, 0x22, 0x2C, 0x23, 0x1C, 0x1C, 0x28, 0x37, 0x29,
            0x2C, 0x30, 0x31, 0x34, 0x34, 0x34, 0x1F, 0x27, 0x39, 0x3D, 0x38, 0x32,
            0x3C, 0x2E, 0x33, 0x34, 0x32, 0xFF, 0xC0, 0x00, 0x0B, 0x08, 0x00, 0x01,
            0x00, 0x01, 0x01, 0x01, 0x11, 0x00, 0xFF, 0xC4, 0x00, 0x1F, 0x00, 0x00,
            0x01, 0x05, 0x01, 0x01, 0x01, 0x01, 0x01, 0x01, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08,
            0x09, 0x0A, 0x0B, 0xFF, 0xC4, 0x00, 0xB5, 0x10, 0x00, 0x02, 0x01, 0x03,
            0x03, 0x02, 0x04, 0x03, 0x05, 0x05, 0x04, 0x04, 0x00, 0x00, 0x01, 0x7D,
            0x01, 0x02, 0x03, 0x00, 0x04, 0x11, 0x05, 0x12, 0x21, 0x31, 0x41, 0x06,
            0x13, 0x51, 0x61, 0x07, 0x22, 0x71, 0x14, 0x32, 0x81, 0x91, 0xA1, 0x08,
            0x23, 0x42, 0xB1, 0xC1, 0x15, 0x52, 0xD1, 0xF0, 0x24, 0x33, 0x62, 0x72,
            0x82, 0x09, 0x0A, 0x16, 0x17, 0x18, 0x19, 0x1A, 0x25, 0x26, 0x27, 0x28,
            0x29, 0x2A, 0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x3A, 0x43, 0x44, 0x45,
            0x46, 0x47, 0x48, 0x49, 0x4A, 0x53, 0x54, 0x55, 0x56, 0x57, 0x58, 0x59,
            0x5A, 0x63, 0x64, 0x65, 0x66, 0x67, 0x68, 0x69, 0x6A, 0x73, 0x74, 0x75,
            0x76, 0x77, 0x78, 0x79, 0x7A, 0x83, 0x84, 0x85, 0x86, 0x87, 0x88, 0x89,
            0x8A, 0x92, 0x93, 0x94, 0x95, 0x96, 0x97, 0x98, 0x99, 0x9A, 0xA2, 0xA3,
            0xA4, 0xA5, 0xA6, 0xA7, 0xA8, 0xA9, 0xAA, 0xAB, 0xAC, 0xAD, 0xAE, 0xAF,
            0xB0, 0xB1, 0xB2, 0xB3, 0xB4, 0xB5, 0xB6, 0xB7, 0xB8, 0xB9, 0xBA, 0xC2,
            0xC3, 0xC4, 0xC5, 0xC6, 0xC7, 0xC8, 0xC9, 0xCA, 0xD2, 0xD3, 0xD4, 0xD5,
            0xD6, 0xD7, 0xD8, 0xD9, 0xDA, 0xE1, 0xE2, 0xE3, 0xE4, 0xE5, 0xE6, 0xE7,
            0xE8, 0xE9, 0xEA, 0xF1, 0xF2, 0xF3, 0xF4, 0xF5, 0xF6, 0xF7, 0xF8, 0xF9,
            0xFA, 0xFF, 0xDA, 0x00, 0x08, 0x01, 0x01, 0x00, 0x00, 0x3F, 0x00, 0xFB,
            0xD5, 0xDB, 0x20, 0xB8, 0xF7, 0xFF, 0xD9
        };
        File.WriteAllBytes(path, jpegBytes);
    }

    private static void CreateMinimalPng(string path)
    {
        // Minimal valid PNG: 1x1 pixel, transparent
        var pngBytes = new byte[]
        {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
            0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
            0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41,
            0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
            0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
            0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
            0x42, 0x60, 0x82
        };
        File.WriteAllBytes(path, pngBytes);
    }

    #endregion

    #region Task 1 — ReplaceText preserves styling

    [Fact]
    public void ReplaceText_PreservesRunAndParagraphFormatting()
    {
        // Arrange: styled shape (bold 24pt run, centered paragraph)
        var path = CreateDeckWithText("Original");
        uint shapeId;
        using (var doc = PresentationDocument.Open(path, true))
        {
            var shape = FirstShape(doc);
            shapeId = ReadElementId(shape.NonVisualShapeProperties!);
            var paragraph = shape.TextBody!.Elements<Drawing.Paragraph>().First();
            paragraph.ParagraphProperties = new Drawing.ParagraphProperties
            {
                Alignment = Drawing.TextAlignmentTypeValues.Center
            };
            paragraph.Elements<Drawing.Run>().First().RunProperties =
                new Drawing.RunProperties { Bold = true, FontSize = 2400 };
            doc.Save();
        }

        // Act
        PptxReplaceResult result;
        using (var doc = PresentationDocument.Open(path, true))
        {
            result = _replacer.ReplaceText(FirstSlidePart(doc), shapeId, "Replacement");
            doc.Save();
        }

        // Assert
        Assert.True(result.Success, result.Error);
        using (var doc = PresentationDocument.Open(path, false))
        {
            var paragraphs = FirstShape(doc).TextBody!.Elements<Drawing.Paragraph>().ToList();
            var paragraph = Assert.Single(paragraphs);
            Assert.Equal(
                Drawing.TextAlignmentTypeValues.Center,
                paragraph.ParagraphProperties?.Alignment?.Value);
            var run = paragraph.Elements<Drawing.Run>().Single();
            Assert.Equal("Replacement", run.Text?.Text);
            Assert.True(run.RunProperties?.Bold?.Value);
            Assert.Equal(2400, run.RunProperties?.FontSize?.Value);
        }
    }

    [Fact]
    public void ReplaceText_MultiLine_PreservesPerParagraphFormatting()
    {
        // Arrange: two paragraphs with distinct pPr (left/bold, right/italic)
        var path = CreateDeckWithText("First");
        uint shapeId;
        using (var doc = PresentationDocument.Open(path, true))
        {
            var shape = FirstShape(doc);
            shapeId = ReadElementId(shape.NonVisualShapeProperties!);
            var textBody = shape.TextBody!;

            var p1 = textBody.Elements<Drawing.Paragraph>().First();
            p1.ParagraphProperties = new Drawing.ParagraphProperties
            {
                Alignment = Drawing.TextAlignmentTypeValues.Left
            };
            p1.Elements<Drawing.Run>().First().RunProperties =
                new Drawing.RunProperties { Bold = true };

            var p2Run = new Drawing.Run(new Drawing.Text("Second"));
            p2Run.RunProperties = new Drawing.RunProperties { Italic = true };
            var p2 = new Drawing.Paragraph(p2Run)
            {
                ParagraphProperties = new Drawing.ParagraphProperties
                {
                    Alignment = Drawing.TextAlignmentTypeValues.Right
                }
            };
            textBody.Append(p2);
            doc.Save();
        }

        // Act: three lines — the third must reuse the last paragraph template
        PptxReplaceResult result;
        using (var doc = PresentationDocument.Open(path, true))
        {
            result = _replacer.ReplaceText(FirstSlidePart(doc), shapeId, "Line1\nLine2\nLine3");
            doc.Save();
        }

        // Assert
        Assert.True(result.Success, result.Error);
        using (var doc = PresentationDocument.Open(path, false))
        {
            var paragraphs = FirstShape(doc).TextBody!.Elements<Drawing.Paragraph>().ToList();
            Assert.Equal(3, paragraphs.Count);

            Assert.Equal(Drawing.TextAlignmentTypeValues.Left,
                paragraphs[0].ParagraphProperties?.Alignment?.Value);
            Assert.True(paragraphs[0].Elements<Drawing.Run>().Single().RunProperties?.Bold?.Value);
            Assert.Equal("Line1", paragraphs[0].Elements<Drawing.Run>().Single().Text?.Text);

            Assert.Equal(Drawing.TextAlignmentTypeValues.Right,
                paragraphs[1].ParagraphProperties?.Alignment?.Value);
            Assert.True(paragraphs[1].Elements<Drawing.Run>().Single().RunProperties?.Italic?.Value);
            Assert.Equal("Line2", paragraphs[1].Elements<Drawing.Run>().Single().Text?.Text);

            Assert.Equal(Drawing.TextAlignmentTypeValues.Right,
                paragraphs[2].ParagraphProperties?.Alignment?.Value);
            Assert.True(paragraphs[2].Elements<Drawing.Run>().Single().RunProperties?.Italic?.Value);
            Assert.Equal("Line3", paragraphs[2].Elements<Drawing.Run>().Single().Text?.Text);
        }
    }

    #endregion

    #region Task 2 — ReplaceImage shared-part safety + content type

    [Fact]
    public void ReplaceImage_KeepsOldPart_WhenSharedWithAnotherSlide()
    {
        // Arrange: deck where slide 2 references the SAME ImagePart as slide 1
        var path = NewDeckPath();
        var jpgPath = NewImagePath(".jpg");
        var pngPath = NewImagePath(".png");
        using (var builder = PresentationBuilder.Create(path))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddImage(jpgPath);
            builder.AddSlide();
            builder.Save();
        }

        uint pictureId;
        string slide2EmbedId;
        using (var doc = PresentationDocument.Open(path, true))
        {
            var slideParts = doc.PresentationPart!.SlideParts.ToList();
            var picture = FirstPicture(slideParts[0]);
            pictureId = ReadElementId(picture.NonVisualPictureProperties!);
            var sharedPart = slideParts[0].ImageParts.Single();

            slideParts[1].AddPart(sharedPart);
            slide2EmbedId = slideParts[1].GetIdOfPart(sharedPart);
            var clone = (P.Picture)picture.CloneNode(true);
            clone.BlipFill!.Blip!.Embed = slide2EmbedId;
            slideParts[1].Slide!.CommonSlideData!.ShapeTree!.Append(clone);
            doc.Save();
        }

        // Act
        PptxReplaceResult result;
        using (var doc = PresentationDocument.Open(path, true))
        {
            result = _replacer.ReplaceImage(FirstSlidePart(doc), pictureId, pngPath);
            doc.Save();
        }

        // Assert
        Assert.True(result.Success, result.Error);
        using (var doc = PresentationDocument.Open(path, false))
        {
            var slideParts = doc.PresentationPart!.SlideParts.ToList();

            // Slide 2's relationship must still resolve — the old part was NOT deleted.
            var oldPart = slideParts[1].GetPartById(slide2EmbedId);
            Assert.Equal("image/jpeg", oldPart.ContentType);
            using (var stream = oldPart.GetStream())
            {
                Assert.True(stream.Length > 0);
            }

            // Slide 1 now points at a new PNG part.
            var newEmbedId = FirstPicture(slideParts[0]).BlipFill!.Blip!.Embed!.Value;
            var newPart = slideParts[0].GetPartById(newEmbedId!);
            Assert.Equal("image/png", newPart.ContentType);
        }
    }

    [Fact]
    public void ReplaceImage_DeletesOldPart_AndDerivesContentType_WhenNotShared()
    {
        // Arrange
        var path = CreateDeckWithImage(out var pictureId);
        var pngPath = NewImagePath(".png");

        // Act
        PptxReplaceResult result;
        using (var doc = PresentationDocument.Open(path, true))
        {
            result = _replacer.ReplaceImage(FirstSlidePart(doc), pictureId, pngPath);
            doc.Save();
        }

        // Assert
        Assert.True(result.Success, result.Error);
        using (var doc = PresentationDocument.Open(path, false))
        {
            // Old JPEG part must be gone — only the new PNG part remains.
            var imageParts = doc.PresentationPart!.SlideParts
                .SelectMany(sp => sp.ImageParts)
                .ToList();
            var only = Assert.Single(imageParts);
            Assert.Equal("image/png", only.ContentType);

            var embedId = FirstPicture(FirstSlidePart(doc)).BlipFill!.Blip!.Embed!.Value;
            Assert.Equal("image/png", FirstSlidePart(doc).GetPartById(embedId!).ContentType);
        }
    }

    [Fact]
    public void ReplaceImage_Throws_WhenFileMissing()
    {
        var path = CreateDeckWithImage(out var pictureId);
        using var doc = PresentationDocument.Open(path, true);
        var missing = Path.Combine(_testDir, "does-not-exist.jpg");
        Assert.Throws<FileNotFoundException>(
            () => _replacer.ReplaceImage(FirstSlidePart(doc), pictureId, missing));
    }

    #endregion

    #region Task 3 — ReplaceTableData style preservation

    [Fact]
    public void ReplaceTableData_SameDimensions_EditsInPlace_PreservingCellAndRunProperties()
    {
        // Arrange: 2x2 table; mark cell(0,0) with tcPr margin + bold run
        var path = CreateDeckWithTable(2, 2, out var frameId);
        using (var doc = PresentationDocument.Open(path, true))
        {
            var cell = FirstTable(FirstGraphicFrame(FirstSlidePart(doc)))
                .Elements<Drawing.TableRow>().First()
                .Elements<Drawing.TableCell>().First();
            cell.TableCellProperties!.SetAttribute(new OpenXmlAttribute("marL", "", "123456"));
            cell.TextBody!.Elements<Drawing.Paragraph>().First()
                .Elements<Drawing.Run>().First().RunProperties =
                new Drawing.RunProperties { Bold = true };
            doc.Save();
        }

        // Act: same 2x2 dimensions
        PptxReplaceResult result;
        using (var doc = PresentationDocument.Open(path, true))
        {
            result = _replacer.ReplaceTableData(FirstSlidePart(doc), frameId, new List<List<string>>
            {
                new() { "New1", "New2" },
                new() { "New3", "New4" }
            });
            doc.Save();
        }

        // Assert
        Assert.True(result.Success, result.Error);
        using (var doc = PresentationDocument.Open(path, false))
        {
            var rows = FirstTable(FirstGraphicFrame(FirstSlidePart(doc)))
                .Elements<Drawing.TableRow>().ToList();
            var firstCell = rows[0].Elements<Drawing.TableCell>().First();

            // Cell properties and run formatting survived → edit was in place.
            var marL = Regex.Match(firstCell.TableCellProperties!.OuterXml,
                @"\bmarL\s*=\s*""([^""]*)""");
            Assert.True(marL.Success, "marL attribute missing after in-place edit");
            Assert.Equal("123456", marL.Groups[1].Value);
            var run = firstCell.TextBody!.Elements<Drawing.Paragraph>().First()
                .Elements<Drawing.Run>().Single();
            Assert.True(run.RunProperties?.Bold?.Value);
            Assert.Equal("New1", run.Text?.Text);
            Assert.Equal("New4", rows[1].Elements<Drawing.TableCell>().ElementAt(1)
                .TextBody!.Elements<Drawing.Paragraph>().First()
                .Elements<Drawing.Run>().Single().Text?.Text);
        }
    }

    [Fact]
    public void ReplaceTableData_DifferentDimensions_ClonesTablePropertiesGridAndRowHeights()
    {
        // Arrange: 2x2 table with custom style id, custom grid widths (total 3600000),
        // custom row heights, and a known frame transform.
        var path = CreateDeckWithTable(2, 2, out var frameId);
        using (var doc = PresentationDocument.Open(path, true))
        {
            var frame = FirstGraphicFrame(FirstSlidePart(doc));
            var table = FirstTable(frame);
            table.TableProperties!.Append(new Drawing.TableStyleId(TestTableStyleId));
            var columns = table.TableGrid!.Elements<Drawing.GridColumn>().ToList();
            columns[0].Width = 2000000;
            columns[1].Width = 1600000;
            foreach (var row in table.Elements<Drawing.TableRow>())
            {
                row.Height = 500000;
            }
            doc.Save();
        }

        // Act: different column count → rebuild path
        PptxReplaceResult result;
        using (var doc = PresentationDocument.Open(path, true))
        {
            result = _replacer.ReplaceTableData(FirstSlidePart(doc), frameId, new List<List<string>>
            {
                new() { "A1", "A2", "A3" },
                new() { "B1", "B2", "B3" }
            });
            doc.Save();
        }

        // Assert
        Assert.True(result.Success, result.Error);
        using (var doc = PresentationDocument.Open(path, false))
        {
            var frame = FirstGraphicFrame(FirstSlidePart(doc));
            var table = FirstTable(frame);

            // tableStyleId survived the rebuild.
            Assert.Equal(TestTableStyleId,
                table.TableProperties!.Elements<Drawing.TableStyleId>().Single().Text);

            // Grid columns derive from the original total width (3600000),
            // not the old hardcoded 7200000/cols split.
            var widths = table.TableGrid!.Elements<Drawing.GridColumn>()
                .Select(c => c.Width!.Value).ToList();
            Assert.Equal(3, widths.Count);
            Assert.Equal(3600000, widths.Sum());
            Assert.DoesNotContain(7200000 / 3, widths);

            // Row heights come from the original rows, not hardcoded 370840.
            Assert.All(table.Elements<Drawing.TableRow>(), row => Assert.Equal(500000, row.Height?.Value));

            // Frame transform (xfrm) untouched — only a:tbl was swapped.
            Assert.Equal(0, frame.Transform!.Offset!.X!.Value);
            Assert.Equal(1440000, frame.Transform.Offset.Y!.Value);
            Assert.Equal(7200000, frame.Transform.Extents!.Cx!.Value);
            Assert.Equal(3600000, frame.Transform.Extents.Cy!.Value);

            var text = table.InnerText;
            Assert.Contains("A1", text);
            Assert.Contains("B3", text);
            Assert.DoesNotContain("R0C0", text);
        }
    }

    #endregion

    #region Task 4 — Per-op success/failure results

    [Fact]
    public void ReplaceText_ReturnsFailure_WhenElementNotFound()
    {
        var path = CreateDeckWithText("Original");
        using var doc = PresentationDocument.Open(path, true);

        var result = _replacer.ReplaceText(FirstSlidePart(doc), 9999u, "Replacement");

        Assert.False(result.Success);
        Assert.Contains("9999", result.Error);
    }

    [Fact]
    public void ReplaceText_ReturnsFailure_WhenMultipleElementsShareId()
    {
        // Arrange: two shapes forced to the same cNvPr id
        var path = NewDeckPath();
        using (var builder = PresentationBuilder.Create(path))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddText("One");
            builder.CurrentSlide.AddText("Two");
            builder.Save();
        }
        uint duplicatedId;
        using (var doc = PresentationDocument.Open(path, true))
        {
            var shapes = FirstSlidePart(doc).Slide!.CommonSlideData!.ShapeTree!
                .Elements<P.Shape>().ToList();
            Assert.True(shapes.Count >= 2, "expected at least two shapes");
            duplicatedId = ReadElementId(shapes[0].NonVisualShapeProperties!);
            var secondCNvPr = shapes[1].NonVisualShapeProperties!.ChildElements
                .First(e => e.LocalName == "cNvPr");
            secondCNvPr.SetAttribute(new OpenXmlAttribute("id", "", duplicatedId.ToString()));
            doc.Save();
        }

        // Act
        PptxReplaceResult result;
        using (var doc = PresentationDocument.Open(path, true))
        {
            result = _replacer.ReplaceText(FirstSlidePart(doc), duplicatedId, "Replacement");
        }

        // Assert: fails cleanly instead of picking one arbitrarily
        Assert.False(result.Success);
        Assert.Contains("Multiple", result.Error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ReplaceImage_ReturnsFailure_WhenElementNotFound()
    {
        var path = CreateDeckWithImage(out _);
        var pngPath = NewImagePath(".png");
        using var doc = PresentationDocument.Open(path, true);

        var result = _replacer.ReplaceImage(FirstSlidePart(doc), 9999u, pngPath);

        Assert.False(result.Success);
        Assert.Contains("9999", result.Error);
    }

    [Fact]
    public void ReplaceTableData_ReturnsFailure_WhenElementNotFound()
    {
        var path = CreateDeckWithTable(1, 1, out _);
        using var doc = PresentationDocument.Open(path, true);

        var result = _replacer.ReplaceTableData(FirstSlidePart(doc), 9999u,
            new List<List<string>> { new() { "X" } });

        Assert.False(result.Success);
        Assert.Contains("9999", result.Error);
    }

    [Fact]
    public void ReplaceTableData_ReturnsFailure_WhenDataIsEmpty()
    {
        var path = CreateDeckWithTable(1, 1, out var frameId);
        using var doc = PresentationDocument.Open(path, true);

        var result = _replacer.ReplaceTableData(FirstSlidePart(doc), frameId, new List<List<string>>());

        Assert.False(result.Success);
        Assert.NotNull(result.Error);
    }

    [Fact]
    public void ReplaceTableData_ReturnsFailure_WhenRowsAreRagged()
    {
        var path = CreateDeckWithTable(1, 1, out var frameId);
        using var doc = PresentationDocument.Open(path, true);

        var result = _replacer.ReplaceTableData(FirstSlidePart(doc), frameId, new List<List<string>>
        {
            new() { "A", "B" },
            new() { "C" }
        });

        Assert.False(result.Success);
        Assert.NotNull(result.Error);
    }

    #endregion

    public void Dispose()
    {
        if (Directory.Exists(_testDir))
        {
            Directory.Delete(_testDir, recursive: true);
        }
    }
}
