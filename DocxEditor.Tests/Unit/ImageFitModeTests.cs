using System.Buffers.Binary;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PptxEditor.Core.Builders;
using PptxEditor.Core.Instructions;
using PptxEditor.Core.Models;
using PptxEditor.Core.Services;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// F7 — image replacement fit modes (stretch|fill|crop|contain): header sniffer
/// fixtures, srcRect fit math, contain frame resize, srcRect replacement hygiene,
/// fallback warnings, and an end-to-end render check on the reference deck.
/// </summary>
public class ImageFitModeTests : IDisposable
{
    // Minimal VALID 1x1 transparent PNG (decodable — used where the bytes must
    // survive a real render; unlike the synthetic header-only fixtures).
    private static readonly byte[] RealPngBytes = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");

    private readonly string _testDir = SlideOpsTestHelpers.CreateTempDirectory(nameof(ImageFitModeTests));
    private readonly PptxElementReplacer _replacer = new();

    public void Dispose() => SlideOpsTestHelpers.BestEffortDelete(_testDir);

    #region Byte fixtures (header-only; OpenXML stores them opaquely)

    private static byte[] PngWithSize(int width, int height)
    {
        var png = new byte[33]; // signature(8) + length(4) + "IHDR"(4) + data(13) + crc(4)
        png[0] = 0x89; png[1] = 0x50; png[2] = 0x4E; png[3] = 0x47;
        png[4] = 0x0D; png[5] = 0x0A; png[6] = 0x1A; png[7] = 0x0A;
        png[11] = 0x0D; // IHDR payload length = 13
        png[12] = (byte)'I'; png[13] = (byte)'H'; png[14] = (byte)'D'; png[15] = (byte)'R';
        BinaryPrimitives.WriteInt32BigEndian(png.AsSpan(16), width);
        BinaryPrimitives.WriteInt32BigEndian(png.AsSpan(20), height);
        png[24] = 0x08; png[25] = 0x06; // bit depth 8, RGBA
        return png;
    }

    private static byte[] JpegWithSize(int width, int height)
    {
        var jpeg = new byte[31];
        jpeg[0] = 0xFF; jpeg[1] = 0xD8; // SOI
        jpeg[2] = 0xFF; jpeg[3] = 0xE0; // APP0 (JFIF), 16-byte segment
        BinaryPrimitives.WriteUInt16BigEndian(jpeg.AsSpan(4), 16);
        "JFIF\0"u8.CopyTo(jpeg.AsSpan(6));
        jpeg[20] = 0xFF; jpeg[21] = 0xC0; // SOF0
        BinaryPrimitives.WriteUInt16BigEndian(jpeg.AsSpan(22), 11);
        jpeg[24] = 0x08; // precision
        BinaryPrimitives.WriteUInt16BigEndian(jpeg.AsSpan(25), (ushort)height);
        BinaryPrimitives.WriteUInt16BigEndian(jpeg.AsSpan(27), (ushort)width);
        jpeg[29] = 0xFF; jpeg[30] = 0xD9; // EOI
        return jpeg;
    }

    private static byte[] GifWithSize(int width, int height)
    {
        var gif = new byte[13];
        "GIF89a"u8.CopyTo(gif.AsSpan());
        BinaryPrimitives.WriteUInt16LittleEndian(gif.AsSpan(6), (ushort)width);
        BinaryPrimitives.WriteUInt16LittleEndian(gif.AsSpan(8), (ushort)height);
        return gif;
    }

    private static byte[] BmpWithSize(int width, int height)
    {
        var bmp = new byte[26];
        bmp[0] = (byte)'B'; bmp[1] = (byte)'M';
        BinaryPrimitives.WriteInt32LittleEndian(bmp.AsSpan(18), width);
        BinaryPrimitives.WriteInt32LittleEndian(bmp.AsSpan(22), height);
        return bmp;
    }

    #endregion

    #region Deck helpers

    private string NewDeckPath() => Path.Combine(_testDir, $"deck_{Guid.NewGuid():N}.pptx");

    private static SlidePart FirstSlidePart(PresentationDocument doc) =>
        doc.PresentationPart!.SlideParts.First();

    private static P.Picture FindPicture(SlidePart slidePart, uint pictureId) =>
        slidePart.Slide.CommonSlideData!.ShapeTree!.Elements<P.Picture>()
            .Single(p => ReadElementId(p.NonVisualPictureProperties!) == pictureId);

    private static uint ReadElementId(DocumentFormat.OpenXml.OpenXmlElement nvProperties)
    {
        var cNvPr = nvProperties.ChildElements.First(e => e.LocalName == "cNvPr");
        var match = Regex.Match(cNvPr.OuterXml, @"\bid\s*=\s*""([^""]*)""");
        Assert.True(match.Success, "cNvPr id attribute not found");
        return uint.Parse(match.Groups[1].Value);
    }

    /// <summary>Creates a single-slide deck with one picture at the given frame (EMU).</summary>
    private string CreateDeckWithFramedImage(long cx, long cy, long offX, long offY, out uint pictureId)
    {
        var path = NewDeckPath();
        var seedPath = Path.Combine(_testDir, $"{Guid.NewGuid():N}.png");
        File.WriteAllBytes(seedPath, RealPngBytes);
        using (var builder = PresentationBuilder.Create(path))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddImage(seedPath);
            builder.Save();
        }
        using (var doc = PresentationDocument.Open(path, true))
        {
            var picture = FirstSlidePart(doc).Slide.CommonSlideData!.ShapeTree!
                .Elements<P.Picture>().First();
            var xfrm = picture.ShapeProperties!.Transform2D!;
            xfrm.Extents!.Cx = cx;
            xfrm.Extents.Cy = cy;
            xfrm.Offset!.X = offX;
            xfrm.Offset.Y = offY;
            pictureId = ReadElementId(picture.NonVisualPictureProperties!);
            doc.Save();
        }
        return path;
    }

    private PptxReplaceResult ReplaceWithFit(
        string path, uint pictureId, byte[] imageBytes, string extension,
        ImageFitMode fit, SourceRect? crop = null)
    {
        using var doc = PresentationDocument.Open(path, true);
        var result = _replacer.ReplaceImage(
            FirstSlidePart(doc), pictureId, new MemoryStream(imageBytes), extension, fit, crop);
        doc.Save();
        return result;
    }

    private static Drawing.SourceRectangle? ReadSourceRect(string path, uint pictureId)
    {
        using var doc = PresentationDocument.Open(path, false);
        return FindPicture(FirstSlidePart(doc), pictureId)
            .BlipFill?.Elements<Drawing.SourceRectangle>().SingleOrDefault();
    }

    private static (long X, long Y, long Cx, long Cy) ReadFrame(string path, uint pictureId)
    {
        using var doc = PresentationDocument.Open(path, false);
        var xfrm = FindPicture(FirstSlidePart(doc), pictureId).ShapeProperties!.Transform2D!;
        return (xfrm.Offset!.X!.Value, xfrm.Offset.Y!.Value,
                xfrm.Extents!.Cx!.Value, xfrm.Extents.Cy!.Value);
    }

    private static void SetExistingSourceRect(string path, uint pictureId, SourceRect rect)
    {
        using var doc = PresentationDocument.Open(path, true);
        var blipFill = FindPicture(FirstSlidePart(doc), pictureId).BlipFill!;
        var srcRect = new Drawing.SourceRectangle
        {
            Left = rect.Left, Top = rect.Top, Right = rect.Right, Bottom = rect.Bottom
        };
        blipFill.InsertAfter(srcRect, blipFill.Blip!);
        doc.Save();
    }

    #endregion

    #region Header sniffer

    [Fact]
    public void Sniffer_ReadsPngDimensions()
    {
        var dimensions = ImageHeaderSniffer.TryGetPixelDimensions(PngWithSize(640, 480));
        Assert.Equal((640, 480), dimensions);
    }

    [Fact]
    public void Sniffer_ReadsJpegDimensions()
    {
        var dimensions = ImageHeaderSniffer.TryGetPixelDimensions(JpegWithSize(2, 3));
        Assert.Equal((2, 3), dimensions);
    }

    [Fact]
    public void Sniffer_ReadsGifDimensions()
    {
        var dimensions = ImageHeaderSniffer.TryGetPixelDimensions(GifWithSize(320, 200));
        Assert.Equal((320, 200), dimensions);
    }

    [Fact]
    public void Sniffer_ReadsBmpDimensions()
    {
        var dimensions = ImageHeaderSniffer.TryGetPixelDimensions(BmpWithSize(1024, 768));
        Assert.Equal((1024, 768), dimensions);
    }

    [Fact]
    public void Sniffer_ReadsTopDownBmpHeightAsAbsolute()
    {
        var dimensions = ImageHeaderSniffer.TryGetPixelDimensions(BmpWithSize(100, -50));
        Assert.Equal((100, 50), dimensions);
    }

    [Fact]
    public void Sniffer_ReturnsNull_ForSvgGarbageAndTruncatedPayloads()
    {
        Assert.Null(ImageHeaderSniffer.TryGetPixelDimensions(
            Encoding.ASCII.GetBytes("<svg xmlns=\"http://www.w3.org/2000/svg\"/>")));
        Assert.Null(ImageHeaderSniffer.TryGetPixelDimensions(Array.Empty<byte>()));
        Assert.Null(ImageHeaderSniffer.TryGetPixelDimensions(new byte[] { 0x89, 0x50 })); // truncated PNG
        Assert.Null(ImageHeaderSniffer.TryGetPixelDimensions(RealPngBytes[..10]));
    }

    #endregion

    #region Fill (cover) srcRect math

    [Fact]
    public void Fill_LandscapeImageIntoPortraitFrame_CropsLeftRight()
    {
        // Frame 3600000x7200000 (aspect 0.5); image 4000x2000 (aspect 2.0)
        // → visible width fraction 0.25 → l = r = 37500.
        var path = CreateDeckWithFramedImage(3600000, 7200000, 500000, 1000000, out var pictureId);

        var result = ReplaceWithFit(path, pictureId, PngWithSize(4000, 2000), ".png", ImageFitMode.Fill);

        Assert.True(result.Success, result.Error);
        Assert.Empty(result.Warnings);
        var srcRect = ReadSourceRect(path, pictureId);
        Assert.NotNull(srcRect);
        Assert.Equal(37500, srcRect.Left!.Value);
        Assert.Equal(37500, srcRect.Right!.Value);
        Assert.Equal(0, srcRect.Top!.Value);
        Assert.Equal(0, srcRect.Bottom!.Value);
        // Frame untouched: only Contain may move a:xfrm.
        Assert.Equal((500000, 1000000, 3600000L, 7200000L), ReadFrame(path, pictureId));
    }

    [Fact]
    public void Fill_PortraitImageIntoLandscapeFrame_CropsTopBottom()
    {
        // Frame 7200000x3600000 (aspect 2.0); image 2000x4000 (aspect 0.5)
        // → visible height fraction 0.25 → t = b = 37500.
        var path = CreateDeckWithFramedImage(7200000, 3600000, 0, 1440000, out var pictureId);

        var result = ReplaceWithFit(path, pictureId, PngWithSize(2000, 4000), ".png", ImageFitMode.Fill);

        Assert.True(result.Success, result.Error);
        var srcRect = ReadSourceRect(path, pictureId);
        Assert.NotNull(srcRect);
        Assert.Equal(0, srcRect.Left!.Value);
        Assert.Equal(0, srcRect.Right!.Value);
        Assert.Equal(37500, srcRect.Top!.Value);
        Assert.Equal(37500, srcRect.Bottom!.Value);
    }

    [Fact]
    public void Fill_SquareImageIntoSquareFrame_RemovesExistingSrcRect()
    {
        // Aspects match (1.0 == 1.0): no crop, and any stale srcRect is cleared.
        var path = CreateDeckWithFramedImage(4000000, 4000000, 0, 0, out var pictureId);
        SetExistingSourceRect(path, pictureId, new SourceRect(11111, 22222, 33333, 44444));

        var result = ReplaceWithFit(path, pictureId, PngWithSize(1000, 1000), ".png", ImageFitMode.Fill);

        Assert.True(result.Success, result.Error);
        Assert.Null(ReadSourceRect(path, pictureId));
    }

    [Fact]
    public void Fill_RoundsCropFractions()
    {
        // Image 3000x2000 (aspect 1.5) into a square frame: visible width = 2/3,
        // each side crops 1/6 = 0.166666… × 100000 = 16666.67 → rounds to 16667.
        var path = CreateDeckWithFramedImage(4000000, 4000000, 0, 0, out var pictureId);

        var result = ReplaceWithFit(path, pictureId, PngWithSize(3000, 2000), ".png", ImageFitMode.Fill);

        Assert.True(result.Success, result.Error);
        var srcRect = ReadSourceRect(path, pictureId);
        Assert.NotNull(srcRect);
        Assert.Equal(16667, srcRect.Left!.Value);
        Assert.Equal(16667, srcRect.Right!.Value);
        Assert.Equal(0, srcRect.Top!.Value);
        Assert.Equal(0, srcRect.Bottom!.Value);
    }

    [Fact]
    public void Fill_ReplacesExistingSrcRect_DoesNotStack()
    {
        // A previously cropped picture re-replaced with a computed fill must end
        // up with exactly ONE srcRect carrying the new values.
        var path = CreateDeckWithFramedImage(3600000, 7200000, 0, 0, out var pictureId);
        SetExistingSourceRect(path, pictureId, new SourceRect(11111, 11111, 11111, 11111));

        var result = ReplaceWithFit(path, pictureId, PngWithSize(4000, 2000), ".png", ImageFitMode.Fill);

        Assert.True(result.Success, result.Error);
        var srcRect = ReadSourceRect(path, pictureId);
        Assert.NotNull(srcRect);
        Assert.Equal(37500, srcRect.Left!.Value);
        Assert.Equal(0, srcRect.Top!.Value);

        using var doc = PresentationDocument.Open(path, false);
        var picture = FindPicture(FirstSlidePart(doc), pictureId);
        Assert.Single(picture.BlipFill!.Elements<Drawing.SourceRectangle>());

        // Schema validity: srcRect must sit between a:blip and a:stretch.
        var validator = new OpenXmlValidator();
        Assert.Empty(validator.Validate(doc));
    }

    #endregion

    #region Crop

    [Fact]
    public void Crop_WritesCallerRectVerbatim()
    {
        var path = CreateDeckWithFramedImage(7200000, 3600000, 0, 1440000, out var pictureId);
        var crop = new SourceRect(10000, 20000, 30000, 40000);

        var result = ReplaceWithFit(path, pictureId, PngWithSize(4000, 2000), ".png", ImageFitMode.Crop, crop);

        Assert.True(result.Success, result.Error);
        var srcRect = ReadSourceRect(path, pictureId);
        Assert.NotNull(srcRect);
        Assert.Equal(10000, srcRect.Left!.Value);
        Assert.Equal(20000, srcRect.Top!.Value);
        Assert.Equal(30000, srcRect.Right!.Value);
        Assert.Equal(40000, srcRect.Bottom!.Value);
    }

    [Fact]
    public void Crop_NullRect_BehavesLikeFill()
    {
        var path = CreateDeckWithFramedImage(3600000, 7200000, 0, 0, out var pictureId);

        var result = ReplaceWithFit(path, pictureId, PngWithSize(4000, 2000), ".png", ImageFitMode.Crop, crop: null);

        Assert.True(result.Success, result.Error);
        var srcRect = ReadSourceRect(path, pictureId);
        Assert.NotNull(srcRect);
        Assert.Equal(37500, srcRect.Left!.Value);
        Assert.Equal(37500, srcRect.Right!.Value);
    }

    #endregion

    #region Contain

    [Fact]
    public void Contain_LandscapeImageIntoSquareFrame_ShrinksCyAroundCenter()
    {
        // Image 2000x1000 (aspect 2.0) into a 3600000x3600000 frame at (1000000, 2000000):
        // width-limited → cy 3600000 → 1800000, y shifts by half the delta (+900000).
        var path = CreateDeckWithFramedImage(3600000, 3600000, 1000000, 2000000, out var pictureId);

        var result = ReplaceWithFit(path, pictureId, PngWithSize(2000, 1000), ".png", ImageFitMode.Contain);

        Assert.True(result.Success, result.Error);
        Assert.Empty(result.Warnings);
        Assert.Null(ReadSourceRect(path, pictureId)); // Contain never carries a srcRect

        var (x, y, cx, cy) = ReadFrame(path, pictureId);
        Assert.Equal(1000000, x);
        Assert.Equal(2900000, y);
        Assert.Equal(3600000, cx);
        Assert.Equal(1800000, cy);
        // Center preserved: (2800000, 3800000) before and after.
        Assert.Equal(1000000 + 3600000 / 2, x + cx / 2);
        Assert.Equal(2000000 + 3600000 / 2, y + cy / 2);
    }

    [Fact]
    public void Contain_PortraitImageIntoSquareFrame_ShrinksCxAroundCenter()
    {
        // Image 1000x2000 (aspect 0.5): height-limited → cx 3600000 → 1800000, x +900000.
        var path = CreateDeckWithFramedImage(3600000, 3600000, 1000000, 2000000, out var pictureId);

        var result = ReplaceWithFit(path, pictureId, PngWithSize(1000, 2000), ".png", ImageFitMode.Contain);

        Assert.True(result.Success, result.Error);
        var (x, y, cx, cy) = ReadFrame(path, pictureId);
        Assert.Equal(1900000, x);
        Assert.Equal(2000000, y);
        Assert.Equal(1800000, cx);
        Assert.Equal(3600000, cy);
        Assert.Equal(1000000 + 3600000 / 2, x + cx / 2);
        Assert.Equal(2000000 + 3600000 / 2, y + cy / 2);
    }

    [Fact]
    public void Contain_MatchingAspect_LeavesFrameAlone()
    {
        var path = CreateDeckWithFramedImage(7200000, 3600000, 500000, 1440000, out var pictureId);

        var result = ReplaceWithFit(path, pictureId, PngWithSize(2000, 1000), ".png", ImageFitMode.Contain);

        Assert.True(result.Success, result.Error);
        Assert.Equal((500000, 1440000, 7200000L, 3600000L), ReadFrame(path, pictureId));
        Assert.Null(ReadSourceRect(path, pictureId));
    }

    #endregion

    #region Stretch hygiene + tile normalization + fallbacks

    [Fact]
    public void Stretch_RemovesExistingSrcRect_AndLeavesFrameUntouched()
    {
        // A picture previously cropped then re-replaced with Stretch comes out clean.
        var path = CreateDeckWithFramedImage(3600000, 7200000, 500000, 1000000, out var pictureId);
        SetExistingSourceRect(path, pictureId, new SourceRect(25000, 25000, 25000, 25000));

        var result = ReplaceWithFit(path, pictureId, PngWithSize(4000, 2000), ".png", ImageFitMode.Stretch);

        Assert.True(result.Success, result.Error);
        Assert.Empty(result.Warnings);
        Assert.Null(ReadSourceRect(path, pictureId));
        Assert.Equal((500000, 1000000, 3600000L, 7200000L), ReadFrame(path, pictureId));
    }

    [Fact]
    public void ReplaceImage_NormalizesTileFillToStretch()
    {
        var path = CreateDeckWithFramedImage(7200000, 3600000, 0, 1440000, out var pictureId);
        using (var doc = PresentationDocument.Open(path, true))
        {
            var blipFill = FindPicture(FirstSlidePart(doc), pictureId).BlipFill!;
            blipFill.RemoveAllChildren<Drawing.Stretch>();
            blipFill.Append(new Drawing.Tile());
            doc.Save();
        }

        var result = ReplaceWithFit(path, pictureId, PngWithSize(4000, 2000), ".png", ImageFitMode.Stretch);

        Assert.True(result.Success, result.Error);
        using (var doc = PresentationDocument.Open(path, false))
        {
            var blipFill = FindPicture(FirstSlidePart(doc), pictureId).BlipFill!;
            Assert.Empty(blipFill.Elements<Drawing.Tile>());
            Assert.Single(blipFill.Elements<Drawing.Stretch>());
        }
    }

    [Fact]
    public void SvgImage_FallsBackToStretch_WithWarning()
    {
        var path = CreateDeckWithFramedImage(3600000, 7200000, 500000, 1000000, out var pictureId);
        var svgBytes = Encoding.ASCII.GetBytes(
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"100\" height=\"50\"></svg>");

        var result = ReplaceWithFit(path, pictureId, svgBytes, ".svg", ImageFitMode.Fill);

        Assert.True(result.Success, result.Error);
        var warning = Assert.Single(result.Warnings);
        Assert.Contains("stretch", warning, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("dimensions", warning, StringComparison.OrdinalIgnoreCase);
        Assert.Null(ReadSourceRect(path, pictureId));
        Assert.Equal((500000, 1000000, 3600000L, 7200000L), ReadFrame(path, pictureId));
    }

    [Fact]
    public void PptxReplaceResult_Warnings_DefaultEmpty()
    {
        Assert.Empty(PptxReplaceResult.Ok().Warnings);
        Assert.Empty(PptxReplaceResult.Fail("nope").Warnings);
    }

    #endregion

    #region Instruction engine wiring

    [Fact]
    public void Engine_ReplaceImage_FitFill_AppliesComputedSrcRect()
    {
        // Real 1x1 PNG (aspect 1.0) into the default AddImage frame 7200000x3600000
        // (aspect 2.0): visible height 0.5 → t = b = 25000.
        var path = CreateDeckWithFramedImage(7200000, 3600000, 0, 1440000, out var pictureId);
        var engine = new PptxInstructionEngine();

        using (var builder = PresentationBuilder.Open(path))
        {
            var result = engine.Apply(builder, new PptxInstructionSet
            {
                Operations = new List<PptxInstruction>
                {
                    new PptxReplaceImageInstruction
                    {
                        Slide = 1,
                        ElementId = pictureId,
                        Image = Convert.ToBase64String(RealPngBytes),
                        Fit = "fill"
                    }
                }
            });
            Assert.True(result.Success, string.Join("; ", result.FailedOps.Select(e => e.Error)));
            builder.Save();
        }

        var srcRect = ReadSourceRect(path, pictureId);
        Assert.NotNull(srcRect);
        Assert.Equal(0, srcRect.Left!.Value);
        Assert.Equal(0, srcRect.Right!.Value);
        Assert.Equal(25000, srcRect.Top!.Value);
        Assert.Equal(25000, srcRect.Bottom!.Value);
    }

    [Fact]
    public void Engine_ReplaceImage_UnknownFit_FailsTheOp()
    {
        // Instructions built directly in code bypass the JSON parser's vocabulary check.
        var path = CreateDeckWithFramedImage(7200000, 3600000, 0, 1440000, out var pictureId);
        var engine = new PptxInstructionEngine();

        using var builder = PresentationBuilder.Open(path);
        var result = engine.Apply(builder, new PptxInstructionSet
        {
            Operations = new List<PptxInstruction>
            {
                new PptxReplaceImageInstruction
                {
                    Slide = 1,
                    ElementId = pictureId,
                    Image = Convert.ToBase64String(RealPngBytes),
                    Fit = "bogus"
                }
            }
        });

        Assert.False(result.Success);
        var error = Assert.Single(result.FailedOps);
        Assert.Contains("fit", error.Error, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Integration: reference deck still renders after fit=fill

    [Fact]
    public void Integration_ReplaceImageFitFill_PresProDeck_StillRenders()
    {
        var referencePath = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory, "..", "..", "..", "..", "examples", "REF", "PPTX", ".pptx"));
        Assert.True(File.Exists(referencePath), $"Reference PPTX not found: {referencePath}");

        // Work on a copy — never mutate the reference deck.
        var path = Path.Combine(_testDir, ".pptx");
        File.Copy(referencePath, path);

        int slideIndex1Based;
        uint pictureId;
        using (var builder = PresentationBuilder.Open(path))
        {
            var anatomy = builder.Analyze();
            var hit = anatomy
                .SelectMany(slide => slide.Elements
                    .Where(e => e.Type == "Image")
                    .Select(e => (slide.SlideIndex, e.Id)))
                .First();
            slideIndex1Based = hit.SlideIndex;
            pictureId = hit.Id;

            var engine = new PptxInstructionEngine();
            var result = engine.Apply(builder, new PptxInstructionSet
            {
                Operations = new List<PptxInstruction>
                {
                    new PptxReplaceImageInstruction
                    {
                        Slide = slideIndex1Based,
                        ElementId = pictureId,
                        Image = Convert.ToBase64String(RealPngBytes),
                        Fit = "fill"
                    }
                }
            });
            Assert.True(result.Success, string.Join("; ", result.FailedOps.Select(e => e.Error)));
            builder.Save();
        }

        // The edited deck must still convert: render the edited slide to a non-empty PNG.
        using (var builder = PresentationBuilder.Open(path))
        {
            var png = builder.ExportThumbnail(slideIndex1Based - 1);
            Assert.True(png.Length > 8, "Rendered PNG is suspiciously small.");
            Assert.Equal(0x89, png[0]);
            Assert.Equal((byte)'P', png[1]);
            Assert.Equal((byte)'N', png[2]);
            Assert.Equal((byte)'G', png[3]);
        }
    }

    #endregion
}
