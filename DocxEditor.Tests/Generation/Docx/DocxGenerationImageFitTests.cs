using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Generation;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocxEditor.Tests.Generation.Docx;

/// <summary>
/// Flow and positioned image emission: fit modes (contain/fill/stretch/crop), explicit vs
/// natural sizing, data-URI and local-file sources, drawing ids, and image-part relationships
/// (including per-owner deduplication).
/// </summary>
public class DocxGenerationImageFitTests
{
    private static string FitJson(string fit, string? crop = null, bool sized = true, string src = "REPLACED")
    {
        var cropPart = crop is null ? string.Empty : $", \"crop\": {crop}";
        var sizePart = sized ? ", \"width\": 100, \"height\": 100" : string.Empty;
        return $$"""
            {
              "version": "1.0",
              "design": { "palette": { "ink": "#1F2937" } },
              "sections": [ { "blocks": [
                { "type": "image", "src": "{{src}}", "fit": "{{fit}}"{{cropPart}}{{sizePart}} }
              ] } ]
            }
            """;
    }

    [Theory]
    [InlineData("fill", 1270000L, 1270000L, true)]
    [InlineData("stretch", 1270000L, 1270000L, false)]
    [InlineData("contain", 1270000L, 635000L, false)]
    public void FlowImage_EmitsExtentsPerFit(string fit, long expectedCx, long expectedCy, bool expectsSrcRect)
    {
        var src = TestImages.DataUriBase64(TestImages.Png(200, 100));
        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(FitJson(fit, src: src));

        var (cx, cy, srcRect) = ReadInlineDrawing(path);
        Assert.Equal(expectedCx, cx);
        Assert.Equal(expectedCy, cy);
        Assert.Equal(expectsSrcRect, srcRect is not null);

        if (expectsSrcRect)
        {
            Assert.Equal(25000, srcRect!.Left!.Value);
            Assert.Equal(25000, srcRect.Right!.Value);
            Assert.Equal(0, srcRect.Top!.Value);
            Assert.Equal(0, srcRect.Bottom!.Value);
        }
    }

    [Fact]
    public void FlowImage_CropHonorsAuthorCropRectangle()
    {
        var src = TestImages.DataUriBase64(TestImages.Png(200, 100));
        var json = FitJson("crop", crop: """{ "left": 0.1, "top": 0.1, "right": 0.1, "bottom": 0.1 }""", src: src);

        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(json);

        var (cx, cy, srcRect) = ReadInlineDrawing(path);
        Assert.Equal(1270000L, cx);
        Assert.Equal(1270000L, cy);
        Assert.NotNull(srcRect);
        Assert.Equal(10000, srcRect!.Left!.Value);
        Assert.Equal(10000, srcRect.Top!.Value);
        Assert.Equal(10000, srcRect.Right!.Value);
        Assert.Equal(10000, srcRect.Bottom!.Value);
    }

    [Fact]
    public void FlowImage_WithoutBoxUsesNaturalSizeAtDefaultDpi()
    {
        // 200×100 PNG with no pHYs → natural 150×75 pt at 96 DPI.
        var src = TestImages.DataUriBase64(TestImages.Png(200, 100));

        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(FitJson("fill", sized: false, src: src));

        var (cx, cy, _) = ReadInlineDrawing(path);
        Assert.Equal(1905000L, cx); // 150 pt
        Assert.Equal(952500L, cy);  // 75 pt
    }

    [Fact]
    public void FlowImage_PercentEncodedDataUri_IsAccepted()
    {
        var src = TestImages.DataUriPercentEncoded(TestImages.Png(16, 16));

        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(FitJson("fill", src: src));

        var mainPart = DocxTestHarness.OpenMainPart(path);
        Assert.Single(mainPart.ImageParts);
        Assert.Equal("image/png", mainPart.ImageParts.Single().ContentType);
    }

    [Fact]
    public void FlowImage_LocalFileSource_IsEmbedded()
    {
        using var temp = new TempDirectory();
        var asset = temp.File("photo.png");
        File.WriteAllBytes(asset, TestImages.Png(64, 32));

        // Allowed root = temp dir so the relative path resolves inside it.
        var json = $$"""
            {
              "version": "1.0",
              "design": { "palette": { "ink": "#1F2937" } },
              "sections": [ { "blocks": [
                { "type": "image", "src": "photo.png", "width": 50, "height": 50 }
              ] } ]
            }
            """;

        var output = temp.File("out.docx");
        new DocxGenerator().Generate(json, output, new DocxGeneratorOptions
        {
            ImageSourceOptions = new DocxEditor.Core.Generation.Assets.ImageSourceOptions { AllowedRoot = temp.Path }
        });

        var mainPart = DocxTestHarness.OpenMainPart(output);
        var part = Assert.Single(mainPart.ImageParts);
        Assert.Equal("image/png", part.ContentType);
        using var read = part.GetStream();
        var payload = new MemoryStream();
        read.CopyTo(payload);
        Assert.Equal(TestImages.Png(64, 32), payload.ToArray());
    }

    [Fact]
    public void SameSource_IsDeduplicatedWithinOneOwner()
    {
        var src = TestImages.DataUriBase64(TestImages.Png(20, 20));
        var json = $$"""
            {
              "version": "1.0",
              "design": { "palette": { "ink": "#1F2937" } },
              "sections": [ { "blocks": [
                { "type": "image", "src": "{{src}}", "width": 20, "height": 20 },
                { "type": "image", "src": "{{src}}", "width": 40, "height": 40 }
              ] } ]
            }
            """;

        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(json);
        var mainPart = DocxTestHarness.OpenMainPart(path);

        Assert.Single(mainPart.ImageParts);
        var drawings = mainPart.Document!.Descendants<Drawing>().ToList();
        Assert.Equal(2, drawings.Count);
        var embedIds = mainPart.Document
            .Descendants<A.Blip>()
            .Select(b => b.Embed!.Value)
            .ToList();
        Assert.Equal(2, embedIds.Count);
        Assert.Equal(embedIds[0], embedIds[1]);
    }

    [Fact]
    public void PositionedImage_FillWritesCoverCropAndExtents()
    {
        var src = TestImages.DataUriBase64(TestImages.Png(200, 100));
        var json = PositionedImageJson(src, "fill");

        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(json);
        var (cx, cy, srcRect, xOffset, yOffset) = ReadAnchoredImage(path);

        Assert.Equal(1270000L, cx);
        Assert.Equal(1270000L, cy);
        Assert.Equal(0L, xOffset); // anchor x = 0pt
        Assert.Equal(0L, yOffset); // anchor y = 0pt
        Assert.NotNull(srcRect);
        Assert.Equal(25000, srcRect!.Left!.Value);
        Assert.Equal(25000, srcRect.Right!.Value);
    }

    [Fact]
    public void PositionedImage_ContainShrinksBoxAndRecenters()
    {
        var src = TestImages.DataUriBase64(TestImages.Png(200, 100));
        var json = PositionedImageJson(src, "contain");

        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(json);
        var (cx, cy, srcRect, xOffset, yOffset) = ReadAnchoredImage(path);

        // Contained box is 100×50, centered vertically: y offset becomes 25pt.
        Assert.Equal(1270000L, cx);
        Assert.Equal(635000L, cy);
        Assert.Null(srcRect);
        Assert.Equal(0L, xOffset);
        Assert.Equal(317500L, yOffset); // 25pt → EMU
    }

    private static string PositionedImageJson(string src, string fit) => $$"""
        {
          "version": "1.0",
          "design": { "palette": { "ink": "#1F2937" } },
          "sections": [ {
            "blocks": [ { "type": "paragraph", "text": "p" } ],
            "positioned": [
              { "type": "image", "src": "{{src}}", "fit": "{{fit}}", "x": 0, "y": 0, "width": 100, "height": 100 }
            ]
          } ]
        }
        """;

    private static (long Cx, long Cy, A.SourceRectangle? SrcRect) ReadInlineDrawing(string path)
    {
        var mainPart = DocxTestHarness.OpenMainPart(path);
        var inline = mainPart.Document!.Descendants<DW.Inline>().Single();
        var extent = inline.Extent!;
        var srcRect = inline.Descendants<A.SourceRectangle>().SingleOrDefault();
        return (extent.Cx!.Value, extent.Cy!.Value, srcRect);
    }

    private static (long Cx, long Cy, A.SourceRectangle? SrcRect, long X, long Y) ReadAnchoredImage(string path)
    {
        var mainPart = DocxTestHarness.OpenMainPart(path);
        var anchor = mainPart.Document!.Descendants<DW.Anchor>().Single();
        var extent = anchor.Extent!;
        var srcRect = anchor.Descendants<A.SourceRectangle>().SingleOrDefault();
        var posH = anchor.Descendants<DW.HorizontalPosition>().Single();
        var posV = anchor.Descendants<DW.VerticalPosition>().Single();
        var xOffset = long.Parse(posH.Descendants<DW.PositionOffset>().Single().Text, System.Globalization.CultureInfo.InvariantCulture);
        var yOffset = long.Parse(posV.Descendants<DW.PositionOffset>().Single().Text, System.Globalization.CultureInfo.InvariantCulture);
        return (extent.Cx!.Value, extent.Cy!.Value, srcRect, xOffset, yOffset);
    }
}
