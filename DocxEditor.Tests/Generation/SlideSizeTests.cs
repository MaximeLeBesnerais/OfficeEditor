using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using PptxEditor.Core.Generation.Emit.Ooxml;
using PptxEditor.Core.Generation.Emit.Typst;
using PptxEditor.Core.Generation.Layout;
using PptxEditor.Core.Generation.Model;
using PptxEditor.Core.Generation.Schema;

namespace DocxEditor.Tests.Generation;

/// <summary>
/// Deck-level slide canvas size ('slideSize'): the preset strings ("16:9" | "4:3") keep
/// working, an object form ({"width": pt, "height": pt}) adds arbitrary dimensions, and
/// both flow end-to-end into the emitted .pptx (sldSz EMU) and Typst preview page size.
/// </summary>
public class SlideSizeTests
{
    private readonly GenerationDocumentParser _parser = new();

    /// <summary>A minimal valid deck; 'slideSize' is injected per test.</summary>
    private const string Deck = """
        {
          "version": "2.0",
          "design": { "palette": { "primary": "#0B3D91", "ink": "#1A1A1A", "paper": "#FFFFFF" } },
          "slides": [ { "type": "container", "children": [] } ]
        }
        """;

    private static string WithSlideSize(string slideSizeJson) =>
        Deck.Replace("\"version\": \"2.0\",", $"\"version\": \"2.0\",\n  \"slideSize\": {slideSizeJson},");

    #region Parser: accepted forms

    [Fact]
    public void Parse_SlideSizeObject_Square540_UsesCustomCanvas()
    {
        var doc = _parser.Parse(WithSlideSize("""{ "width": 540, "height": 540 }"""));

        Assert.Equal(540, doc.SlideSize.WidthPt);
        Assert.Equal(540, doc.SlideSize.HeightPt);
        Assert.Equal(SlideSize.MaxDimPt, 4032);
        Assert.Equal(SlideSize.MinDimPt, 1);
    }

    [Fact]
    public void Parse_SlideSizeObject_NonSquare800x600_UsesCustomCanvas()
    {
        var doc = _parser.Parse(WithSlideSize("""{ "width": 800, "height": 600 }"""));

        Assert.Equal(800, doc.SlideSize.WidthPt);
        Assert.Equal(600, doc.SlideSize.HeightPt);
    }

    [Fact]
    public void Parse_SlideSizeObject_BoundaryDims1And4032_AreAccepted()
    {
        var doc = _parser.Parse(WithSlideSize("""{ "width": 1, "height": 4032 }"""));

        Assert.Equal(1, doc.SlideSize.WidthPt);
        Assert.Equal(4032, doc.SlideSize.HeightPt);
    }

    [Fact]
    public void Parse_SlideSizeString16x9_StillUsesWidescreenCanvas()
    {
        var doc = _parser.Parse(WithSlideSize("\"16:9\""));

        Assert.Equal(SlideSize.Widescreen16x9, doc.SlideSize);
        Assert.Equal(960, doc.SlideSize.WidthPt);
        Assert.Equal(540, doc.SlideSize.HeightPt);
    }

    [Fact]
    public void Parse_SlideSizeString4x3_StillUsesStandardCanvas()
    {
        var doc = _parser.Parse(WithSlideSize("\"4:3\""));

        Assert.Equal(SlideSize.Standard4x3, doc.SlideSize);
        Assert.Equal(720, doc.SlideSize.WidthPt);
        Assert.Equal(540, doc.SlideSize.HeightPt);
    }

    [Fact]
    public void Parse_SlideSizeAbsent_DefaultsTo16x9()
    {
        var doc = _parser.Parse(Deck);

        Assert.Equal(SlideSize.Widescreen16x9, doc.SlideSize);
    }

    #endregion

    #region Parser: loud rejections

    [Fact]
    public void Validate_SlideSizeObject_MissingWidth_RejectsAtSlideSizeWidth()
    {
        var result = _parser.Validate(WithSlideSize("""{ "height": 540 }"""));

        Assert.False(result.IsValid);
        var error = Assert.Single(result.Errors);
        Assert.Equal("$.slideSize.width", error.Path);
        Assert.Contains("'width' is required", error.Message);
    }

    [Fact]
    public void Validate_SlideSizeObject_MissingHeight_RejectsAtSlideSizeHeight()
    {
        var result = _parser.Validate(WithSlideSize("""{ "width": 540 }"""));

        Assert.False(result.IsValid);
        var error = Assert.Single(result.Errors);
        Assert.Equal("$.slideSize.height", error.Path);
        Assert.Contains("'height' is required", error.Message);
    }

    [Fact]
    public void Validate_SlideSizeObject_MissingBothDims_RejectsBothPaths()
    {
        var result = _parser.Validate(WithSlideSize("{}"));

        Assert.False(result.IsValid);
        Assert.Equal(2, result.Errors.Count);
        Assert.Contains(result.Errors, e => e.Path == "$.slideSize.width" && e.Message.Contains("'width' is required"));
        Assert.Contains(result.Errors, e => e.Path == "$.slideSize.height" && e.Message.Contains("'height' is required"));
    }

    [Theory]
    [InlineData("""{ "width": "540", "height": 540 }""", "$.slideSize.width")]
    [InlineData("""{ "width": 540, "height": [600] }""", "$.slideSize.height")]
    [InlineData("""{ "width": true, "height": 540 }""", "$.slideSize.width")]
    public void Validate_SlideSizeObject_NonNumberDimension_RejectsAtPath(string slideSize, string expectedPath)
    {
        var result = _parser.Validate(WithSlideSize(slideSize));

        Assert.False(result.IsValid);
        var error = Assert.Single(result.Errors);
        Assert.Equal(expectedPath, error.Path);
        Assert.Contains("must be a finite JSON number in points", error.Message);
    }

    [Theory]
    [InlineData("""{ "width": 0, "height": 540 }""", "$.slideSize.width")]
    [InlineData("""{ "width": 540, "height": -10 }""", "$.slideSize.height")]
    public void Validate_SlideSizeObject_NonPositiveDimension_RejectsAtPath(string slideSize, string expectedPath)
    {
        var result = _parser.Validate(WithSlideSize(slideSize));

        Assert.False(result.IsValid);
        var error = Assert.Single(result.Errors);
        Assert.Equal(expectedPath, error.Path);
        Assert.Contains("must be ≥ 1 pt", error.Message);
    }

    [Fact]
    public void Validate_SlideSizeObject_OverMax_RejectsAtPath()
    {
        var result = _parser.Validate(WithSlideSize("""{ "width": 540, "height": 5000 }"""));

        Assert.False(result.IsValid);
        var error = Assert.Single(result.Errors);
        Assert.Equal("$.slideSize.height", error.Path);
        Assert.Contains("must be ≤ 4032 pt (got 5000)", error.Message);
    }

    [Fact]
    public void Validate_SlideSizeObject_ExtraProperty_RejectsAtPath()
    {
        var result = _parser.Validate(WithSlideSize("""{ "width": 540, "height": 540, "foo": 1 }"""));

        Assert.False(result.IsValid);
        var error = Assert.Single(result.Errors);
        Assert.Equal("$.slideSize.foo", error.Path);
        Assert.Contains("unknown property 'foo'", error.Message);
    }

    [Theory]
    [InlineData("540")]
    [InlineData("""[ "16:9" ]""")]
    [InlineData("true")]
    public void Validate_SlideSizeNonStringNonObjectType_Rejects(string slideSize)
    {
        var result = _parser.Validate(WithSlideSize(slideSize));

        Assert.False(result.IsValid);
        var error = Assert.Single(result.Errors);
        Assert.Equal("$.slideSize", error.Path);
        Assert.Contains("must be a string (\"16:9\" | \"4:3\") or an object", error.Message);
    }

    [Fact]
    public void Validate_SlideSizeUnknownPresetString_RejectsLikeBefore()
    {
        var result = _parser.Validate(WithSlideSize("\"1:1\""));

        Assert.False(result.IsValid);
        var error = Assert.Single(result.Errors);
        Assert.Equal("$.slideSize", error.Path);
        Assert.Contains("'1:1' is not a valid slide size", error.Message);
    }

    #endregion

    #region End-to-end: emitted .pptx slide size (no Typst)

    private static byte[] GeneratePptx(string json)
    {
        var document = new GenerationDocumentParser().Parse(json);
        var layout = new LayoutResolver().Resolve(document);
        return new OoxmlEmitter().Emit(layout).Bytes;
    }

    [Fact]
    public void Generate_Square540_EmittedSlideSizeIsSquareEmu()
    {
        var bytes = GeneratePptx(WithSlideSize("""{ "width": 540, "height": 540 }"""));

        using var package = PresentationDocument.Open(new MemoryStream(bytes), false);
        var slideSize = package.PresentationPart!.Presentation!.SlideSize!;
        Assert.Equal(540 * 12700, slideSize.Cx!.Value);
        Assert.Equal(540 * 12700, slideSize.Cy!.Value);
    }

    [Fact]
    public void Generate_Custom800x600_EmittedSlideSizeIsCustomEmu()
    {
        var bytes = GeneratePptx(WithSlideSize("""{ "width": 800, "height": 600 }"""));

        using var package = PresentationDocument.Open(new MemoryStream(bytes), false);
        var slideSize = package.PresentationPart!.Presentation!.SlideSize!;
        Assert.Equal(800 * 12700, slideSize.Cx!.Value);
        Assert.Equal(600 * 12700, slideSize.Cy!.Value);
    }

    [Fact]
    public void Generate_16x9_EmittedSlideSizeIsWidescreenEmu()
    {
        var bytes = GeneratePptx(WithSlideSize("\"16:9\""));

        using var package = PresentationDocument.Open(new MemoryStream(bytes), false);
        var slideSize = package.PresentationPart!.Presentation!.SlideSize!;
        Assert.Equal(960 * 12700, slideSize.Cx!.Value);
        Assert.Equal(540 * 12700, slideSize.Cy!.Value);
    }

    [Fact]
    public void Generate_4x3_EmittedSlideSizeIsStandardEmu()
    {
        var bytes = GeneratePptx(WithSlideSize("\"4:3\""));

        using var package = PresentationDocument.Open(new MemoryStream(bytes), false);
        var slideSize = package.PresentationPart!.Presentation!.SlideSize!;
        Assert.Equal(720 * 12700, slideSize.Cx!.Value);
        Assert.Equal(540 * 12700, slideSize.Cy!.Value);
    }

    [Fact]
    public void Generate_CustomSize_DeckOpensCleanly()
    {
        var bytes = GeneratePptx(WithSlideSize("""{ "width": 800, "height": 600 }"""));

        using var package = PresentationDocument.Open(new MemoryStream(bytes), false);
        var errors = new OpenXmlValidator().Validate(package).ToList();
        Assert.True(errors.Count == 0, string.Join("\n", errors.Select(e => $"{e.Path}: {e.Description}")));
    }

    #endregion

    #region Typst preview: page size follows SlideSize

    [Fact]
    public void Generate_CustomSize_TypstPreviewPageSizeComesFromSlideSize()
    {
        var document = new GenerationDocumentParser().Parse(WithSlideSize("""{ "width": 800, "height": 600 }"""));
        var layout = new LayoutResolver().Resolve(document);
        var source = new TypstEmitter().Emit(layout);

        Assert.Contains("#set page(width: 800pt, height: 600pt, margin: 0pt)", source);
    }

    #endregion
}
