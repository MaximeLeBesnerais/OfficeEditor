using System.Text.Json;
using DocxEditor.Tests.Unit;
using DocumentFormat.OpenXml.Packaging;
using OfficeEditor.Core.Rendering;
using OfficeEditor.Testing;
using PptxEditor.Core.Builders;
using PptxEditor.Core.Generation;
using PptxEditor.Core.Generation.Fixtures;
using PptxEditor.Core.Generation.Layout;

namespace DocxEditor.Tests.Generation;

public sealed class PptxGeneratorTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "pptx-generator-" + Guid.NewGuid().ToString("N"));

    public PptxGeneratorTests() => Directory.CreateDirectory(_directory);
    public void Dispose() => Directory.Delete(_directory, recursive: true);

    [Fact]
    public void Generate_Default_ExpandsArchetypesAndPreservesMetadataWithoutRendering()
    {
        var result = new PptxGenerator().Generate(GenerationContract.Json);
        Assert.True(result.Success, string.Join("; ", result.Errors));
        Assert.Equal(2, result.SlideCount);
        Assert.Empty(result.Previews);
        Assert.Null(result.PreviewError);
        Assert.NotEmpty(result.Warnings);
        Assert.True(result.TotalMilliseconds >= result.GenerationMilliseconds);
        Assert.Equal(800, result.Layout!.Slides[0].WidthPt);
        Assert.Equal("intro", result.Layout.Slides[0].Id);
        Assert.Equal("Speaker notes survive generation.", result.Layout.Slides[0].Notes);
        Assert.True(result.Layout.TryGetRect(1, "fitted", out var rect));
        Assert.Equal(40, rect.Width);
        using var stream = new MemoryStream(result.PptxBytes!);
        using var pptx = PresentationDocument.Open(stream, false);
        Assert.Contains("Shared generation", pptx.PresentationPart!.SlideParts.First().Slide!.InnerText);
        Assert.Contains("Speaker notes survive generation.", pptx.PresentationPart.SlideParts.First().NotesSlidePart!.NotesSlide!.InnerText);
    }

    [Fact]
    public void Generate_TextFit_UsesConfiguredFontsAndEmitsMeasuredSize()
    {
        TextFitTestFont.WriteFont(_directory, "font.ttf", "GeneratorTestFont");
        var json = TextDocument("shrink");
        var result = new PptxGenerator().Generate(json, new PptxGeneratorOptions { FontDirectory = _directory });
        Assert.True(result.Success, string.Join("; ", result.Errors));
        var text = Assert.IsType<ResolvedText>(Assert.Single(result.Layout!.Slides[0].Root.Children));
        Assert.True(text.FontScale < 1);
        using var stream = new MemoryStream(result.PptxBytes!);
        using var pptx = PresentationDocument.Open(stream, false);
        var run = Assert.Single(pptx.PresentationPart!.SlideParts.Single().Slide!.Descendants<DocumentFormat.OpenXml.Drawing.RunProperties>());
        Assert.Equal((int)Math.Round(30 * text.FontScale * 100, MidpointRounding.AwayFromZero), run.FontSize!.Value);
    }

    [Fact]
    public void Generate_TextOverflowError_IsRejectedBeforeEmission()
    {
        TextFitTestFont.WriteFont(_directory, "font.ttf", "GeneratorTestFont");
        var result = new PptxGenerator().Generate(TextDocument("error"), new PptxGeneratorOptions { FontDirectory = _directory });
        Assert.False(result.Success);
        Assert.Null(result.PptxBytes);
        Assert.Null(result.Layout);
        Assert.Contains("text does not fit", Assert.Single(result.Errors).Message);
    }

    [Fact]
    public void Generate_ReusedInstance_DoesNotRetainErrorsOrSlides()
    {
        var generator = new PptxGenerator();
        Assert.False(generator.Generate("{ malformed").Success);
        var first = generator.Generate(GenerationContract.Json);
        var second = generator.Generate(GenerationContract.Json);
        Assert.True(first.Success);
        Assert.True(second.Success);
        Assert.Equal(GenerationContract.Snapshot(first.PptxBytes!), GenerationContract.Snapshot(second.PptxBytes!));
    }

    [Fact]
    public void Generate_RelativeImage_RequiresExplicitDirectoryAndEmbedsThatFile()
    {
        var bytes = FixtureAssets.CreateCheckerboardPng();
        File.WriteAllBytes(Path.Combine(_directory, "image.png"), bytes);
        var json = ImageDocument("image.png");
        var missingRoot = new PptxGenerator().Generate(json);
        Assert.False(missingRoot.Success);
        Assert.Contains("relative path", Assert.Single(missingRoot.Errors).Message);
        var result = new PptxGenerator().Generate(json, new PptxGeneratorOptions { DocumentDirectory = _directory });
        Assert.True(result.Success);
        using var stream = new MemoryStream(result.PptxBytes!);
        using var pptx = PresentationDocument.Open(stream, false);
        using var image = pptx.PresentationPart!.SlideParts.Single().ImageParts.Single().GetStream();
        using var output = new MemoryStream();
        image.CopyTo(output);
        Assert.Equal(bytes, output.ToArray());
    }

    [Fact]
    public void Generate_ImageTraversal_IsRejected()
    {
        var result = new PptxGenerator().Generate(ImageDocument("../secret.png"), new PptxGeneratorOptions { DocumentDirectory = _directory });
        Assert.False(result.Success);
        Assert.Null(result.PptxBytes);
        Assert.Contains("outside", Assert.Single(result.Errors).Message);
    }

    [Fact]
    public void Generate_HostRoot_RejectsAbsoluteFilesOutsideRoot()
    {
        var result = new PptxGenerator().Generate(ImageDocument(Path.Combine(Path.GetTempPath(), "secret.png")),
            new PptxGeneratorOptions { AllowedImageRoot = _directory });
        Assert.False(result.Success);
        Assert.Contains("outside the allowed root", Assert.Single(result.Errors).Message);
    }

    [Fact]
    public void Generate_PreviewEmissionFailure_PreservesPptxAndReportsError()
    {
        var uri = "data:image/png;base64," + Convert.ToBase64String(FixtureAssets.CreateCheckerboardPng());
        var result = new PptxGenerator().Generate(ImageDocument(uri), new PptxGeneratorOptions { PreviewFormat = "svg" });
        Assert.True(result.Success);
        Assert.NotNull(result.PptxBytes);
        Assert.Empty(result.Previews);
        Assert.Contains("Typst emission failed", result.PreviewError);
    }

    [Theory]
    [InlineData("svg")]
    [InlineData("png")]
    public void Generate_Preview_UsesDocumentDirectory(string format)
    {
        File.WriteAllBytes(Path.Combine(_directory, "image.png"), FixtureAssets.CreateCheckerboardPng());
        var result = new PptxGenerator().Generate(ImageDocument("image.png"), new PptxGeneratorOptions
        {
            DocumentDirectory = _directory, PreviewFormat = format, Ppi = 36
        });
        Assert.True(result.Success);
        // This integration check requires a working backend, like DocumentRendererFacadeTests.
        Assert.Null(result.PreviewError);
        var preview = Assert.Single(result.Previews);
        Assert.Equal(format == "svg" ? "image/svg+xml" : "image/png", preview.ContentType);
        Assert.NotEmpty(preview.Bytes);
    }

    [Theory]
    [InlineData(float.NaN)]
    [InlineData(float.PositiveInfinity)]
    [InlineData(0f)]
    [InlineData(601f)]
    public void Generate_InvalidPpi_Throws(float ppi)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new PptxGenerator().Generate(GenerationContract.Json, new PptxGeneratorOptions { Ppi = ppi }));
    }

    [Fact]
    public void Render_Json_UsesSameGeneratedDeliveryAsCore()
    {
        var result = new PptxGenerator().Generate(GenerationContract.Json);
        Assert.True(result.Success);
        var source = Path.Combine(_directory, "deck.json");
        File.WriteAllText(source, GenerationContract.Json);
        using var expected = PresentationBuilder.Open(result.PptxBytes!);
        var actual = new DocumentRenderer().Render(new DocumentRenderRequest { SourcePath = source, Format = DocumentOutputFormat.Typ });
        Assert.True(actual.Success, actual.ErrorMessage);
        Assert.Equal(expected.ExportToTypst(), System.Text.Encoding.UTF8.GetString(Assert.Single(actual.Pages)));
    }

    private static string TextDocument(string overflow) => $$"""
        { "version": "2.0", "design": { "palette": { "primary": "#000000" } }, "slides": [
          { "type": "container", "children": [
            { "type": "text", "font": "GeneratorTestFont", "fontSize": 30,
              "text": "Many words that cannot fit", "overflow": "{{overflow}}",
              "at": { "x": 0, "y": 0 }, "size": { "w": 40, "h": 15 } }
          ] }
        ] }
        """;

    private static string ImageDocument(string source) => $$"""
        { "version": "2.0", "design": { "palette": { "primary": "#000000" } }, "slides": [
          { "type": "container", "children": [
            { "type": "image", "src": {{JsonSerializer.Serialize(source)}},
              "at": { "x": 0, "y": 0 }, "size": { "w": 100, "h": 100 } }
          ] }
        ] }
        """;
}
