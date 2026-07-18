using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using OfficeEditor.Core.Services;
using PptxEditor.Core.Converters;
using Xunit;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Tests for the single-slide conversion API (<see cref="PptxToTypstConverter.ConvertSingleSlide"/>).
/// The emitted document must be self-contained (global <c>#set text</c> header plus the
/// slide's own <c>#set page</c> block), compile to exactly one page, and match the
/// corresponding page of the whole-deck emission at the source level.
/// </summary>
public sealed class PptxToTypstConverterSingleSlideTests : IDisposable
{
    private readonly string _tempDir;

    public PptxToTypstConverterSingleSlideTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), nameof(PptxToTypstConverterSingleSlideTests), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        SlideOpsTestHelpers.BestEffortDelete(_tempDir);
    }

    [Fact]
    public void ConvertSingleSlide_OutOfRange_Throws()
    {
        var path = SlideOpsTestHelpers.CreateDeckWithTitledSlides(_tempDir, "one", "two", "three");
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        Assert.Throws<ArgumentOutOfRangeException>(() => converter.ConvertSingleSlide(-1));
        Assert.Throws<ArgumentOutOfRangeException>(() => converter.ConvertSingleSlide(3));
    }

    [Fact]
    public void ConvertSingleSlide_ReturnsExactlyOneSlideWithOneBasedSlideIndex()
    {
        var path = SlideOpsTestHelpers.CreateDeckWithTitledSlides(_tempDir, "one", "two", "three");
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.ConvertSingleSlide(1);

        var slide = Assert.Single(presentation.Slides);
        Assert.Equal(2, slide.SlideIndex); // TypstSlide.SlideIndex stays 1-based
        Assert.Contains(slide.Elements, e => e.Text?.Content.Contains("two") == true);
        Assert.DoesNotContain(slide.Elements, e => e.Text?.Content.Contains("one") == true);
        Assert.DoesNotContain(slide.Elements, e => e.Text?.Content.Contains("three") == true);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void ConvertSingleSlide_SourceMatchesWholeDeckSlideBlock(int slideIndex)
    {
        var path = SlideOpsTestHelpers.CreateDeckWithTitledSlides(_tempDir, "one", "two", "three");

        string wholeDeckSource;
        using (var document = PresentationDocument.Open(path, false))
        using (var converter = new PptxToTypstConverter(document))
        {
            wholeDeckSource = converter.GenerateTypstSource(converter.Convert());
        }

        string singleSlideSource;
        using (var document = PresentationDocument.Open(path, false))
        using (var converter = new PptxToTypstConverter(document))
        {
            singleSlideSource = converter.GenerateTypstSource(converter.ConvertSingleSlide(slideIndex));
        }

        // The global header (#set text ...) must be identical in both emissions.
        var wholeHeader = HeaderOf(wholeDeckSource);
        Assert.Equal(wholeHeader, HeaderOf(singleSlideSource));

        // The slide block (#set page ... content) of the single-slide emission must match
        // the corresponding block of the whole-deck emission verbatim.
        var wholeBlock = ExtractSlideBlock(wholeDeckSource, slideIndex);
        var singleBlock = ExtractSlideBlock(singleSlideSource, 0);
        Assert.Equal(wholeBlock, singleBlock);
    }

    [Fact]
    public void ConvertSingleSlide_EmitsCompilableOnePageDocument()
    {
        var referencePath = Path.Combine(ResolveReferenceDirectory(), "Presentation1.pptx");
        Assert.True(File.Exists(referencePath), $"Reference PPTX file not found: {referencePath}");

        using var document = PresentationDocument.Open(referencePath, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.ConvertSingleSlide(0);
        var source = converter.GenerateTypstSource(presentation);

        using var compiler = new TypstCompilerService();
        var result = compiler.Compile(source, new CompileOptions
        {
            Format = OutputFormat.Pdf,
            WorkingDirectory = presentation.TempDirectory,
            FontDirectory = Path.Combine(presentation.TempDirectory, "fonts"),
            ProcessTimeout = TimeSpan.FromMinutes(3)
        });

        Assert.True(result.Success, $"Single-slide PDF compilation failed: {result.ErrorMessage}");
        Assert.Single(result.Pages);
        Assert.True(result.Pages[0].Length > 0, "Single-slide PDF output is empty.");
    }

    /// <summary>Returns the source up to (excluding) the first slide's page setup.</summary>
    private static string HeaderOf(string source)
    {
        var index = source.IndexOf("#set page(width:", StringComparison.Ordinal);
        Assert.True(index > 0, "Generated source contains no '#set page(width:' marker.");
        return source[..index];
    }

    /// <summary>
    /// Extracts the Nth slide block: from its <c>#set page(width:</c> up to the next
    /// slide's <c>#set page(width:</c> or end of source. Whole-deck non-first slides carry
    /// a <c>#pagebreak()</c> between page setup and content which the single-slide emission
    /// (always "first") omits, so it is stripped. Extracted asset file names are normalized
    /// (image_N) because they are numbered per conversion run.
    /// </summary>
    private static string ExtractSlideBlock(string source, int slideIndex)
    {
        var matches = Regex.Matches(source, "#set page\\(width:", RegexOptions.None);
        Assert.True(matches.Count > slideIndex,
            $"Expected at least {slideIndex + 1} slide block(s), found {matches.Count}.");

        var start = matches[slideIndex].Index;
        var end = matches.Count > slideIndex + 1 ? matches[slideIndex + 1].Index : source.Length;
        var block = source[start..end];

        block = block.Replace("#pagebreak()", "");
        block = Regex.Replace(block, @"image_\d+", "image_N");
        block = Regex.Replace(block, @"\n{3,}", "\n\n");
        return block.Trim();
    }

    private static string ResolveReferenceDirectory()
    {
        var baseDir = AppContext.BaseDirectory;
        return Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "..", "examples", "REF", "PPTX"));
    }
}
