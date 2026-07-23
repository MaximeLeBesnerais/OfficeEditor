using DocumentFormat.OpenXml.Packaging;
using OfficeEditor.Core.Services;
using PptxEditor.Core.Converters;
using Xunit;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Smoke tests for every reference PPTX in examples/REF/PPTX.
/// Ensures each file can be opened, converted to a Typst model, and rendered
/// to non-empty Typst source. The northwind-demo deck is also compiled to PDF
/// to verify end-to-end generation. (The original third-party reference decks
/// were removed for licensing; replacements pending.)
/// </summary>
public sealed class PptxReferenceSmokeTests : IDisposable
{
    private readonly string _referenceDirectory;
    private readonly string _tempDirectory;

    private static readonly string[] OriginalReferenceFileNames =
    [
        "northwind-demo.pptx"
    ];

    public PptxReferenceSmokeTests()
    {
        _referenceDirectory = ResolveReferenceDirectory();
        _tempDirectory = Path.Combine(Path.GetTempPath(), "PptxReferenceSmokeTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDirectory);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDirectory))
            {
                Directory.Delete(_tempDirectory, true);
            }
        }
        catch
        {
            // Best-effort cleanup.
        }
    }

    public static IEnumerable<object[]> ReferencePptxFiles()
    {
        var referenceDirectory = ResolveReferenceDirectory();

        if (!Directory.Exists(referenceDirectory))
        {
            throw new DirectoryNotFoundException($"Reference PPTX directory not found: {referenceDirectory}");
        }

        foreach (var filePath in Directory.GetFiles(referenceDirectory, "*.pptx").OrderBy(Path.GetFileName, StringComparer.OrdinalIgnoreCase))
        {
            yield return new object[] { Path.GetFileName(filePath) };
        }
    }

    [Theory]
    [MemberData(nameof(ReferencePptxFiles))]
    public void GenerateTypstSource_ReferencePptx_ProducesNonEmptySourceWithExpectedMarkers(string fileName)
    {
        var path = Path.Combine(_referenceDirectory, fileName);
        Assert.True(File.Exists(path), $"Reference PPTX file not found: {path}");

        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        Assert.True(presentation.Slides.Count > 0, $"Expected at least one slide in {fileName}, got {presentation.Slides.Count}.");

        var source = converter.GenerateTypstSource(presentation);

        Assert.False(string.IsNullOrWhiteSpace(source), $"Generated Typst source is empty for {fileName}.");
        Assert.Contains("#set page", source);
        Assert.Contains("#place", source);
    }

    [Fact]
    public void CompileToPdf_OriginalReferencePptxFiles_Succeeds()
    {
        foreach (var fileName in OriginalReferenceFileNames)
        {
            var path = Path.Combine(_referenceDirectory, fileName);
            Assert.True(File.Exists(path), $"Reference PPTX file not found: {path}");

            CompileReferenceToPdfAndAssert(path, fileName);
        }
    }

    private static void CompileReferenceToPdfAndAssert(string path, string fileName)
    {
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var presentation = converter.Convert();
        var source = converter.GenerateTypstSource(presentation);

        using var compiler = new TypstCompilerService();
        var result = compiler.Compile(source, new CompileOptions
        {
            Format = OutputFormat.Pdf,
            WorkingDirectory = presentation.TempDirectory,
            FontDirectory = Path.Combine(presentation.TempDirectory, "fonts"),
            ProcessTimeout = TimeSpan.FromMinutes(3)
        });

        Assert.True(result.Success, $"PDF compilation failed for {fileName}: {result.ErrorMessage}");
        Assert.Single(result.Pages);
        Assert.True(result.Pages[0].Length > 0, $"PDF output is empty for {fileName}.");
    }

    private static string ResolveReferenceDirectory()
    {
        var baseDir = AppContext.BaseDirectory;
        var referenceDirectory = Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "..", "examples", "REF", "PPTX"));
        return referenceDirectory;
    }
}
