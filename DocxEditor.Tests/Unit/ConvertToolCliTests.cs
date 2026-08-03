using System.Reflection;
using System.Text;
using DocxEditor.Core.Builders;
using PptxEditor.Core.Builders;
using XlsxEditor.Core.Builders;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// CLI parity tests for the convert tools (<c>tools/convert-pptx</c>, <c>tools/convert-docx</c>,
/// <c>tools/convert-xlsx</c>): every tool accepts <c>--format pdf|png|svg|typ</c> through the
/// shared render facade (<see cref="OfficeEditor.Core.Rendering.DocumentRenderer"/>) and writes
/// the expected output kind — a single file for <c>pdf</c>/<c>typ</c>, a directory of
/// <c>page-NNN.ext</c> for <c>png</c>/<c>svg</c>. Compiles run through the repository's own Typst
/// pipeline on tiny in-memory fixtures (same convention as
/// <see cref="DocumentRendererFacadeTests"/>) — no office suite is involved.
/// </summary>
public abstract class ConvertToolTestBase : IDisposable
{
    private readonly string _tempDirectory;

    protected ConvertToolTestBase()
    {
        _tempDirectory = Path.Combine(Path.GetTempPath(), "OfficeEditorConvertToolTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDirectory);
    }

    protected string WriteInput(string extension, byte[] content)
    {
        var path = Path.Combine(_tempDirectory, $"input-{Guid.NewGuid():N}{extension}");
        File.WriteAllBytes(path, content);
        return path;
    }

    protected string OutputPath(string fileName) => Path.Combine(_tempDirectory, fileName);

    protected static byte[] CreatePptxBytes(int slideCount = 2)
    {
        using var builder = PresentationBuilder.Create();
        for (var i = 0; i < slideCount; i++)
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTitle($"Slide {i + 1}");
        }

        return builder.SaveToBytes();
    }

    protected static byte[] CreateDocxBytes()
    {
        using var builder = DocumentBuilder.Create();
        builder.AddParagraph("Hello convert tool", "Normal");
        return builder.SaveToBytes();
    }

    protected static byte[] CreateXlsxBytes()
    {
        using var builder = WorkbookBuilder.Create();
        var sheet = builder.AddWorksheet("Data");
        sheet.AddHeaderRow(["A", "B"], rowIndex: 1);
        sheet.AddDataRow(["1", "2"], rowIndex: 2);
        return builder.SaveToBytes();
    }

    /// <summary>
    /// Invokes a convert tool's generated top-level <c>Program.&lt;Main&gt;$</c> from its own
    /// assembly (top-level statements compile the entry point under the <c>&lt;Main&gt;$</c> name)
    /// and returns its exit code.
    /// </summary>
    protected static int InvokeTool(string toolName, string[] args)
    {
        var assembly = AppDomain.CurrentDomain.GetAssemblies()
            .FirstOrDefault(assembly => assembly.GetName().Name == toolName)
            ?? Assembly.Load(toolName);

        var programType = assembly.GetType("Program")
            ?? throw new InvalidOperationException($"Program type not found in {toolName} assembly.");

        var method = programType.GetMethod("<Main>$",
            BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.Public)
            ?? throw new InvalidOperationException("Main method not found.");

        var parameters = method.GetParameters();
        if (parameters.Length != 1 || parameters[0].ParameterType != typeof(string[]))
        {
            throw new InvalidOperationException($"Main signature changed: {method}");
        }

        var result = method.Invoke(null, [args]);
        return (int)(result ?? throw new InvalidOperationException("Main signature changed; update InvokeTool"));
    }

    protected static void AssertPdfFile(string path)
    {
        Assert.True(File.Exists(path), $"PDF file missing: {path}");
        var bytes = File.ReadAllBytes(path);
        Assert.True(bytes.Length > 8, "PDF file is suspiciously small.");
        Assert.Equal(0x25, bytes[0]); // '%'
        Assert.Equal((byte)'P', bytes[1]);
        Assert.Equal((byte)'D', bytes[2]);
        Assert.Equal((byte)'F', bytes[3]);
    }

    protected static void AssertPngFile(string path)
    {
        Assert.True(File.Exists(path), $"PNG file missing: {path}");
        var bytes = File.ReadAllBytes(path);
        Assert.True(bytes.Length > 8, "PNG page is suspiciously small.");
        Assert.Equal(0x89, bytes[0]);
        Assert.Equal((byte)'P', bytes[1]);
        Assert.Equal((byte)'N', bytes[2]);
        Assert.Equal((byte)'G', bytes[3]);
    }

    protected static void AssertSvgFile(string path)
    {
        Assert.True(File.Exists(path), $"SVG file missing: {path}");
        var head = Encoding.ASCII.GetString(File.ReadAllBytes(path), 0, 64);
        Assert.Contains("<svg", head);
    }

    protected static void AssertTypstSourceFile(string path)
    {
        Assert.True(File.Exists(path), $"Typst source missing: {path}");
        var source = File.ReadAllText(path);
        Assert.False(string.IsNullOrWhiteSpace(source));
        Assert.Contains("#set", source);
    }

    protected static void AssertPageDirectory(string directory, string extension, int expectedMinPages = 1)
    {
        Assert.True(Directory.Exists(directory), $"Output directory missing: {directory}");
        var pages = Directory.GetFiles(directory, $"page-*.{extension}")
            .OrderBy(Path.GetFileName, StringComparer.Ordinal)
            .ToList();
        Assert.True(pages.Count >= expectedMinPages, $"Expected at least {expectedMinPages} page(s), found {pages.Count}.");
        Assert.Equal($"page-001.{extension}", Path.GetFileName(pages[0]));
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDirectory))
            {
                Directory.Delete(_tempDirectory, recursive: true);
            }
        }
        catch
        {
            // Best-effort cleanup of the test scratch directory.
        }
    }
}

public sealed class ConvertPptxToolTests : ConvertToolTestBase
{
    private const string Tool = "convert-pptx";

    [Fact]
    public void Pdf_ProducesSinglePdfFile()
    {
        var output = OutputPath("deck.pdf");
        Assert.Equal(0, InvokeTool(Tool, [WriteInput(".pptx", CreatePptxBytes(1)), output, "--format", "pdf"]));
        AssertPdfFile(output);
    }

    [Fact]
    public void Png_ProducesDirectoryOfPngPages()
    {
        var output = OutputPath("deck-png");
        Assert.Equal(0, InvokeTool(Tool, [WriteInput(".pptx", CreatePptxBytes(2)), output, "--format", "png", "--ppi", "72"]));
        AssertPageDirectory(output, "png", expectedMinPages: 2);
        AssertPngFile(Path.Combine(output, "page-001.png"));
        AssertPngFile(Path.Combine(output, "page-002.png"));
    }

    [Fact]
    public void Svg_ProducesDirectoryOfSvgPages()
    {
        var output = OutputPath("deck-svg");
        Assert.Equal(0, InvokeTool(Tool, [WriteInput(".pptx", CreatePptxBytes(2)), output, "--format", "svg"]));
        AssertPageDirectory(output, "svg", expectedMinPages: 2);
        AssertSvgFile(Path.Combine(output, "page-001.svg"));
        AssertSvgFile(Path.Combine(output, "page-002.svg"));
    }

    [Fact]
    public void Typ_ProducesTypstSourceFile()
    {
        var output = OutputPath("deck.typ");
        Assert.Equal(0, InvokeTool(Tool, [WriteInput(".pptx", CreatePptxBytes(1)), output, "--format", "typ"]));
        AssertTypstSourceFile(output);
    }

    [Fact]
    public void NoArgs_ReturnsOne()
    {
        Assert.Equal(1, InvokeTool(Tool, []));
    }

    [Fact]
    public void UnknownFormat_ReturnsOne()
    {
        Assert.Equal(1, InvokeTool(Tool, [WriteInput(".pptx", CreatePptxBytes(1)), OutputPath("out.pdf"), "--format", "bogus"]));
    }

    [Fact]
    public void MissingInputFile_ReturnsOne()
    {
        Assert.Equal(1, InvokeTool(Tool, [OutputPath("nope.pptx"), OutputPath("out.pdf")]));
    }

    [Fact]
    public void UnsupportedSourceType_ReturnsOne()
    {
        Assert.Equal(1, InvokeTool(Tool, [WriteInput(".txt", Encoding.UTF8.GetBytes("not a document")), OutputPath("out.pdf"), "--format", "pdf"]));
    }
}

public sealed class ConvertDocxToolTests : ConvertToolTestBase
{
    private const string Tool = "convert-docx";

    [Fact]
    public void Pdf_ProducesSinglePdfFile()
    {
        var output = OutputPath("doc.pdf");
        Assert.Equal(0, InvokeTool(Tool, [WriteInput(".docx", CreateDocxBytes()), output, "--format", "pdf"]));
        AssertPdfFile(output);
    }

    [Fact]
    public void Png_ProducesDirectoryOfNonEmptyPngPages()
    {
        // DOCX → PNG is the headline parity gap: assert the facade path really renders pages.
        var output = OutputPath("doc-png");
        Assert.Equal(0, InvokeTool(Tool, [WriteInput(".docx", CreateDocxBytes()), output, "--format", "png", "--ppi", "72"]));
        AssertPageDirectory(output, "png");
        AssertPngFile(Path.Combine(output, "page-001.png"));
    }

    [Fact]
    public void Svg_ProducesDirectoryOfNonEmptySvgPages()
    {
        var output = OutputPath("doc-svg");
        Assert.Equal(0, InvokeTool(Tool, [WriteInput(".docx", CreateDocxBytes()), output, "--format", "svg"]));
        AssertPageDirectory(output, "svg");
        AssertSvgFile(Path.Combine(output, "page-001.svg"));
    }

    [Fact]
    public void Typ_ProducesTypstSourceFile()
    {
        var output = OutputPath("doc.typ");
        Assert.Equal(0, InvokeTool(Tool, [WriteInput(".docx", CreateDocxBytes()), output, "--format", "typ"]));
        AssertTypstSourceFile(output);
    }

    [Fact]
    public void NoArgs_ReturnsOne()
    {
        Assert.Equal(1, InvokeTool(Tool, []));
    }

    [Fact]
    public void UnknownFormat_ReturnsOne()
    {
        Assert.Equal(1, InvokeTool(Tool, [WriteInput(".docx", CreateDocxBytes()), OutputPath("out.pdf"), "--format", "bogus"]));
    }

    [Fact]
    public void MissingInputFile_ReturnsOne()
    {
        Assert.Equal(1, InvokeTool(Tool, [OutputPath("nope.docx"), OutputPath("out.pdf")]));
    }
}

public sealed class ConvertXlsxToolTests : ConvertToolTestBase
{
    private const string Tool = "convert-xlsx";

    [Fact]
    public void Pdf_ProducesSinglePdfFile()
    {
        var output = OutputPath("book.pdf");
        Assert.Equal(0, InvokeTool(Tool, [WriteInput(".xlsx", CreateXlsxBytes()), output, "--format", "pdf"]));
        AssertPdfFile(output);
    }

    [Fact]
    public void Png_ProducesDirectoryOfPngPages()
    {
        var output = OutputPath("book-png");
        Assert.Equal(0, InvokeTool(Tool, [WriteInput(".xlsx", CreateXlsxBytes()), output, "--format", "png", "--ppi", "72"]));
        AssertPageDirectory(output, "png");
        AssertPngFile(Path.Combine(output, "page-001.png"));
    }

    [Fact]
    public void Svg_ProducesDirectoryOfSvgPages()
    {
        var output = OutputPath("book-svg");
        Assert.Equal(0, InvokeTool(Tool, [WriteInput(".xlsx", CreateXlsxBytes()), output, "--format", "svg"]));
        AssertPageDirectory(output, "svg");
        AssertSvgFile(Path.Combine(output, "page-001.svg"));
    }

    [Fact]
    public void Typ_ProducesTypstSourceFile()
    {
        var output = OutputPath("book.typ");
        Assert.Equal(0, InvokeTool(Tool, [WriteInput(".xlsx", CreateXlsxBytes()), output, "--format", "typ"]));
        AssertTypstSourceFile(output);
    }

    [Fact]
    public void Json_ProducesInstructionSetFile()
    {
        var output = OutputPath("book.json");
        Assert.Equal(0, InvokeTool(Tool, [WriteInput(".xlsx", CreateXlsxBytes()), output, "--format", "json"]));
        Assert.True(File.Exists(output), $"JSON file missing: {output}");
        var json = File.ReadAllText(output);
        Assert.Contains("\"worksheets\"", json);
    }

    [Fact]
    public void JsonSource_ToPdf_ProducesSinglePdfFile()
    {
        var json = WriteInput(".json", Encoding.UTF8.GetBytes("""
            {
              "version": "1.0",
              "worksheets": [
                { "name": "Sheet1", "headers": ["A", "B"], "rows": [["1", "2"]] }
              ]
            }
            """));
        var output = OutputPath("from-json.pdf");
        Assert.Equal(0, InvokeTool(Tool, [json, output, "--format", "pdf"]));
        AssertPdfFile(output);
    }

    [Fact]
    public void NoArgs_ReturnsOne()
    {
        Assert.Equal(1, InvokeTool(Tool, []));
    }

    [Fact]
    public void UnknownFormat_ReturnsOne()
    {
        Assert.Equal(1, InvokeTool(Tool, [WriteInput(".xlsx", CreateXlsxBytes()), OutputPath("out.pdf"), "--format", "bogus"]));
    }

    [Fact]
    public void MissingInputFile_ReturnsOne()
    {
        Assert.Equal(1, InvokeTool(Tool, [OutputPath("nope.xlsx"), OutputPath("out.pdf")]));
    }

    [Fact]
    public void UnknownArgument_ReturnsOne()
    {
        Assert.Equal(1, InvokeTool(Tool, [WriteInput(".xlsx", CreateXlsxBytes()), OutputPath("out.pdf"), "--bogus"]));
    }
}
