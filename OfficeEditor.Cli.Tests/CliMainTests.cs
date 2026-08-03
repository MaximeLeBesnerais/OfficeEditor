using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using PptxEditor.Core.Builders;
using XlsxEditor.Core.Builders;
using Xunit;

namespace OfficeEditor.Cli.Tests;

public sealed class CliMainTests : IDisposable
{
    private readonly string _tempDir;

    public CliMainTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "OfficeEditorCliTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    [Fact]
    public void NoArgs_ReturnsOne()
    {
        Assert.Equal(1, InvokeMain([]));
    }

    [Fact]
    public void Help_ReturnsZero()
    {
        Assert.Equal(0, InvokeMain(["help"]));
        Assert.Equal(0, InvokeMain(["--help"]));
    }

    [Fact]
    public void UnknownCommand_ReturnsOne()
    {
        Assert.Equal(1, InvokeMain(["nonexistent"]));
    }

    [Fact]
    public void Create_Docx_ProducesValidFile()
    {
        var path = TempPath("test.docx");
        Assert.Equal(0, InvokeMain(["create", path, "--text", "Hello OfficeEditor"]));

        Assert.True(File.Exists(path));
        Assert.True(new FileInfo(path).Length > 0);
    }

    [Fact]
    public void Create_Pptx_ProducesValidFile()
    {
        var path = TempPath("test.pptx");
        Assert.Equal(0, InvokeMain(["create", path, "--title", "Test Deck"]));

        Assert.True(File.Exists(path));
        Assert.True(new FileInfo(path).Length > 0);

        using var builder = PresentationBuilder.Open(path);
        Assert.Equal(1, builder.SlideCount);
    }

    [Fact]
    public void Create_Xlsx_ProducesValidFile()
    {
        var path = TempPath("test.xlsx");
        Assert.Equal(0, InvokeMain(["create", path, "--sheet", "Data"]));

        Assert.True(File.Exists(path));
        Assert.True(new FileInfo(path).Length > 0);
    }

    [Fact]
    public void Create_MissingOutputPath_ReturnsOne()
    {
        Assert.Equal(1, InvokeMain(["create"]));
    }

    [Fact]
    public void Create_UnsupportedFormat_ReturnsOne()
    {
        var path = TempPath("test.pdf");
        Assert.Equal(1, InvokeMain(["create", path]));
    }

    [Fact]
    public void Detect_OnExistingPptx_FindsNoVariablesOnEmptyDeck()
    {
        var pptxPath = CreateMinimalPptx();
        Assert.Equal(0, InvokeMain(["detect", pptxPath]));
    }

    [Fact]
    public void Detect_WithNoArgs_ReturnsOne()
    {
        Assert.Equal(1, InvokeMain(["detect"]));
    }

    [Fact]
    public void Generate_WithValidJson_ProducesPptx()
    {
        var jsonPath = TempPath("deck.json");
        var genJson = ValidGenerationDocument();
        File.WriteAllText(jsonPath, genJson);

        var outputPath = TempPath("out.pptx");
        Assert.Equal(0, InvokeMain(["generate", jsonPath, "--output", outputPath]));

        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    [Fact]
    public void Generate_WithMissingInputFile_ReturnsOne()
    {
        var path = TempPath("nonexistent.json");
        Assert.Equal(1, InvokeMain(["generate", path]));
    }

    [Fact]
    public void Generate_WithNoArgs_ReturnsOne()
    {
        Assert.Equal(1, InvokeMain(["generate"]));
    }

    [Fact]
    public void Generate_MissingRequiredToken_ReturnsOne()
    {
        var jsonPath = TempPath("missing.json");
        var genJson = """
            {
              "version": "2.0",
              "design": {
                "palette": { "paper": "#FFFFFF", "ink": "#111111" },
                "fonts": { "display": "Aptos Display", "body": "Aptos" },
                "metrics": { "marginPt": 48, "gutterPt": 18, "titleSizePt": 30, "bodySizePt": 14 }
              },
              "slides": [
                {
                  "type": "cover",
                  "content": { "title": "Test" }
                }
              ]
            }
            """;
        File.WriteAllText(jsonPath, genJson);
        var outputPath = TempPath("out.pptx");
        var code = InvokeMain(["generate", jsonPath, "--output", outputPath]);
        Assert.Equal(1, code);
    }

    [Fact]
    public void Merge_WithValidData_ProducesFile()
    {
        var templatePath = CreateMinimalPptx();
        var dataPath = TempPath("data.json");
        File.WriteAllText(dataPath, """{"company":"TestCorp"}""");
        var outputPath = TempPath("merged.pptx");

        Assert.Equal(0, InvokeMain(["merge", templatePath, dataPath, outputPath]));

        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    [Fact]
    public void Merge_WithInvalidJson_ReturnsOne()
    {
        var templatePath = CreateMinimalPptx();
        var dataPath = TempPath("bad.json");
        File.WriteAllText(dataPath, "not json");
        var outputPath = TempPath("merged.pptx");

        Assert.Equal(1, InvokeMain(["merge", templatePath, dataPath, outputPath]));
    }

    [Fact]
    public void Merge_WithNoArgs_ReturnsOne()
    {
        Assert.Equal(1, InvokeMain(["merge"]));
    }

    [Fact]
    public void Edit_ReturnsOne_NotYetImplemented()
    {
        var inputPath = CreateMinimalPptx();
        var instrPath = TempPath("instructions.json");
        File.WriteAllText(instrPath, "{}");

        Assert.Equal(1, InvokeMain(["edit", inputPath, "--instructions", instrPath]));
    }

    [Fact]
    public void Edit_WithNoArgs_ReturnsOne()
    {
        Assert.Equal(1, InvokeMain(["edit"]));
    }

    [Fact]
    public void Generate_WithArchetypeSlide_ProducesValidPptx()
    {
        var jsonPath = TempPath("archetype.json");
        var genJson = """
            {
              "version": "2.0",
              "design": {
                "palette": { "primary": "#0B3D91", "accent": "#1E7BC6", "ink": "#1A1A1A", "paper": "#FFFFFF", "muted": "#8A94A6" }
              },
              "slides": [
                {
                  "type": "cover",
                  "content": {
                    "title": "Test",
                    "subtitle": "Generated"
                  }
                }
              ]
            }
            """;
        File.WriteAllText(jsonPath, genJson);

        var outputPath = TempPath("out.pptx");
        var code = InvokeMain(["generate", jsonPath, "--output", outputPath]);

        Assert.Equal(0, code);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Generate_WithValidXlsxJson_ProducesXlsx()
    {
        var jsonPath = TempPath("workbook.json");
        File.WriteAllText(jsonPath, """
        {
          "version": "1.0",
          "worksheets": [
            { "name": "Sheet1", "headers": ["A", "B"], "rows": [["1", "2"]] }
          ]
        }
        """);

        var outputPath = TempPath("out.xlsx");
        var code = InvokeMain(["generate", jsonPath, "--output", outputPath]);

        Assert.Equal(0, code);
        Assert.True(File.Exists(outputPath));
        using var reader = WorkbookBuilder.Open(outputPath);
        var ws = reader.GetWorksheet("Sheet1");
        Assert.Equal("A", ws.GetCellValue("A1"));
        Assert.Equal("2", ws.GetCellValue("B2"));
    }

    [Fact]
    public void Generate_WithValidDocxJson_ProducesDocx()
    {
        var jsonPath = TempPath("report.json");
        File.WriteAllText(jsonPath, """
        {
          "version": "1.0",
          "design": { "theme": "editorial" },
          "sections": [
            { "blocks": [ { "type": "paragraph", "text": "Hello DOCX" } ] }
          ]
        }
        """);

        var outputPath = TempPath("out.docx");
        var code = InvokeMain(["generate", jsonPath, "--output", outputPath]);

        Assert.Equal(0, code);
        Assert.True(File.Exists(outputPath));
        using var doc = WordprocessingDocument.Open(outputPath, false);
        Assert.Contains("Hello DOCX", doc.MainDocumentPart!.Document!.InnerText);
    }

    [Fact]
    public void Generate_DocxWithTheme_ProducesDocx()
    {
        var jsonPath = TempPath("report.json");
        File.WriteAllText(jsonPath, """
        {
          "version": "1.0",
          "sections": [
            { "blocks": [ { "type": "paragraph", "text": "Corporate theme" } ] }
          ]
        }
        """);

        var outputPath = TempPath("themed.docx");
        var code = InvokeMain(["generate", jsonPath, "--output", outputPath, "--theme", "corporate"]);

        Assert.Equal(0, code);
        Assert.True(File.Exists(outputPath));
        using var doc = WordprocessingDocument.Open(outputPath, false);
        Assert.Contains("Corporate theme", doc.MainDocumentPart!.Document!.InnerText);
    }

    [Fact]
    public void Generate_ThemeWithPptx_ReturnsOne()
    {
        var jsonPath = TempPath("deck.json");
        File.WriteAllText(jsonPath, ValidGenerationDocument());
        var outputPath = TempPath("out.pptx");

        var code = InvokeMain(["generate", jsonPath, "--output", outputPath, "--theme", "editorial"]);

        Assert.Equal(1, code);
        Assert.False(File.Exists(outputPath));
    }

    [Fact]
    public void Generate_UnknownTheme_ReturnsOne()
    {
        var jsonPath = TempPath("report.json");
        File.WriteAllText(jsonPath, """
        {
          "version": "1.0",
          "sections": [
            { "blocks": [ { "type": "paragraph", "text": "Hello" } ] }
          ]
        }
        """);
        var outputPath = TempPath("out.docx");

        var code = InvokeMain(["generate", jsonPath, "--output", outputPath, "--theme", "nonexistent"]);

        Assert.Equal(1, code);
        Assert.False(File.Exists(outputPath));
    }

    [Fact]
    public void Generate_UnknownFlag_ReturnsOne()
    {
        var jsonPath = TempPath("deck.json");
        File.WriteAllText(jsonPath, ValidGenerationDocument());
        var outputPath = TempPath("out.pptx");

        var code = InvokeMain(["generate", jsonPath, "--output", outputPath, "--bogus"]);

        Assert.Equal(1, code);
    }

    [Fact]
    public void Generate_UnsupportedExtension_ReturnsOne()
    {
        var jsonPath = TempPath("deck.json");
        File.WriteAllText(jsonPath, ValidGenerationDocument());
        var outputPath = TempPath("out.pdf");

        var code = InvokeMain(["generate", jsonPath, "--output", outputPath]);

        Assert.Equal(1, code);
        Assert.False(File.Exists(outputPath));
    }

    [Fact]
    public void Generate_InvalidXlsxJson_ReturnsOneAndNoOutput()
    {
        var jsonPath = TempPath("bad.json");
        File.WriteAllText(jsonPath, """{ "version": "9.9", "worksheets": [{"name": "S"}] }""");
        var outputPath = TempPath("out.xlsx");

        var code = InvokeMain(["generate", jsonPath, "--output", outputPath]);

        Assert.Equal(1, code);
        Assert.False(File.Exists(outputPath));
    }

    [Fact]
    public void Generate_InvalidDocxJson_ReturnsOneAndNoOutput()
    {
        var jsonPath = TempPath("bad.json");
        File.WriteAllText(jsonPath, """{ "version": "1.0", "sections": [] }""");
        var outputPath = TempPath("out.docx");

        var code = InvokeMain(["generate", jsonPath, "--output", outputPath]);

        Assert.Equal(1, code);
        Assert.False(File.Exists(outputPath));
    }

    [Fact]
    public void Generate_ThemeWithoutValue_ReturnsOne()
    {
        var jsonPath = TempPath("report.json");
        File.WriteAllText(jsonPath, """
        {
          "version": "1.0",
          "sections": [
            { "blocks": [ { "type": "paragraph", "text": "Hello" } ] }
          ]
        }
        """);
        var outputPath = TempPath("out.docx");

        var code = InvokeMain(["generate", jsonPath, "--output", outputPath, "--theme"]);

        Assert.Equal(1, code);
    }

    private static int InvokeMain(string[] args)
    {
        var cliAssembly = AppDomain.CurrentDomain.GetAssemblies()
            .FirstOrDefault(a => a.GetName().Name == "OfficeEditor.Cli");
        if (cliAssembly == null)
        {
            cliAssembly = Assembly.Load("OfficeEditor.Cli");
        }

        var programType = cliAssembly.GetType("OfficeEditor.Cli.Program")
            ?? throw new InvalidOperationException("Program type not found in OfficeEditor.Cli assembly.");

        var method = programType.GetMethod("Main",
            BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.Public,
            [typeof(string[])])
            ?? throw new InvalidOperationException("Main method not found.");

        var result = method.Invoke(null, [args]);
        return (int)(result ?? throw new InvalidOperationException("Main signature changed; update InvokeMain"));
    }

    private string TempPath(string fileName)
        => Path.Combine(_tempDir, fileName);

    private string CreateMinimalPptx()
    {
        var path = TempPath("minimal.pptx");
        using var builder = PresentationBuilder.Create(path);
        builder.AddSlide();
        builder.CurrentSlide.AddTitle("Test Slide");
        builder.Save();
        return path;
    }

    private static string ValidGenerationDocument()
        => """
        {
          "version": "2.0",
          "design": {
            "palette": { "primary": "#0B3D91", "accent": "#1E7BC6", "ink": "#1A1A1A", "paper": "#FFFFFF", "muted": "#8A94A6" },
            "fonts": { "display": "Aptos Display", "body": "Aptos" },
            "metrics": { "marginPt": 43, "gutterPt": 18, "titleSizePt": 30, "bodySizePt": 14 }
          },
          "slides": [
            {
              "type": "container",
              "fill": "paper",
              "layout": { "mode": "row", "gap": 18 },
              "overflow": "clip",
              "padding": 43,
              "children": [
                {
                  "type": "card",
                  "size": { "grow": 1, "aspect": "4:3" },
                  "content": { "title": "Hello", "subtitle": "CLI test" }
                }
              ]
            }
          ]
        }
        """;

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDir))
                Directory.Delete(_tempDir, true);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }
}
