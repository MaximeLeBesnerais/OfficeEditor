using OfficeEditor.Testing;
using PptxEditor.Core.Generation;
using System.Reflection;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using PptxEditor.Core.Builders;
using Spectre.Console;
using XlsxEditor.Core.Builders;
using Xunit;

namespace OfficeEditor.Cli.Tests;

public sealed class CliMainTests : IDisposable
{
    /// <summary>
    /// PDF/PNG/SVG output compiles through the repository's Typst pipeline, so those tests are
    /// opt-in with the same env var the rest of the suite uses (CI does not always have the
    /// native bridge). Dispatch/validation behaviour is covered by ungated tests below.
    /// </summary>
    private const string EnableRenderEnvVar = "OE_RUN_TYPST_COMPILE_TESTS";

    private readonly string _tempDir;

    public CliMainTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "OfficeEditorCliTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    [Fact]
    public void Ids_DryRunAndBackup_PreserveOriginalAndRejectExistingBackup()
    {
        var source = TempPath("source.json");
        var original = ValidGenerationDocument();
        File.WriteAllText(source, original);
        Assert.Equal(0, InvokeMain(["ids", source, "--dry-run", "--backup"]));
        Assert.Equal(original, File.ReadAllText(source));
        Assert.False(File.Exists(source + ".bak"));
        Assert.Equal(0, InvokeMain(["ids", source, "--backup"]));
        Assert.Equal(original, File.ReadAllText(source + ".bak"));
        Assert.Contains("\"id\"", File.ReadAllText(source));
        File.WriteAllText(source, original);
        Assert.Equal(1, InvokeMain(["ids", source, "--backup"]));
        Assert.Equal(original, File.ReadAllText(source));
    }

    [Fact]
    public void EditSource_GeneratesEditedDeck_AndNeverOverwritesExistingOutput()
    {
        var source = TempPath("source.json");
        var operations = TempPath("operations.json");
        var output = TempPath("edited.json");
        var original = ValidGenerationDocument();
        File.WriteAllText(source, original);
        File.WriteAllText(operations, """[{"type":"rename","id":"slide-0","newId":"cover"}]""");
        Assert.Equal(0, InvokeMain(["edit", source, "--instructions", operations, "--output", output]));
        Assert.Equal(original, File.ReadAllText(source));
        Assert.Contains("cover", File.ReadAllText(output));
        Assert.Equal(0, InvokeMain(["inspect", output]));
        Assert.Equal(0, InvokeMain(["generate", output, "--output", TempPath("edited.pptx")]));
        var saved = File.ReadAllText(output);
        Assert.Equal(1, InvokeMain(["edit", source, "--instructions", operations, "--output", output]));
        Assert.Equal(saved, File.ReadAllText(output));
        File.WriteAllText(operations, """[{"type":"rename","id":"absent","newId":"cover"}]""");
        var failed = TempPath("failed.json");
        Assert.Equal(1, InvokeMain(["edit", source, "--instructions", operations, "--output", failed]));
        Assert.False(File.Exists(failed));
    }

    [Fact]
    public void IdsCommand_NormalizesGenerationOnly_AndIsIdempotent()
    {
        var source = TempPath("ids.json");
        var config = TempPath("settings.json");
        File.WriteAllText(source, ValidGenerationDocument());
        File.WriteAllText(config, "{\"theme\":\"dark\"}");
        Assert.Equal(0, InvokeMain(["ids", _tempDir, "--recursive"]));
        var normalized = File.ReadAllText(source);
        Assert.Contains("\"id\"", normalized);
        Assert.Equal("{\"theme\":\"dark\"}", File.ReadAllText(config));
        Assert.Equal(0, InvokeMain(["ids", _tempDir, "--recursive"]));
        Assert.Equal(normalized, File.ReadAllText(source));
    }

    [Fact]
    public void Generate_PersistsIdsIntoSource()
    {
        var source = TempPath("source.json");
        File.WriteAllText(source, ValidGenerationDocument());
        Assert.Equal(0, InvokeMain(["generate", source, "--output", TempPath("source.pptx")]));
        var normalized = File.ReadAllText(source);
        Assert.Contains("\"id\"", normalized);
        Assert.Equal(0, InvokeMain(["generate", source, "--output", TempPath("source.pptx")]));
        Assert.Equal(normalized, File.ReadAllText(source));
    }

    [Fact]
    public void Generate_ContractDocument_MatchesCoreDelivery()
    {
        var expected = new PptxGenerator().Generate(GenerationContract.Json);
        Assert.True(expected.Success);
        var source = TempPath("contract.json");
        var output = TempPath("contract.pptx");
        File.WriteAllText(source, GenerationContract.Json);
        Assert.Equal(0, InvokeMain(["generate", source, "--output", output]));
        Assert.Equal(GenerationContract.Snapshot(expected.PptxBytes!), GenerationContract.Snapshot(File.ReadAllBytes(output)));
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
        var outputPath = TempPath("out.txt");

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

    // ─── Render output (.pdf/.png/.svg) dispatches by the source JSON's vocabulary ───

    [Fact]
    public void Generate_PptxJsonToPdf_RendersTheDeck()
    {
        if (!RenderTestsEnabled()) return;
        var jsonPath = TempPath("deck.json");
        File.WriteAllText(jsonPath, ValidGenerationDocument());
        var outputPath = TempPath("deck.pdf");

        Assert.Equal(0, InvokeMain(["generate", jsonPath, "--output", outputPath]));

        Assert.True(File.Exists(outputPath));
        Assert.Equal("%PDF", Magic(outputPath, 4));
    }

    [Fact]
    public void Generate_PptxJsonToPng_WritesOneImagePerSlide()
    {
        if (!RenderTestsEnabled()) return;
        var jsonPath = TempPath("deck.json");
        File.WriteAllText(jsonPath, ValidGenerationDocument());
        var outputDirectory = TempPath("deck.png");

        Assert.Equal(0, InvokeMain(["generate", jsonPath, "--output", outputDirectory]));

        Assert.True(Directory.Exists(outputDirectory));
        var page = Path.Combine(outputDirectory, "page-001.png");
        Assert.True(File.Exists(page));
        Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47 }, File.ReadAllBytes(page)[..4]);
    }

    [Fact]
    public void Generate_PptxJsonToSvg_WritesOneImagePerSlide()
    {
        if (!RenderTestsEnabled()) return;
        var jsonPath = TempPath("deck.json");
        File.WriteAllText(jsonPath, ValidGenerationDocument());
        var outputDirectory = TempPath("deck.svg");

        Assert.Equal(0, InvokeMain(["generate", jsonPath, "--output", outputDirectory]));

        var page = Path.Combine(outputDirectory, "page-001.svg");
        Assert.True(File.Exists(page));
        Assert.Contains("<svg", File.ReadAllText(page));
    }

    [Fact]
    public void Generate_DocxJsonToPdf_RendersTheReport()
    {
        if (!RenderTestsEnabled()) return;
        var jsonPath = TempPath("report.json");
        File.WriteAllText(jsonPath, ValidDocxGenerationDocument());
        var outputPath = TempPath("report.pdf");

        Assert.Equal(0, InvokeMain(["generate", jsonPath, "--output", outputPath]));

        Assert.True(File.Exists(outputPath));
        Assert.Equal("%PDF", Magic(outputPath, 4));
    }

    [Fact]
    public void Generate_XlsxJsonToPdf_RendersTheWorkbook()
    {
        if (!RenderTestsEnabled()) return;
        var jsonPath = TempPath("workbook.json");
        File.WriteAllText(jsonPath, ValidXlsxInstructionSet());
        var outputPath = TempPath("workbook.pdf");

        Assert.Equal(0, InvokeMain(["generate", jsonPath, "--output", outputPath]));

        Assert.True(File.Exists(outputPath));
        Assert.Equal("%PDF", Magic(outputPath, 4));
    }

    [Fact]
    public void Generate_XlsxJsonToPng_WritesOneImagePerPage()
    {
        if (!RenderTestsEnabled()) return;
        var jsonPath = TempPath("workbook.json");
        File.WriteAllText(jsonPath, ValidXlsxInstructionSet());
        var outputDirectory = TempPath("workbook.png");

        Assert.Equal(0, InvokeMain(["generate", jsonPath, "--output", outputDirectory]));

        Assert.True(File.Exists(Path.Combine(outputDirectory, "page-001.png")));
    }

    [Fact]
    public void Generate_DocxJsonToPng_WritesOneImagePerPage()
    {
        if (!RenderTestsEnabled()) return;
        var jsonPath = TempPath("report.json");
        File.WriteAllText(jsonPath, ValidDocxGenerationDocument());
        var outputDirectory = TempPath("report.png");

        Assert.Equal(0, InvokeMain(["generate", jsonPath, "--output", outputDirectory]));

        Assert.True(File.Exists(Path.Combine(outputDirectory, "page-001.png")));
    }

    [Fact]
    public void Generate_InvalidPptxJsonToPdf_ReportsPptxVocabulary_NotXlsx()
    {
        var jsonPath = TempPath("bad-deck.json");
        File.WriteAllText(jsonPath, """
        {
          "version": "2.0",
          "slides": [ { "type": "cover", "content": { "title": 42 } } ]
        }
        """);
        var outputPath = TempPath("bad.pdf");

        var (code, output) = InvokeMainCapturingOutput(["generate", jsonPath, "--output", outputPath]);

        Assert.Equal(1, code);
        Assert.False(File.Exists(outputPath));
        Assert.Contains("invalid PPTX generation JSON", output);
        Assert.Contains("--output <path.pptx>", output);
        Assert.DoesNotContain("XLSX", output);
    }

    [Fact]
    public void Generate_InvalidDocxJsonToPdf_ReportsDocxVocabulary()
    {
        var jsonPath = TempPath("bad-report.json");
        File.WriteAllText(jsonPath, """{ "version": "1.0", "sections": [] }""");
        var outputPath = TempPath("bad.pdf");

        var (code, output) = InvokeMainCapturingOutput(["generate", jsonPath, "--output", outputPath]);

        Assert.Equal(1, code);
        Assert.False(File.Exists(outputPath));
        Assert.Contains("invalid DOCX generation JSON", output);
        Assert.Contains("at least one section is required", output);
        Assert.DoesNotContain("XLSX", output);
    }

    [Fact]
    public void Generate_NonGenerationJsonToPdf_NamesTheSupportedVocabularies()
    {
        var jsonPath = TempPath("settings.json");
        File.WriteAllText(jsonPath, """{ "theme": "dark" }""");
        var outputPath = TempPath("settings.pdf");

        var (code, output) = InvokeMainCapturingOutput(["generate", jsonPath, "--output", outputPath]);

        Assert.Equal(1, code);
        Assert.False(File.Exists(outputPath));
        Assert.Contains("unrecognized generation JSON", output);
        Assert.Contains("PPTX", output);
        Assert.Contains("DOCX", output);
        Assert.Contains("XLSX", output);
    }

    [Fact]
    public void Generate_SvgOutputIsAccepted()
    {
        if (!RenderTestsEnabled()) return;
        var jsonPath = TempPath("workbook.json");
        File.WriteAllText(jsonPath, ValidXlsxInstructionSet());
        var outputDirectory = TempPath("workbook.svg");

        Assert.Equal(0, InvokeMain(["generate", jsonPath, "--output", outputDirectory]));

        Assert.False(File.Exists(outputDirectory));
        Assert.True(File.Exists(Path.Combine(outputDirectory, "page-001.svg")));
    }

    // ─── Help/usage text matches the parser ───

    [Fact]
    public void Help_DocumentsRenderFormatsAndGenerationFlags()
    {
        var (code, output) = InvokeMainCapturingOutput(["help"]);

        Assert.Equal(0, code);
        Assert.Contains("--instructions", output);
        Assert.Contains("--no-write-ids", output);
        Assert.Contains("--backup", output);
        Assert.Contains(".pdf", output);
        Assert.Contains(".png", output);
        Assert.Contains(".svg", output);
        Assert.Contains("page-001.png", output);
    }

    [Fact]
    public void Generate_OutputFlagWithoutValue_ReportsTheMissingValue()
    {
        var jsonPath = TempPath("deck.json");
        File.WriteAllText(jsonPath, ValidGenerationDocument());

        var (code, output) = InvokeMainCapturingOutput(["generate", jsonPath, "--output"]);

        Assert.Equal(1, code);
        Assert.Contains("--output requires a value", output);
        Assert.DoesNotContain("ending in .pptx, .docx, or .xlsx", output);
    }

    // ─── create --instructions generates real content per output format ───

    [Fact]
    public void Create_XlsxFromInstructions_ContainsTheInstructedCell()
    {
        var instructions = TempPath("workbook.json");
        File.WriteAllText(instructions, """
        {
          "version": "1.0",
          "worksheets": [
            { "name": "Data", "headers": ["Region"], "rows": [["Northwind"]] }
          ]
        }
        """);
        var outputPath = TempPath("created.xlsx");

        Assert.Equal(0, InvokeMain(["create", outputPath, "--instructions", instructions]));

        Assert.True(File.Exists(outputPath));
        using var reader = WorkbookBuilder.Open(outputPath);
        var sheet = reader.GetWorksheet("Data");
        Assert.Equal("Region", sheet.GetCellValue("A1"));
        Assert.Equal("Northwind", sheet.GetCellValue("A2"));
    }

    [Fact]
    public void Create_DocxFromInstructions_ContainsTheInstructedText()
    {
        var instructions = TempPath("report.json");
        File.WriteAllText(instructions, ValidDocxGenerationDocument("Instructed DOCX paragraph"));
        var outputPath = TempPath("created.docx");

        Assert.Equal(0, InvokeMain(["create", outputPath, "--instructions", instructions]));

        Assert.True(File.Exists(outputPath));
        using var document = WordprocessingDocument.Open(outputPath, false);
        Assert.Contains("Instructed DOCX paragraph", document.MainDocumentPart!.Document!.InnerText);
    }

    [Fact]
    public void Create_PptxFromInstructions_MatchesCoreDelivery_AndNeverRewritesTheSource()
    {
        var instructions = TempPath("contract.json");
        var original = GenerationContract.Json;
        File.WriteAllText(instructions, original);
        var outputPath = TempPath("created.pptx");

        Assert.Equal(0, InvokeMain(["create", outputPath, "--instructions", instructions]));

        var expected = new PptxGenerator().Generate(original);
        Assert.True(expected.Success);
        Assert.Equal(
            GenerationContract.Snapshot(expected.PptxBytes!),
            GenerationContract.Snapshot(File.ReadAllBytes(outputPath)));
        Assert.Equal(original, File.ReadAllText(instructions));
        Assert.False(File.Exists(instructions + ".bak"));
    }

    [Fact]
    public void Create_WithoutInstructions_StillWritesBlankDocuments()
    {
        var xlsxPath = TempPath("blank.xlsx");
        Assert.Equal(0, InvokeMain(["create", xlsxPath, "--sheet", "Data"]));
        using (var reader = WorkbookBuilder.Open(xlsxPath))
        {
            Assert.Null(reader.GetWorksheet("Data").GetCellValue("A1"));
        }

        var pptxPath = TempPath("blank.pptx");
        Assert.Equal(0, InvokeMain(["create", pptxPath, "--title", "Blank title"]));
        using (var builder = PresentationBuilder.Open(pptxPath))
        {
            Assert.Equal(1, builder.SlideCount);
        }

        var docxPath = TempPath("blank.docx");
        Assert.Equal(0, InvokeMain(["create", docxPath, "--text", "Hello World"]));
        using (var document = WordprocessingDocument.Open(docxPath, false))
        {
            Assert.Contains("Hello World", document.MainDocumentPart!.Document!.InnerText);
        }
    }

    [Fact]
    public void Create_WithInstructionsForAnotherVocabulary_FailsWithoutWritingABlankDocument()
    {
        var instructions = TempPath("deck.json");
        File.WriteAllText(instructions, ValidGenerationDocument());
        var outputPath = TempPath("mismatch.xlsx");

        Assert.Equal(1, InvokeMain(["create", outputPath, "--instructions", instructions]));

        Assert.False(File.Exists(outputPath));
    }

    [Fact]
    public void Create_WithMissingInstructionsFile_FailsWithoutWriting()
    {
        var outputPath = TempPath("missing.docx");

        Assert.Equal(1, InvokeMain(["create", outputPath, "--instructions", TempPath("nope.json")]));

        Assert.False(File.Exists(outputPath));
    }

    [Fact]
    public void Create_WithInstructionsFlagButNoValue_FailsWithoutWriting()
    {
        var outputPath = TempPath("novalue.docx");

        Assert.Equal(1, InvokeMain(["create", outputPath, "--instructions"]));

        Assert.False(File.Exists(outputPath));
    }

    // ─── generate --no-write-ids / --backup ───

    [Fact]
    public void Generate_NoWriteIds_LeavesTheInputByteIdentical()
    {
        var source = TempPath("source.json");
        var original = ValidGenerationDocument();
        File.WriteAllText(source, original);
        var outputPath = TempPath("source.pptx");

        Assert.Equal(0, InvokeMain(["generate", source, "--output", outputPath, "--no-write-ids"]));

        Assert.Equal(original, File.ReadAllText(source));
        Assert.DoesNotContain("\"id\"", File.ReadAllText(source));
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Generate_Backup_CopiesTheSourceBeforeTheIdWriteBack()
    {
        var source = TempPath("source.json");
        var original = ValidGenerationDocument();
        File.WriteAllText(source, original);

        Assert.Equal(0, InvokeMain(["generate", source, "--output", TempPath("source.pptx"), "--backup"]));

        Assert.Equal(original, File.ReadAllText(source + ".bak"));
        Assert.Contains("\"id\"", File.ReadAllText(source));
    }

    [Fact]
    public void Generate_BackupWithExistingBackup_FailsAndPreservesTheSource()
    {
        var source = TempPath("source.json");
        var original = ValidGenerationDocument();
        File.WriteAllText(source, original);
        File.WriteAllText(source + ".bak", "existing backup");
        var outputPath = TempPath("source.pptx");

        Assert.Equal(1, InvokeMain(["generate", source, "--output", outputPath, "--backup"]));

        Assert.Equal(original, File.ReadAllText(source));
        Assert.Equal("existing backup", File.ReadAllText(source + ".bak"));
        Assert.False(File.Exists(outputPath));
    }

    [Fact]
    public void Generate_BackupWithNoWriteIds_Fails()
    {
        var source = TempPath("source.json");
        File.WriteAllText(source, ValidGenerationDocument());

        Assert.Equal(1, InvokeMain(["generate", source, "--output", TempPath("source.pptx"), "--backup", "--no-write-ids"]));

        Assert.False(File.Exists(source + ".bak"));
    }

    private static bool RenderTestsEnabled()
        => Environment.GetEnvironmentVariable(EnableRenderEnvVar) == "1";

    private static string Magic(string path, int length)
        => Encoding.ASCII.GetString(File.ReadAllBytes(path), 0, length);

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

    /// <summary>
    /// Runs the CLI against a plain, non-ANSI console so tests can assert on user-facing
    /// messages. The console is process-wide, so it is restored immediately afterwards.
    /// </summary>
    private static (int Code, string Output) InvokeMainCapturingOutput(string[] args)
    {
        var writer = new StringWriter();
        var previous = AnsiConsole.Console;
        AnsiConsole.Console = AnsiConsole.Create(new AnsiConsoleSettings
        {
            Ansi = AnsiSupport.No,
            ColorSystem = ColorSystemSupport.NoColors,
            Interactive = InteractionSupport.No,
            Out = new AnsiConsoleOutput(writer)
        });

        try
        {
            return (InvokeMain(args), writer.ToString());
        }
        finally
        {
            AnsiConsole.Console = previous;
        }
    }

    private static string ValidDocxGenerationDocument(string text = "Hello DOCX")
        => $$"""
        {
          "version": "1.0",
          "sections": [
            { "blocks": [ { "type": "paragraph", "text": "{{text}}" } ] }
          ]
        }
        """;

    private static string ValidXlsxInstructionSet()
        => """
        {
          "version": "1.0",
          "worksheets": [
            { "name": "Sheet1", "headers": ["A", "B"], "rows": [["1", "2"]] }
          ]
        }
        """;

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
