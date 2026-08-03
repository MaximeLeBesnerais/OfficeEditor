using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Generation;
using DocxEditor.Core.Generation.Contracts;

namespace DocxEditor.Tests.Generation.Docx;

/// <summary>
/// Wrong-kind style references: resolution validates the expected style kind (paragraph,
/// character, table) for both author references and baseline-style reuse. A wrong-kind
/// reference never lands in the OOXML — it records the canonical <c>WrongStyleKind</c> warning,
/// uses the proper baseline fallback, and never mutates the template's definitions.
/// </summary>
public class DocxGenerationStyleKindTests
{
    private static readonly string TemplateDir = Path.Combine(Path.GetTempPath(), "officeeditor-tests");

    [Fact]
    public void ParagraphReferencingCharacterStyle_WarnsFallsBackToNormal_AndLeavesTemplateUntouched()
    {
        var templatePath = CreateTemplateWithKindStyles();
        var templateStyleXml = ReadStyleOuterXml(templatePath, "CharStyle");

        var generated = GenerateWithTemplate(templatePath, "paragraph", "CharStyle");
        var doc = OpenDocument(generated.Content);
        var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>()
            .Single(p => p.InnerText == "x");

        // The wrong-kind style is never referenced; the paragraph falls back to the Normal baseline.
        Assert.Equal("Normal", paragraph.ParagraphProperties!.ParagraphStyleId!.Val!.Value);

        var warning = Assert.Single(generated.Result.Warnings, w => w.Message.Contains("WrongStyleKind"));
        Assert.Equal("$.sections[0].blocks[0]", warning.Path);
        Assert.Contains("CharStyle", warning.Message);
        Assert.Contains("character", warning.Message);
        Assert.Contains("Normal", warning.Message);

        // The template definition is byte-identical (never mutated).
        var preserved = doc.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<Style>().Single(s => s.StyleId?.Value == "CharStyle");
        Assert.Equal(templateStyleXml, preserved.OuterXml);

        Cleanup(templatePath);
    }

    [Fact]
    public void ParagraphReferencingTableStyle_WarnsAndFallsBackToNormal()
    {
        var templatePath = CreateTemplateWithKindStyles();

        var generated = GenerateWithTemplate(templatePath, "paragraph", "TableStyle");
        var doc = OpenDocument(generated.Content);
        var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>()
            .Single(p => p.InnerText == "x");

        Assert.Equal("Normal", paragraph.ParagraphProperties!.ParagraphStyleId!.Val!.Value);
        Assert.Contains(generated.Result.Warnings, w =>
            w.Message.Contains("WrongStyleKind") && w.Message.Contains("TableStyle") && w.Message.Contains("table"));

        Cleanup(templatePath);
    }

    [Fact]
    public void TableReferencingParagraphStyle_WarnsAndFallsBackToTableGrid()
    {
        var templatePath = CreateTemplateWithKindStyles();

        var generated = GenerateWithTemplate(templatePath, "table", "ParaStyle");
        var doc = OpenDocument(generated.Content);
        var table = doc.MainDocumentPart!.Document!.Body!.Descendants<Table>().Single();

        // The wrong-kind style is never referenced; the table falls back to the TableGrid baseline.
        Assert.Equal("TableGrid", table.TableProperties!.TableStyle!.Val!.Value);
        Assert.Contains(generated.Result.Warnings, w =>
            w.Message.Contains("WrongStyleKind") && w.Message.Contains("ParaStyle") && w.Message.Contains("paragraph"));

        Cleanup(templatePath);
    }

    [Fact]
    public void TableReferencingTableStyle_AppliesWithoutWarning()
    {
        var templatePath = CreateTemplateWithKindStyles();

        var generated = GenerateWithTemplate(templatePath, "table", "TableStyle");
        var doc = OpenDocument(generated.Content);
        var table = doc.MainDocumentPart!.Document!.Body!.Descendants<Table>().Single();

        Assert.Equal("TableStyle", table.TableProperties!.TableStyle!.Val!.Value);
        Assert.DoesNotContain(generated.Result.Warnings, w => w.Message.Contains("WrongStyleKind"));

        Cleanup(templatePath);
    }

    [Fact]
    public void RunReferencingParagraphStyle_WarnsAndSkipsTheRunStyle()
    {
        var templatePath = CreateTemplateWithKindStyles();

        var generated = GenerateWithTemplate(templatePath, "run", "ParaStyle");
        var doc = OpenDocument(generated.Content);
        var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>()
            .Single(p => p.InnerText == "x");
        var run = paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>().Single();

        // No character baseline exists, so a wrong-kind run style is skipped (never referenced).
        Assert.Null(run.RunProperties?.RunStyle);
        Assert.Contains(generated.Result.Warnings, w =>
            w.Message.Contains("WrongStyleKind") && w.Message.Contains("ParaStyle"));

        Cleanup(templatePath);
    }

    [Fact]
    public void TemplateBaselineOfWrongKind_IsNotReused_AndFreshStyleGenerated()
    {
        var templatePath = CreateTemplateWithWrongKindBaseline();

        var generated = GenerateWithTemplate(templatePath, "table", null);
        var doc = OpenDocument(generated.Content);
        var styles = doc.MainDocumentPart!.StyleDefinitionsPart!.Styles!;

        // The template's paragraph-kind "TableGrid" is never reused as a table style; a fresh,
        // collision-free table baseline is generated instead and referenced.
        var generatedTableGrid = styles.Elements<Style>().Single(s => s.StyleId?.Value == "TableGrid1");
        Assert.Equal(StyleValues.Table, generatedTableGrid.Type?.Value);

        var table = doc.MainDocumentPart!.Document!.Body!.Descendants<Table>().Single();
        Assert.Equal("TableGrid1", table.TableProperties!.TableStyle!.Val!.Value);

        Assert.Contains(generated.Result.Warnings, w =>
            w.Message.Contains("WrongStyleKind") && w.Message.Contains("Table Grid"));

        Cleanup(templatePath);
    }

    // ---- helpers ----

    private static string TemplateJson(string blockJson) => $$"""
        {"version":"1.0","sections":[{"blocks":[ {{blockJson}} ]}]}
        """;

    private static GeneratedDocx GenerateWithTemplate(string templatePath, string kind, string? style)
    {
        var styleJson = style is null ? string.Empty : $",\"style\":\"{style}\"";
        var block = kind switch
        {
            "paragraph" => $$"""{"type":"paragraph","text":"x"{{styleJson}}}""",
            "table" => $$"""{"type":"table"{{styleJson}},"rows":[{"cells":[{"text":"a"}]}]}""",
            "run" => $$"""{"type":"paragraph","runs":[{"text":"x"{{styleJson}}}]}""",
            _ => throw new ArgumentOutOfRangeException(nameof(kind))
        };
        return new DocxGenerator().GenerateToBytes(
            TemplateJson(block),
            new DocxGeneratorOptions { TemplatePath = templatePath });
    }

    /// <summary>A template with one paragraph, one character, and one table style (correct kinds).</summary>
    private static string CreateTemplateWithKindStyles()
    {
        var path = NextTemplatePath();
        using (var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(new Run(new Text("Template body")))));
            var styles = main.AddNewPart<StyleDefinitionsPart>();
            styles.Styles = new Styles(
                new Style { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true, StyleName = new StyleName { Val = "Normal" } },
                new Style { Type = StyleValues.Paragraph, StyleId = "ParaStyle", StyleName = new StyleName { Val = "Para Style" } },
                new Style { Type = StyleValues.Character, StyleId = "CharStyle", StyleName = new StyleName { Val = "Char Style" } },
                new Style { Type = StyleValues.Table, StyleId = "TableStyle", StyleName = new StyleName { Val = "Table Style" } });
            document.Save();
        }
        return path;
    }

    /// <summary>A template whose "Table Grid" style is (wrongly) a paragraph style.</summary>
    private static string CreateTemplateWithWrongKindBaseline()
    {
        var path = NextTemplatePath();
        using (var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(new Run(new Text("Template body")))));
            var styles = main.AddNewPart<StyleDefinitionsPart>();
            styles.Styles = new Styles(
                new Style { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true, StyleName = new StyleName { Val = "Normal" } },
                new Style { Type = StyleValues.Paragraph, StyleId = "TableGrid", StyleName = new StyleName { Val = "Table Grid" } });
            document.Save();
        }
        return path;
    }

    private static string NextTemplatePath()
    {
        Directory.CreateDirectory(TemplateDir);
        return Path.Combine(TemplateDir, $"template-kind-{Guid.NewGuid():N}.docx");
    }

    private static string ReadStyleOuterXml(string path, string styleId)
    {
        using var document = WordprocessingDocument.Open(path, false);
        return document.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<Style>()
            .Single(s => s.StyleId?.Value == styleId)
            .OuterXml;
    }

    private static WordprocessingDocument OpenDocument(byte[] bytes)
    {
        var stream = new MemoryStream(bytes);
        return WordprocessingDocument.Open(stream, false);
    }

    private static void Cleanup(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
            // Best-effort cleanup must never mask the primary assertion.
        }
    }
}
