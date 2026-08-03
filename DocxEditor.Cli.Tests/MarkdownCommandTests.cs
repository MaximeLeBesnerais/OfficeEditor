using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Builders;
using Xunit;

namespace DocxEditor.Cli.Tests;

/// <summary>
/// End-to-end CLI tests for the <c>docxeditor markdown</c> command: rich Markdown rendering,
/// relative-image root confinement, template-based style resolution, strict-mode diagnostics,
/// atomic output and nonzero exit codes on errors.
/// </summary>
public sealed class MarkdownCommandTests : IDisposable
{
    private readonly string _tempDir;

    public MarkdownCommandTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "DocxEditorCliTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    [Fact]
    public void Markdown_Basic_ProducesDocx()
    {
        var input = WriteInput("basic.md", "# Hello\n\nWorld with **bold**.");
        var output = TempPath("basic.docx");

        var code = InvokeMain(["markdown", input, output]);

        Assert.Equal(0, code);
        Assert.True(File.Exists(output));
        using var doc = WordprocessingDocument.Open(output, false);
        Assert.Contains("Hello", doc.MainDocumentPart!.Document!.InnerText);
    }

    [Fact]
    public void Markdown_RelativeImageWithinRoot_IsEmbedded()
    {
        var inputDir = Path.Combine(_tempDir, "docs");
        Directory.CreateDirectory(inputDir);
        var imageDir = Path.Combine(inputDir, "images");
        Directory.CreateDirectory(imageDir);
        var imagePath = Path.Combine(imageDir, "logo.png");
        File.WriteAllBytes(imagePath, OnePixelPng());

        var input = Path.Combine(inputDir, "page.md");
        File.WriteAllText(input, "# Page\n\n![logo](images/logo.png)");
        var output = TempPath("with-image.docx");

        var code = InvokeMain(["markdown", input, output]);

        Assert.Equal(0, code);
        Assert.True(File.Exists(output));
        using var doc = WordprocessingDocument.Open(output, false);
        Assert.Contains(doc.MainDocumentPart!.Document!.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline>(), _ => true);
    }

    [Fact]
    public void Markdown_WithTemplateAndStyleMap_ResolvesCustomStyle()
    {
        var template = CreateTemplateWithCustomStyle();
        var styleMap = TempPath("styles.json");
        File.WriteAllText(styleMap, """{ "paragraph": "CustomBody" }""");
        var input = WriteInput("styled.md", "Hello styled body.");
        var output = TempPath("styled.docx");

        var code = InvokeMain(["markdown", input, output, "--template", template, "--style-map", styleMap]);

        Assert.Equal(0, code);
        Assert.True(File.Exists(output));
        using var doc = WordprocessingDocument.Open(output, false);
        var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>()
            .First(p => p.InnerText.Contains("Hello styled body."));
        Assert.Equal("CustomBody", paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
    }

    [Fact]
    public void Markdown_Strict_MissingTemplateStyle_ReturnsOne()
    {
        var template = CreateTemplateWithCustomStyle();
        var styleMap = TempPath("styles.json");
        File.WriteAllText(styleMap, """{ "paragraph": "DoesNotExist" }""");
        var input = WriteInput("strict.md", "Hello.");
        var output = TempPath("strict.docx");

        var code = InvokeMain(["markdown", input, output, "--template", template, "--style-map", styleMap, "--strict"]);

        Assert.Equal(1, code);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public void Markdown_Failure_PreservesExistingOutput()
    {
        var template = CreateTemplateWithCustomStyle();
        var styleMap = TempPath("styles.json");
        File.WriteAllText(styleMap, """{ "paragraph": "DoesNotExist" }""");
        var input = WriteInput("atomic.md", "Hello.");
        var output = TempPath("existing.docx");
        File.WriteAllText(output, "previous output");

        var code = InvokeMain(["markdown", input, output, "--template", template, "--style-map", styleMap, "--strict"]);

        Assert.Equal(1, code);
        Assert.True(File.Exists(output));
        Assert.Equal("previous output", File.ReadAllText(output));
    }

    [Fact]
    public void Markdown_Permissive_MissingTemplateStyle_WarnsButSucceeds()
    {
        var template = CreateTemplateWithCustomStyle();
        var styleMap = TempPath("styles.json");
        File.WriteAllText(styleMap, """{ "paragraph": "MissingPermissive" }""");
        var input = WriteInput("perm.md", "Hello.");
        var output = TempPath("perm.docx");

        var code = InvokeMain(["markdown", input, output, "--template", template, "--style-map", styleMap]);

        Assert.Equal(0, code);
        Assert.True(File.Exists(output));
    }

    [Fact]
    public void Markdown_Strict_BlankDocument_DefaultMapping_Succeeds()
    {
        // A fresh document has no styles part; strict mode must seed the trusted built-in
        // fallback styles for the constructs the input uses instead of failing on the default
        // "Normal" mapping.
        var input = WriteInput("strict-blank.md", "# Hello\n\nWorld.");
        var output = TempPath("strict-blank.docx");

        var code = InvokeMain(["markdown", input, output, "--strict"]);

        Assert.Equal(0, code);
        Assert.True(File.Exists(output));
        using var doc = WordprocessingDocument.Open(output, false);
        Assert.Contains("Hello", doc.MainDocumentPart!.Document!.InnerText);
        Assert.NotNull(doc.MainDocumentPart.StyleDefinitionsPart);
    }

    [Fact]
    public void Markdown_Strict_Template_UnusedDefaultEntries_NotRejected()
    {
        // The template defines only the styles the input actually uses. Default-map entries for
        // constructs the input never uses (table, code, quote, ...) must not fail the conversion.
        var template = CreateTemplateWithStyles("Normal", "Heading1");
        var input = WriteInput("strict-template.md", "# Heading\n\nBody text.");
        var output = TempPath("strict-template.docx");

        var code = InvokeMain(["markdown", input, output, "--template", template, "--strict"]);

        Assert.Equal(0, code);
        Assert.True(File.Exists(output));
        using var doc = WordprocessingDocument.Open(output, false);
        var heading = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>()
            .First(p => p.InnerText.Contains("Heading"));
        Assert.Equal("Heading1", heading.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
    }

    [Fact]
    public void Markdown_Strict_Template_ReferencedStyleMissing_StillRejected()
    {
        // The input uses a table, which maps to "TableGrid" by default; the template lacks it, so
        // strict mode must still reject the reference because the input actually uses the table.
        var template = CreateTemplateWithStyles("Normal", "Heading1");
        var input = WriteInput("strict-table.md", "# H\n\n| A | B |\n|---|---|\n| 1 | 2 |");
        var output = TempPath("strict-table.docx");

        var code = InvokeMain(["markdown", input, output, "--template", template, "--strict"]);

        Assert.Equal(1, code);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public void Markdown_Strict_RenderStyleError_DoesNotCreateOutput()
    {
        // Without a template, unresolved style references surface during render, not in the
        // pre-render template check. A strict render error must not publish any output.
        var styleMap = TempPath("styles.json");
        File.WriteAllText(styleMap, """{ "paragraph": "DoesNotExist" }""");
        var input = WriteInput("render-strict.md", "Hello.");
        var output = TempPath("render-strict.docx");

        var code = InvokeMain(["markdown", input, output, "--style-map", styleMap, "--strict"]);

        Assert.Equal(1, code);
        Assert.False(File.Exists(output));
        Assert.Empty(TempLeftovers());
    }

    [Fact]
    public void Markdown_Strict_RenderStyleError_PreservesExistingOutputByteForByte()
    {
        var styleMap = TempPath("styles.json");
        File.WriteAllText(styleMap, """{ "paragraph": "DoesNotExist" }""");
        var input = WriteInput("render-atomic.md", "Hello.");
        var output = TempPath("render-atomic.docx");
        var previous = new byte[] { 0xDE, 0xAD, 0xBE, 0xEF, 0x01, 0x02, 0x03 };
        File.WriteAllBytes(output, previous);

        var code = InvokeMain(["markdown", input, output, "--style-map", styleMap, "--strict"]);

        Assert.Equal(1, code);
        Assert.True(File.Exists(output));
        Assert.Equal(previous, File.ReadAllBytes(output));
        Assert.Empty(TempLeftovers());
    }

    [Fact]
    public void Markdown_Strict_RenderStyleError_LeavesNoTempFiles()
    {
        var styleMap = TempPath("styles.json");
        File.WriteAllText(styleMap, """{ "paragraph": "DoesNotExist" }""");
        var input = WriteInput("render-tmp.md", "Hello.");
        var output = TempPath("render-tmp.docx");

        var code = InvokeMain(["markdown", input, output, "--style-map", styleMap, "--strict"]);

        Assert.Equal(1, code);
        Assert.Empty(Directory.EnumerateFiles(_tempDir, "*.tmp"));
    }

    [Fact]
    public void Markdown_Permissive_RenderStyleWarning_PublishesOutput()
    {
        // In permissive mode an unresolved render-time style is only a warning; the conversion
        // must still publish atomically and leave no temp file behind.
        var styleMap = TempPath("styles.json");
        File.WriteAllText(styleMap, """{ "paragraph": "MissingPermissiveRender" }""");
        var input = WriteInput("perm-render.md", "Hello.");
        var output = TempPath("perm-render.docx");

        var code = InvokeMain(["markdown", input, output, "--style-map", styleMap]);

        Assert.Equal(0, code);
        Assert.True(File.Exists(output));
        using var doc = WordprocessingDocument.Open(output, false);
        Assert.Contains("Hello", doc.MainDocumentPart!.Document!.InnerText);
        Assert.Empty(TempLeftovers());
    }

    [Fact]
    public void Markdown_InvalidStyleMapJson_ReturnsOne()
    {
        var styleMap = TempPath("bad.json");
        File.WriteAllText(styleMap, "{ not json");
        var input = WriteInput("bad.md", "Hello.");
        var output = TempPath("bad.docx");

        var code = InvokeMain(["markdown", input, output, "--style-map", styleMap]);

        Assert.Equal(1, code);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public void Markdown_MissingInputFile_ReturnsOne()
    {
        var output = TempPath("missing.docx");
        var code = InvokeMain(["markdown", TempPath("nope.md"), output]);

        Assert.Equal(1, code);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public void Markdown_MissingTemplateFile_ReturnsOne()
    {
        var input = WriteInput("t.md", "Hello.");
        var output = TempPath("t.docx");

        var code = InvokeMain(["markdown", input, output, "--template", TempPath("nope.docx")]);

        Assert.Equal(1, code);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public void Markdown_UnknownOption_ReturnsOne()
    {
        var input = WriteInput("u.md", "Hello.");
        var output = TempPath("u.docx");

        var code = InvokeMain(["markdown", input, output, "--bogus"]);

        Assert.Equal(1, code);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public void Markdown_NonDocxOutput_ReturnsOne()
    {
        var input = WriteInput("w.md", "Hello.");
        var output = TempPath("out.pdf");

        var code = InvokeMain(["markdown", input, output]);

        Assert.Equal(1, code);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public void Markdown_NoArgs_ReturnsOne()
    {
        Assert.Equal(1, InvokeMain(["markdown"]));
    }

    [Fact]
    public void Markdown_MissingOutput_ReturnsOne()
    {
        var input = WriteInput("one.md", "Hello.");
        Assert.Equal(1, InvokeMain(["markdown", input]));
    }

    private string CreateTemplateWithCustomStyle()
    {
        var path = TempPath("template.docx");
        using var builder = DocumentBuilder.Create(path);
        builder.AddParagraph("Template marker", "CustomBody");
        builder.Save();
        return path;
    }

    /// <summary>Builds a template whose styles part defines exactly the given style ids.</summary>
    private string CreateTemplateWithStyles(params string[] styleIds)
    {
        var path = TempPath($"template-{Guid.NewGuid():N}.docx");
        using var builder = DocumentBuilder.Create(path);
        foreach (var styleId in styleIds)
        {
            builder.AddParagraph("Template marker", styleId);
        }
        builder.Save();
        return path;
    }

    private string WriteInput(string fileName, string content)
    {
        var path = TempPath(fileName);
        File.WriteAllText(path, content);
        return path;
    }

    private static int InvokeMain(string[] args)
    {
        var cliAssembly = AppDomain.CurrentDomain.GetAssemblies()
            .FirstOrDefault(a => a.GetName().Name == "DocxEditor.Cli");
        if (cliAssembly == null)
        {
            cliAssembly = Assembly.Load("DocxEditor.Cli");
        }

        var programType = cliAssembly.GetType("DocxEditor.Cli.Program")
            ?? throw new InvalidOperationException("Program type not found in DocxEditor.Cli assembly.");

        var method = programType.GetMethod("Main",
            BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.Public,
            [typeof(string[])])
            ?? throw new InvalidOperationException("Main method not found.");

        var result = method.Invoke(null, [args]);
        return (int)(result ?? throw new InvalidOperationException("Main signature changed; update InvokeMain"));
    }

    private string TempPath(string fileName) => Path.Combine(_tempDir, fileName);

    private IReadOnlyList<string> TempLeftovers() =>
        Directory.Exists(_tempDir)
            ? Directory.EnumerateFiles(_tempDir).Where(p => Path.GetFileName(p).EndsWith(".tmp", StringComparison.Ordinal)).ToList()
            : [];

    private static byte[] OnePixelPng() =>
        Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==");

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
