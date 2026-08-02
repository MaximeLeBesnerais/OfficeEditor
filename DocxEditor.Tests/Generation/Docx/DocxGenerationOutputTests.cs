using DocxEditor.Core.Generation;
using DocxEditor.Core.Generation.Contracts;
using DocxEditor.Core.Generation.Schema;
using OfficeEditor.Core.Exceptions;

namespace DocxEditor.Tests.Generation.Docx;

/// <summary>
/// Generator output semantics: byte buffers, caller-owned streams, atomic file writes, and
/// failure behavior that never corrupts an existing destination.
/// </summary>
public class DocxGenerationOutputTests
{
    [Fact]
    public void GenerateToBytes_ReturnsZippedPackage()
    {
        var generated = DocxTestHarness.GenerateToBytes(DocxTestHarness.MinimalJson);

        Assert.True(generated.Content.Length > 1_000);
        Assert.Equal((byte)'P', generated.Content[0]);
        Assert.Equal((byte)'K', generated.Content[1]);
        Assert.Empty(generated.Result.Outputs); // byte mode produces no file artifact
        Assert.Single(generated.Result.Document.Sections);
    }

    [Fact]
    public void StreamOutput_WritesPackageAndLeavesStreamOpen()
    {
        var stream = new MemoryStream();
        var result = new DocxGenerator().Generate(DocxTestHarness.MinimalJson, stream);

        Assert.True(stream.CanWrite);
        Assert.True(stream.Length > 1_000);
        Assert.Equal(stream.Length, stream.Position); // positioned after the last written byte
        Assert.Single(result.Document.Sections);

        // The caller-owned stream stays usable.
        stream.WriteByte(0x42);
        Assert.Equal(stream.Length, stream.Position);
    }

    [Fact]
    public void StreamOutput_ProducesReopenablePackage()
    {
        var stream = new MemoryStream();
        _ = new DocxGenerator().Generate(DocxTestHarness.MinimalJson, stream);

        Assert.True(stream.Length > 1_000);
        stream.Position = 0;
        using var document = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(stream, false);
        Assert.NotNull(document.MainDocumentPart);
        Assert.Contains("Hello world", document.MainDocumentPart!.Document!.InnerText);
    }

    [Fact]
    public void NonWritableStream_IsRejected()
    {
        var stream = new MemoryStream(Array.Empty<byte>(), writable: false);

        var exception = Assert.Throws<ArgumentException>(
            () => new DocxGenerator().Generate(DocxTestHarness.MinimalJson, stream));
        Assert.Contains("writable", exception.Message);
    }

    [Fact]
    public void NullStream_IsRejected()
    {
        Assert.Throws<ArgumentNullException>(
            () => new DocxGenerator().Generate(DocxTestHarness.MinimalJson, (Stream)null!));
    }

    [Fact]
    public void FileOutput_CreatesFileAndReportsOutput()
    {
        using var temp = new TempDirectory();
        var output = temp.File("out.docx");

        var result = new DocxGenerator().Generate(DocxTestHarness.MinimalJson, output);

        Assert.True(File.Exists(output));
        Assert.Equal(DocxOutputKind.Document, result.Outputs.Single().Kind);
        Assert.Equal(Path.GetFullPath(output), result.Outputs.Single().FilePath);
    }

    [Fact]
    public void FileOutput_OverwritesExistingDestination()
    {
        using var temp = new TempDirectory();
        var output = temp.File("out.docx");
        File.WriteAllBytes(output, new byte[] { 1, 2, 3, 4, 5 });

        var result = new DocxGenerator().Generate(DocxTestHarness.MinimalJson, output);

        var onDisk = File.ReadAllBytes(output);
        Assert.NotEqual(new byte[] { 1, 2, 3, 4, 5 }, onDisk);
        Assert.Equal((byte)'P', onDisk[0]);
        Assert.Equal((byte)'K', onDisk[1]);
        Assert.NotEmpty(result.Outputs);
    }

    [Fact]
    public void FailedValidation_DoesNotTouchExistingDestination()
    {
        using var temp = new TempDirectory();
        var output = temp.File("out.docx");
        var sentinel = new byte[] { 0xDE, 0xAD, 0xBE, 0xEF };
        File.WriteAllBytes(output, sentinel);

        Assert.Throws<DocxGenerationValidationException>(
            () => new DocxGenerator().Generate("{ not json", output));

        Assert.Equal(sentinel, File.ReadAllBytes(output));
    }

    [Fact]
    public void FailedEmit_ImageError_LeavesDestinationAndDirectoryClean()
    {
        using var temp = new TempDirectory();
        var output = temp.File("out.docx");
        var sentinel = new byte[] { 0xDE, 0xAD, 0xBE, 0xEF };
        File.WriteAllBytes(output, sentinel);

        var json = """
            {
              "version": "1.0",
              "design": { "palette": { "ink": "#1F2937" } },
              "sections": [ { "blocks": [
                { "type": "image", "src": "data:image/png;base64,not-really-a-png-payload" }
              ] } ]
            }
            """;

        Assert.Throws<DocxEditor.Core.Generation.Assets.ImageSourceException>(
            () => new DocxGenerator().Generate(json, output));

        Assert.Equal(sentinel, File.ReadAllBytes(output));
        // No staged temp files remain in the destination directory.
        Assert.Single(Directory.GetFiles(temp.Path));
    }

    [Fact]
    public void FailedEmit_MissingTemplate_LeavesDestinationUntouched()
    {
        using var temp = new TempDirectory();
        var output = temp.File("out.docx");
        var sentinel = "SENTINEL"u8.ToArray();
        File.WriteAllBytes(output, sentinel);

        var json = $$"""
            {
              "version": "1.0",
              "template": "{{temp.File("missing-template.docx")}}",
              "sections": [ { "blocks": [ { "type": "paragraph", "text": "x" } ] } ]
            }
            """;

        Assert.Throws<OfficeEditorException>(
            () => new DocxGenerator().Generate(json, output));

        Assert.Equal(sentinel, File.ReadAllBytes(output));
        Assert.Single(Directory.GetFiles(temp.Path));
    }

    [Fact]
    public void ModelOverload_GeneratesToStream()
    {
        var model = DocxTestHarness.Parse(DocxTestHarness.MinimalJson);
        var stream = new MemoryStream();

        var result = new DocxGenerator().Generate(model, stream);

        Assert.True(stream.Length > 1_000);
        Assert.Single(result.Document.Sections);
    }

    [Fact]
    public void ModelOverload_GeneratesToBytes()
    {
        var model = DocxTestHarness.Parse(DocxTestHarness.MinimalJson);

        var generated = new DocxGenerator().GenerateToBytes(model);

        Assert.True(generated.Content.Length > 1_000);
        Assert.Equal((byte)'P', generated.Content[0]);
    }

    [Fact]
    public void TemplatePathOption_OverridesModelTemplate()
    {
        using var temp = new TempDirectory();
        var templatePath = BuildBlankTemplate(temp.File("tpl.docx"));

        var generated = new DocxGenerator().GenerateToBytes(
            DocxTestHarness.MinimalJson,
            new DocxGeneratorOptions { TemplatePath = templatePath });

        Assert.True(generated.Content.Length > 1_000);
        Assert.Single(generated.Result.Document.Sections);
    }

    [Fact]
    public void Emitter_IsSingleShot()
    {
        using var emitter = new DocxEditor.Core.Generation.Emit.Ooxml.DocxOoxmlEmitter();
        var document = DocxTestHarness.Parse(DocxTestHarness.MinimalJson);

        _ = emitter.Emit(document);
        Assert.Throws<InvalidOperationException>(() => emitter.Emit(document));
    }

    [Fact]
    public void NullJson_IsRejected()
    {
        Assert.Throws<ArgumentNullException>(
            () => new DocxGenerator().Generate((string)null!, Path.GetTempPath() + "x.docx"));
    }

    [Fact]
    public void EmptyOutputPath_IsRejected()
    {
        Assert.Throws<ArgumentException>(
            () => new DocxGenerator().Generate(DocxTestHarness.MinimalJson, "   "));
    }

    private static string BuildBlankTemplate(string path)
    {
        using (var document = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(path, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(
                new DocumentFormat.OpenXml.Wordprocessing.Body());
        }
        return path;
    }
}
