using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocxEditor.Core.Generation;
using DocxEditor.Core.Generation.Assets;
using DocxEditor.Core.Generation.Contracts;
using DocxEditor.Core.Generation.Model;
using DocxEditor.Core.Generation.Schema;

namespace DocxEditor.Tests.Generation.Docx;

/// <summary>
/// Shared harness for the DOCX generation test suite: builds minimal valid documents,
/// generates packages to bytes/files/streams, and reopens them for OOXML inspection.
/// </summary>
public static class DocxTestHarness
{
    /// <summary>A minimal valid generation document (single A4 section, one paragraph).</summary>
    public static string MinimalJson => """
        {
          "version": "1.0",
          "design": {
            "palette": { "ink": "#1F2937", "navy": "#1F4E79" },
            "fonts": { "display": "Aptos Display", "body": "Aptos" },
            "page": { "defaultFont": "body", "defaultTextColor": "ink" }
          },
          "sections": [
            { "blocks": [ { "type": "paragraph", "text": "Hello world" } ] }
          ]
        }
        """;

    public static DocxGenerationDocument Parse(string json) =>
        new DocxGenerationDocumentParser().Parse(json);

    public static DocxGenerationValidationResult Validate(string json) =>
        new DocxGenerationDocumentParser().Validate(json);

    /// <summary>Generates a package to a temporary file and returns its path.</summary>
    public static string GenerateToTempFile(
        string json,
        DocxGeneratorOptions? options = null)
    {
        var path = Path.Combine(Path.GetTempPath(), $"docxgen-{Guid.NewGuid():N}.docx");
        new DocxGenerator().Generate(json, path, options);
        return path;
    }

    /// <summary>Generates a package in memory and returns it.</summary>
    public static GeneratedDocx GenerateToBytes(string json, DocxGeneratorOptions? options = null) =>
        new DocxGenerator().GenerateToBytes(json, options);

    /// <summary>Locates a repo-relative file by walking up from the test output directory.</summary>
    public static string FindRepoFile(string relative)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        for (var depth = 0; depth < 12 && directory is not null; depth++)
        {
            var candidate = Path.Combine(directory.FullName, relative);
            if (File.Exists(candidate))
            {
                return candidate;
            }
            directory = directory.Parent;
        }
        throw new FileNotFoundException($"Could not locate '{relative}' above {AppContext.BaseDirectory}.");
    }

    /// <summary>Reads the built-in sample document (<c>examples/Docx/generation/comprehensive.json</c>).</summary>
    public static string ReadBuiltInSample() =>
        File.ReadAllText(FindRepoFile(Path.Combine("examples", "Docx", "generation", "comprehensive.json")));

    /// <summary>Reads the report archetype fixture (<c>examples/Docx/generation/editorial-report.json</c>).</summary>
    public static string ReadEditorialReport() =>
        File.ReadAllText(FindRepoFile(Path.Combine("examples", "Docx", "generation", "editorial-report.json")));

    /// <summary>Opens a generated DOCX read-only and returns its main document part.</summary>
    public static MainDocumentPart OpenMainPart(string path)
    {
        var document = WordprocessingDocument.Open(path, false);
        return document.MainDocumentPart
            ?? throw new InvalidOperationException("Generated package has no main document part.");
    }

    /// <summary>Writes <paramref name="bytes"/> to a file under a fresh temporary directory.</summary>
    public static string WriteTempFile(byte[] bytes, string extension = ".png")
    {
        var directory = Path.Combine(Path.GetTempPath(), $"docxgen-assets-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, $"asset{extension}");
        File.WriteAllBytes(path, bytes);
        return path;
    }

    /// <summary>Builds a fresh temporary directory (caller disposes).</summary>
    public static TempDirectory NewTempDirectory() => new();

    /// <summary>Serializes a model back to its canonical JSON round-trip shape (for reuse tests).</summary>
    public static string ToJson(DocxGenerationDocument document) =>
        JsonSerializer.Serialize(document, new JsonSerializerOptions { WriteIndented = true });
}

/// <summary>Auto-deleting temporary directory for filesystem-sensitive tests.</summary>
public sealed class TempDirectory : IDisposable
{
    public TempDirectory()
    {
        Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"docxgen-dir-{Guid.NewGuid():N}");
        Directory.CreateDirectory(Path);
    }

    public string Path { get; }

    public string File(string name) => System.IO.Path.Combine(Path, name);

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(Path))
            {
                Directory.Delete(Path, recursive: true);
            }
        }
        catch
        {
            // Best-effort cleanup; never mask the test result.
        }
    }
}
