namespace OfficeEditor.Api.Services;

public sealed class SampleFileService : ISampleFileService
{
    private static readonly IReadOnlyList<SampleFile> SampleFiles = new List<SampleFile>
    {
        new(
            ".pptx",
            "examples/REF/PPTX/.pptx",
            OfficeEditor.Core.Services.OfficeDocumentFormat.Pptx,
            "A professional PowerPoint deck used for advanced PPTX regression tests."),
        new(
            "Annual reporting template ENGLISH_0.docx",
            "examples/REF/DOCX/Annual reporting template ENGLISH_0.docx",
            OfficeEditor.Core.Services.OfficeDocumentFormat.Docx,
            "An annual reporting DOCX template for conversion regression tests."),
        new(
            "Monitoring Report Template.docx",
            "examples/REF/DOCX/Monitoring Report Template.docx",
            OfficeEditor.Core.Services.OfficeDocumentFormat.Docx,
            "A monitoring report DOCX template for conversion regression tests."),
        new(
            "-entreprise-bcp-pme.docx",
            "examples/REF/DOCX/-entreprise-bcp-pme.docx",
            OfficeEditor.Core.Services.OfficeDocumentFormat.Docx,
            "A French risk-management DOCX document for conversion regression tests."),
        new(
            "sample.md",
            "examples/Docx/sample.md",
            OfficeEditor.Core.Services.OfficeDocumentFormat.Markdown,
            "A Markdown sample used to test Markdown-to-DOCX conversion.")
    };

    public IReadOnlyList<SampleFile> GetSampleFiles() => SampleFiles;

    public async Task<byte[]> LoadAsync(string name, CancellationToken ct = default)
    {
        var sample = SampleFiles.FirstOrDefault(s => s.Name == name)
            ?? throw new FileNotFoundException($"Sample file '{name}' was not found.");

        var repoRoot = FindRepositoryRoot();
        var fullPath = Path.GetFullPath(Path.Combine(repoRoot, sample.Path.Replace('/', Path.DirectorySeparatorChar)));

        if (!File.Exists(fullPath))
        {
            throw new FileNotFoundException($"Sample file not found on disk: {fullPath}");
        }

        return await File.ReadAllBytesAsync(fullPath, ct).ConfigureAwait(false);
    }

    private static string FindRepositoryRoot()
    {
        var baseDirectory = AppContext.BaseDirectory;
        var directory = new DirectoryInfo(baseDirectory);

        while (directory != null)
        {
            var gitDirectory = new DirectoryInfo(Path.Combine(directory.FullName, ".git"));
            if (gitDirectory.Exists)
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        // Fallback to the parent of the API project folder (bin/Debug/net9.0 -> OfficeEditor.Api -> repo root).
        var fallback = new DirectoryInfo(baseDirectory);
        for (var i = 0; i < 4 && fallback != null; i++)
        {
            fallback = fallback.Parent;
        }

        return fallback?.FullName ?? baseDirectory;
    }
}

public interface ISampleFileService
{
    IReadOnlyList<SampleFile> GetSampleFiles();
    Task<byte[]> LoadAsync(string name, CancellationToken ct = default);
}

public record SampleFile(string Name, string Path, OfficeEditor.Core.Services.OfficeDocumentFormat Format, string Description);
