using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using PptxEditor.Core.Models;
using PptxEditor.Core.Services;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Snapshot tests: extract the <see cref="BrandProfile"/> from the reference deck(s) and
/// compare against committed brand.json fixtures. Set OE_UPDATE_SNAPSHOTS=1 to regenerate
/// the fixtures (review the diff before committing).
/// </summary>
public sealed class BrandProfileSnapshotTests
{
    private const string UpdateSnapshotsEnvVar = "OE_UPDATE_SNAPSHOTS";

    [Theory]
    [InlineData("northwind-demo.pptx", "northwind-demo.brand.json")]
    [InlineData("AetherLink-Glass-Shareholder-Overview.pptx", "aetherlink-glass-shareholder-overview.brand.json")]
    public void Extract_ReferenceDeck_MatchesCommittedBrandSnapshot(string deckFileName, string snapshotFileName)
    {
        var deckPath = Path.Combine(ResolveReferenceDirectory(), deckFileName);
        Assert.True(File.Exists(deckPath), $"Reference PPTX file not found: {deckPath}");

        using var document = PresentationDocument.Open(deckPath, false);
        var profile = new BrandProfileExtractor().Extract(document);
        var actual = Serialize(profile);

        var snapshotPath = Path.Combine(ResolveFixtureDirectory(), snapshotFileName);
        if (Environment.GetEnvironmentVariable(UpdateSnapshotsEnvVar) == "1")
        {
            Directory.CreateDirectory(Path.GetDirectoryName(snapshotPath)!);
            File.WriteAllText(snapshotPath, actual);
        }

        Assert.True(File.Exists(snapshotPath),
            $"Brand snapshot not found: {snapshotPath}. Run the test with {UpdateSnapshotsEnvVar}=1 to create it.");

        var expected = File.ReadAllText(snapshotPath).Replace("\r\n", "\n");
        Assert.Equal(expected, actual);
    }

    /// <summary>
    /// Deterministic JSON shape: camelCase properties in declaration order, theme color
    /// keys sorted (the extractor stores them in a SortedDictionary), trailing newline.
    /// </summary>
    internal static string Serialize(BrandProfile profile)
        => JsonSerializer.Serialize(profile, SerializerOptions) + "\n";

    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    private static string ResolveReferenceDirectory()
    {
        var baseDir = AppContext.BaseDirectory;
        return Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "..", "examples", "REF", "PPTX"));
    }

    private static string ResolveFixtureDirectory()
    {
        var baseDir = AppContext.BaseDirectory;
        return Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "Fixtures", "BrandProfile"));
    }
}
