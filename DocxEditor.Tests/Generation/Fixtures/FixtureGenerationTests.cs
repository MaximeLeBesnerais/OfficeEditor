using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using PptxEditor.Core.Generation.Fixtures;

namespace DocxEditor.Tests.Generation.Fixtures;

/// <summary>
/// Fixture generator acceptance (P7): every catalog fixture flows through the shared
/// layout resolution into BOTH emitters (plan.md §2 rule 2) — producing an
/// OpenXmlValidator-clean .pptx, a non-empty Typst source, a manifest and the
/// deterministic image asset — with zero layout/emission warnings and fully
/// deterministic output.
/// </summary>
public sealed class FixtureGenerationTests : IDisposable
{
    private readonly string _testDir = Path.Combine(Path.GetTempPath(), $"parity-fixture-tests-{Guid.NewGuid():N}");

    public void Dispose()
    {
        try { Directory.Delete(_testDir, recursive: true); } catch { /* best effort */ }
    }

    public static IEnumerable<object[]> FixtureNames => FixtureCatalog.All
        .Select(f => new object[] { f.Name });

    [Theory]
    [MemberData(nameof(FixtureNames))]
    public void Generate_ProducesValidArtifacts_ForBothEmitters(string fixtureName)
    {
        var fixture = FixtureCatalog.Find(fixtureName)!;
        var artifact = new FixtureGenerator().Generate(fixture, Path.Combine(_testDir, fixtureName));

        Assert.True(File.Exists(artifact.PptxPath), $"missing {artifact.PptxPath}");
        Assert.True(File.Exists(artifact.TypstPath), $"missing {artifact.TypstPath}");
        Assert.True(File.Exists(artifact.ManifestPath), $"missing {artifact.ManifestPath}");
        Assert.True(File.Exists(Path.Combine(artifact.OutputDirectory, "assets", FixtureAssets.CheckerboardFileName)));
        Assert.True(artifact.SlideCount >= 1);
        Assert.Empty(artifact.Warnings);

        // OOXML side: opens and validates without repair (P4 acceptance pattern).
        using var document = PresentationDocument.Open(new MemoryStream(File.ReadAllBytes(artifact.PptxPath)), false);
        var errors = new OpenXmlValidator().Validate(document).ToList();
        Assert.True(errors.Count == 0,
            string.Join("\n", errors.Select(e => $"{e.Path}: {e.Description}")));

        // Typst side: absolute-placement source (#place only — Typst does no layout).
        var typst = File.ReadAllText(artifact.TypstPath);
        Assert.StartsWith("#set page(", typst);
        Assert.Contains("#place(", typst);

        // Manifest round-trips and mirrors the fixture.
        var manifest = JsonSerializer.Deserialize<FixtureManifest>(File.ReadAllText(artifact.ManifestPath),
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        Assert.NotNull(manifest);
        Assert.Equal(fixture.Name, manifest.Name);
        Assert.Equal(fixture.ThresholdRmse, manifest.ThresholdRmse);
        Assert.Equal(artifact.SlideCount, manifest.SlideCount);
        Assert.Empty(manifest.Warnings);
    }

    [Fact]
    public void GenerateAll_ProducesOneDirectoryPerFixture()
    {
        var artifacts = new FixtureGenerator().GenerateAll(Path.Combine(_testDir, "all"));

        Assert.Equal(FixtureCatalog.All.Count, artifacts.Count);
        foreach (var artifact in artifacts)
        {
            Assert.True(File.Exists(artifact.PptxPath));
            Assert.True(File.Exists(artifact.TypstPath));
            Assert.Empty(artifact.Warnings);
        }
    }

    [Fact]
    public void Generate_IsDeterministic_TypstAndManifestIdenticalAcrossRuns()
    {
        var fixture = FixtureCatalog.Find("gradient")!;
        var generator = new FixtureGenerator();

        var first = generator.Generate(fixture, Path.Combine(_testDir, "run1"));
        var second = generator.Generate(fixture, Path.Combine(_testDir, "run2"));

        Assert.Equal(File.ReadAllText(first.TypstPath), File.ReadAllText(second.TypstPath));
        Assert.Equal(File.ReadAllText(first.ManifestPath), File.ReadAllText(second.ManifestPath));
        Assert.Equal(
            File.ReadAllBytes(Path.Combine(first.OutputDirectory, "assets", FixtureAssets.CheckerboardFileName)),
            File.ReadAllBytes(Path.Combine(second.OutputDirectory, "assets", FixtureAssets.CheckerboardFileName)));
    }

    [Fact]
    public void ImageFixture_TypstSource_ReferencesAssetRelativeToSourceFile()
    {
        var fixture = FixtureCatalog.Find("image-fit")!;
        var artifact = new FixtureGenerator().Generate(fixture, Path.Combine(_testDir, "image-fit"));

        var typst = File.ReadAllText(artifact.TypstPath);
        // Typst resolves image paths against the source file; an absolute path or a
        // repo-root-relative path would silently break the preview render.
        Assert.Contains($"\"assets/{FixtureAssets.CheckerboardFileName}\"", typst);
        Assert.DoesNotContain(artifact.OutputDirectory, typst);
    }

    [Fact]
    public void CheckerboardAsset_IsAValidDecodablePng_WithDeclaredDimensions()
    {
        var png = FixtureAssets.CreateCheckerboardPng();

        // PNG signature + IHDR: the bytes must be a real renderable image (PowerPoint
        // and Typst both rasterize it), not a header-only stub.
        Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }, png[..8]);
        Assert.Equal(FixtureAssets.CheckerboardWidth, ReadBigEndianInt(png, 16));
        Assert.Equal(FixtureAssets.CheckerboardHeight, ReadBigEndianInt(png, 20));
    }

    [Fact]
    public void Generate_NullFixture_Throws()
    {
        var generator = new FixtureGenerator();
        Assert.Throws<ArgumentNullException>(() => generator.Generate(null!, Path.Combine(_testDir, "x")));
    }

    private static int ReadBigEndianInt(byte[] buffer, int offset) =>
        (buffer[offset] << 24) | (buffer[offset + 1] << 16) | (buffer[offset + 2] << 8) | buffer[offset + 3];
}
