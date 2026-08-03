using System.Text.Json;
using PptxEditor.Core.Generation.Emit.Ooxml;
using PptxEditor.Core.Generation.Emit.Typst;
using PptxEditor.Core.Generation.Layout;

namespace PptxEditor.Core.Generation.Fixtures;

/// <summary>Artifacts produced for one fixture by <see cref="FixtureGenerator"/>.</summary>
public sealed record FixtureArtifact
{
    /// <summary>Fixture name (catalog slug).</summary>
    public required string Name { get; init; }

    /// <summary>Directory the artifacts were written to.</summary>
    public required string OutputDirectory { get; init; }

    /// <summary>The OOXML-emitter deck (delivery side of the parity pair).</summary>
    public required string PptxPath { get; init; }

    /// <summary>The Typst-emitter source (preview side of the parity pair).</summary>
    public required string TypstPath { get; init; }

    /// <summary>Machine-readable manifest describing the fixture and its render instructions.</summary>
    public required string ManifestPath { get; init; }

    /// <summary>Number of slides in the generated deck.</summary>
    public required int SlideCount { get; init; }

    /// <summary>Non-fatal diagnostics from layout and OOXML emission.</summary>
    public required IReadOnlyList<string> Warnings { get; init; }
}

/// <summary>Serialized next to each generated fixture as manifest.json.</summary>
public sealed record FixtureManifest
{
    /// <summary>Fixture name (catalog slug).</summary>
    public required string Name { get; init; }

    /// <summary>What the fixture exercises.</summary>
    public required string Description { get; init; }

    /// <summary>Per-primitive normalized-RMSE ceiling.</summary>
    public required double ThresholdRmse { get; init; }

    /// <summary>Number of slides in the deck.</summary>
    public required int SlideCount { get; init; }

    /// <summary>Non-fatal diagnostics from layout and OOXML emission.</summary>
    public required IReadOnlyList<string> Warnings { get; init; }

    /// <summary>Deck file name (OOXML side), relative to the manifest directory.</summary>
    public required string PptxFile { get; init; }

    /// <summary>Typst source file name (preview side), relative to the manifest directory.</summary>
    public required string TypstFile { get; init; }

    /// <summary>Directory (relative) the PowerPoint ground-truth PNG pages belong in.</summary>
    public required string GroundTruthDirectory { get; init; }

    /// <summary>Directory (relative) the Typst preview PNG pages are rendered into.</summary>
    public required string PreviewDirectory { get; init; }
}

/// <summary>
/// Fixture generator (P7): resolves each catalog fixture once and runs the
/// absolute draw tree through BOTH emitters (rule 2), writing fixture.pptx, fixture.typ
/// and manifest.json per fixture. Deterministic and dependency-free — the PowerPoint and
/// typst render steps that turn these artifacts into diffable PNGs live in
/// tools/visual-diff and are opt-in (see baselines/gen/README.md).
/// <para>
/// The image asset is materialized per fixture under assets/ and referenced two ways:
/// the OOXML pass embeds it from its absolute path, the Typst pass references it
/// relative to the .typ file (Typst resolves image paths against the source file).
/// Both sides see the same pixels, which is what the parity diff measures.
/// </para>
/// </summary>
public sealed class FixtureGenerator
{
    /// <summary>Deck file name written per fixture directory.</summary>
    public const string PptxFileName = "fixture.pptx";

    /// <summary>Typst source file name written per fixture directory.</summary>
    public const string TypstFileName = "fixture.typ";

    /// <summary>Manifest file name written per fixture directory.</summary>
    public const string ManifestFileName = "manifest.json";

    /// <summary>Ground-truth (PowerPoint render) subdirectory name per fixture.</summary>
    public const string GroundTruthDirectoryName = "ground-truth";

    /// <summary>Typst preview render subdirectory name per fixture.</summary>
    public const string PreviewDirectoryName = "typst";

    private const string AssetsDirectoryName = "assets";

    /// <summary>Generates every catalog fixture under <paramref name="outputRoot"/> (one subdirectory per fixture).</summary>
    public IReadOnlyList<FixtureArtifact> GenerateAll(string outputRoot)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputRoot);
        var artifacts = new List<FixtureArtifact>(FixtureCatalog.All.Count);
        foreach (var fixture in FixtureCatalog.All)
        {
            artifacts.Add(Generate(fixture, Path.Combine(outputRoot, fixture.Name)));
        }
        return artifacts;
    }

    /// <summary>Generates one fixture into <paramref name="outputDir"/> (created when missing).</summary>
    public FixtureArtifact Generate(ParityFixture fixture, string outputDir)
    {
        ArgumentNullException.ThrowIfNull(fixture);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputDir);

        Directory.CreateDirectory(outputDir);
        var assetsDir = Path.Combine(outputDir, AssetsDirectoryName);
        Directory.CreateDirectory(assetsDir);
        var assetPath = Path.Combine(assetsDir, FixtureAssets.CheckerboardFileName);
        File.WriteAllBytes(assetPath, FixtureAssets.CreateCheckerboardPng());

        // OOXML pass: the emitter reads the image file (absolute path) and embeds it.
        var ooxmlLayout = new LayoutResolver().Resolve(fixture.BuildDocument(Path.GetFullPath(assetPath)));
        var pptx = new OoxmlEmitter().Emit(ooxmlLayout);

        // Typst pass: the source references the image relative to the .typ file.
        var typstLayout = new LayoutResolver().Resolve(fixture.BuildDocument($"{AssetsDirectoryName}/{FixtureAssets.CheckerboardFileName}"));
        var typstSource = new TypstEmitter().Emit(typstLayout);

        var pptxPath = Path.Combine(outputDir, PptxFileName);
        var typstPath = Path.Combine(outputDir, TypstFileName);
        File.WriteAllBytes(pptxPath, pptx.Bytes);
        File.WriteAllText(typstPath, typstSource);

        var warnings = ooxmlLayout.Warnings
            .Concat(typstLayout.Warnings)
            .Concat(pptx.Warnings)
            .ToList();
        var manifest = new FixtureManifest
        {
            Name = fixture.Name,
            Description = fixture.Description,
            ThresholdRmse = fixture.ThresholdRmse,
            SlideCount = ooxmlLayout.Slides.Count,
            Warnings = warnings,
            PptxFile = PptxFileName,
            TypstFile = TypstFileName,
            GroundTruthDirectory = GroundTruthDirectoryName,
            PreviewDirectory = PreviewDirectoryName
        };
        var manifestPath = Path.Combine(outputDir, ManifestFileName);
        File.WriteAllText(manifestPath, JsonSerializer.Serialize(manifest, new JsonSerializerOptions { WriteIndented = true }));

        return new FixtureArtifact
        {
            Name = fixture.Name,
            OutputDirectory = Path.GetFullPath(outputDir),
            PptxPath = Path.GetFullPath(pptxPath),
            TypstPath = Path.GetFullPath(typstPath),
            ManifestPath = Path.GetFullPath(manifestPath),
            SlideCount = ooxmlLayout.Slides.Count,
            Warnings = warnings
        };
    }
}
