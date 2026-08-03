using PptxEditor.Core.Generation.Fixtures;

namespace VisualDiff;

/// <summary>
/// The Phase 5 parity suite (P7): for every fixture in
/// <see cref="FixtureCatalog"/>, diff the PowerPoint ground-truth render
/// (<c>ground-truth/*.png</c>) against the Typst preview render (<c>typst/*.png</c>)
/// in PNG-pair mode. Decks and Typst sources are produced in-process via
/// <see cref="FixtureGenerator"/> (<c>--generate</c>); the actual renders are external
/// and opt-in: PowerPoint for ground truth (manual, steps printed per fixture) and the
/// typst CLI for the preview (<c>--render</c>, probed). Missing renders are loud skips
/// with exact instructions — never a silent pass.
/// </summary>
internal static class GenSuite
{
    /// <summary>Root the fixture artifacts and renders live under (generated, git-ignored).</summary>
    public const string FixturesRoot = "examples/output/gen/parity";

    /// <summary>Committed per-primitive thresholds, applied by default to gen-suite runs.</summary>
    public const string DefaultThresholdsPath = "tools/visual-diff/baselines/gen/thresholds.json";

    public static List<ComparisonInput> BuildInputs(CliOptions options, ToolPaths tools)
    {
        List<ComparisonInput> inputs = [];
        List<string> skipped = [];

        foreach (ParityFixture fixture in FixtureCatalog.All)
        {
            string fixtureDir = Path.Combine(FixturesRoot, fixture.Name);
            string groundTruthDir = Path.Combine(fixtureDir, FixtureGenerator.GroundTruthDirectoryName);
            string previewDir = Path.Combine(fixtureDir, FixtureGenerator.PreviewDirectoryName);

            if (options.GenerateMissing)
            {
                GenerateFixture(fixture, fixtureDir);
            }

            if (options.RenderTypst)
            {
                RenderPreview(fixture, fixtureDir, previewDir, options.Dpi, tools);
            }

            if (!HasPngPages(groundTruthDir))
            {
                skipped.Add(fixture.Name);
                Console.Error.WriteLine($"visual-diff: skipping '{fixture.Name}': ground-truth renders missing: {Path.GetFullPath(groundTruthDir)}");
                Console.Error.WriteLine("  produce them with PowerPoint (ground truth is a manual, opt-in step):");
                Console.Error.WriteLine($"    1. dotnet run --project tools/visual-diff -- --suite gen --generate   # writes {fixtureDir}/fixture.pptx");
                Console.Error.WriteLine($"    2. open {fixtureDir}/fixture.pptx in PowerPoint and export it as PDF (File > Export > PDF) to {fixtureDir}/ground-truth.pdf");
                Console.Error.WriteLine($"    3. pdftocairo -png -r {options.Dpi} {fixtureDir}/ground-truth.pdf {groundTruthDir}/page");
                continue;
            }

            if (!HasPngPages(previewDir))
            {
                skipped.Add(fixture.Name);
                Console.Error.WriteLine($"visual-diff: skipping '{fixture.Name}': Typst preview renders missing: {Path.GetFullPath(previewDir)}");
                Console.Error.WriteLine($"  produce them with: dotnet run --project tools/visual-diff -- --suite gen --render   # needs the typst CLI");
                continue;
            }

            inputs.Add(new(fixture.Name, groundTruthDir, previewDir, InputKind.Png));
        }

        if (inputs.Count == 0)
        {
            throw new InvalidOperationException(
                "No parity fixture renders were available under " + Path.GetFullPath(FixturesRoot) + ". "
                + "The render steps are opt-in by design: run\n"
                + "  dotnet run --project tools/visual-diff -- --suite gen --generate\n"
                + "then produce the two render sets per fixture (PowerPoint ground truth; '--render' for the Typst preview).\n"
                + "See tools/visual-diff/baselines/gen/README.md for the full pipeline.");
        }

        if (skipped.Count > 0)
        {
            Console.Error.WriteLine($"visual-diff: {skipped.Count} fixture(s) skipped: {string.Join(", ", skipped)}");
        }

        return inputs;
    }

    /// <summary>Deterministic in-process generation: fixture.pptx + fixture.typ + manifest.json + assets/.</summary>
    private static void GenerateFixture(ParityFixture fixture, string fixtureDir)
    {
        Console.WriteLine($"Generating fixture '{fixture.Name}' via both emitters...");
        FixtureArtifact artifact = new FixtureGenerator().Generate(fixture, fixtureDir);
        foreach (string warning in artifact.Warnings)
        {
            Console.Error.WriteLine($"visual-diff: fixture '{fixture.Name}' warning: {warning}");
        }
    }

    /// <summary>Opt-in preview render: typst compile --format png, one page-NNN.png per slide.</summary>
    private static void RenderPreview(ParityFixture fixture, string fixtureDir, string previewDir, int dpi, ToolPaths tools)
    {
        if (tools.Typst is null)
        {
            throw new InvalidOperationException(
                "--render needs the typst CLI but it was not found. Install it (macOS: 'brew install typst') or drop --render.");
        }

        string typstSource = Path.Combine(fixtureDir, FixtureGenerator.TypstFileName);
        if (!File.Exists(typstSource))
        {
            throw new InvalidOperationException(
                $"Cannot render '{fixture.Name}': {typstSource} is missing. Run with --generate first to emit the fixture sources.");
        }

        Directory.CreateDirectory(previewDir);
        foreach (string stale in Directory.EnumerateFiles(previewDir, "*.png"))
        {
            File.Delete(stale);
        }

        string outputTemplate = Path.Combine(previewDir, "page-{p}.png");
        Console.WriteLine($"Rendering '{fixture.Name}' preview via typst CLI at {dpi} PPI...");
        // typst resolves relative paths against the input file, so the fixture's
        // assets/checkerboard.png reference works without a --root override.
        ProcessResult result = ComparisonRunner.RunProcess(tools.Typst,
            ["compile", "--format", "png", "--ppi", dpi.ToString(System.Globalization.CultureInfo.InvariantCulture), typstSource, outputTemplate]);
        if (result.ExitCode != 0 || !HasPngPages(previewDir))
        {
            throw new InvalidOperationException(
                $"typst compile failed for fixture '{fixture.Name}' (exit {result.ExitCode}): {result.ErrorOrOutput}");
        }
    }

    private static bool HasPngPages(string directory) =>
        Directory.Exists(directory) && Directory.EnumerateFiles(directory, "*.png", SearchOption.TopDirectoryOnly).Any();
}
