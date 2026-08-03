using PptxEditor.Core.Generation.Emit.Typst;
using PptxEditor.Core.Generation.Layout;
using PptxEditor.Core.Generation.Schema;

namespace DocxEditor.Tests.Generation.Emit.Typst;

/// <summary>
/// Golden-file tests for the Typst emitter: a generation document in →
/// Typst source out, compared against committed snapshots. These deterministic source
/// snapshots carry the acceptance weight: actual Typst compilation is covered by
/// <see cref="TypstEmitterCompileTests"/> and is opt-in (no typst CLI / TypstBridge
/// native library in sandbox environments). Set OE_UPDATE_SNAPSHOTS=1 to regenerate
/// the expected files (review the diff before committing).
/// </summary>
public sealed class TypstEmitterGoldenTests
{
    private const string UpdateSnapshotsEnvVar = "OE_UPDATE_SNAPSHOTS";

    [Theory]
    [InlineData("rect")]
    [InlineData("gradient")]
    [InlineData("text")]
    [InlineData("line")]
    [InlineData("ellipse")]
    [InlineData("image")]
    [InlineData("group")]
    [InlineData("container")]
    [InlineData("document")]
    public void Emit_GoldenFixture_MatchesExpectedTypstSource(string name)
    {
        var inputPath = Path.Combine(ResolveGoldenDirectory(), name + ".input.json");
        Assert.True(File.Exists(inputPath), $"Golden input not found: {inputPath}");

        var document = new GenerationDocumentParser().Parse(File.ReadAllText(inputPath));
        var result = new LayoutResolver().Resolve(document);
        var actual = new TypstEmitter().Emit(result);

        var expectedPath = Path.Combine(ResolveGoldenDirectory(), name + ".expected.typ");
        if (Environment.GetEnvironmentVariable(UpdateSnapshotsEnvVar) == "1")
        {
            File.WriteAllText(expectedPath, actual);
        }

        Assert.True(File.Exists(expectedPath),
            $"Golden snapshot not found: {expectedPath}. Run the test with {UpdateSnapshotsEnvVar}=1 to create it.");

        var expected = File.ReadAllText(expectedPath).Replace("\r\n", "\n");
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void Emit_IsDeterministic_RepeatedEmissionProducesIdenticalSource()
    {
        var inputPath = Path.Combine(ResolveGoldenDirectory(), "container.input.json");
        var document = new GenerationDocumentParser().Parse(File.ReadAllText(inputPath));
        var result = new LayoutResolver().Resolve(document);

        var first = new TypstEmitter().Emit(result);
        var second = new TypstEmitter().Emit(result);

        Assert.Equal(first, second);
    }

    private static string ResolveGoldenDirectory()
    {
        var baseDir = AppContext.BaseDirectory;
        return Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "Generation", "Emit", "Typst", "Golden"));
    }
}
