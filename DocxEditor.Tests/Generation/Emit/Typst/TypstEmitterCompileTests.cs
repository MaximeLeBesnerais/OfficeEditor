using OfficeEditor.Core.Services;
using PptxEditor.Core.Generation.Emit.Typst;
using PptxEditor.Core.Generation.Layout;
using PptxEditor.Core.Generation.Schema;

namespace DocxEditor.Tests.Generation.Emit.Typst;

/// <summary>
/// Compile smoke test for emitted Typst source via the TypstBridge → CLI chain
/// (TypstBridge-primary backend convention). OPT-IN: set OE_RUN_TYPST_COMPILE_TESTS=1 to enable.
/// Skipped by default because sandboxed environments have neither the typst CLI nor the
/// TypstBridge native library — the deterministic golden snapshots in
/// <see cref="TypstEmitterGoldenTests"/> carry the acceptance weight there.
/// </summary>
public sealed class TypstEmitterCompileTests
{
    private const string EnableEnvVar = "OE_RUN_TYPST_COMPILE_TESTS";

    [Fact]
    public void EmittedSource_AllFixtures_CompilesToSvg()
    {
        if (Environment.GetEnvironmentVariable(EnableEnvVar) != "1")
        {
            return; // no Typst backend in this environment (see class summary)
        }

        var goldenDir = ResolveGoldenDirectory();
        using var compiler = new TypstCompilerService();

        foreach (var name in new[]
                 {
                     "rect", "gradient", "text", "line", "ellipse", "image", "group", "container", "document"
                 })
        {
            var document = new GenerationDocumentParser().Parse(
                File.ReadAllText(Path.Combine(goldenDir, name + ".input.json")));
            var source = new TypstEmitter().Emit(new LayoutResolver().Resolve(document));

            var result = compiler.Compile(source, new CompileOptions
            {
                Format = OutputFormat.Svg,
                WorkingDirectory = goldenDir
            });

            Assert.True(result.Success, $"fixture '{name}' failed to compile: {result.ErrorMessage}");
            Assert.NotEmpty(result.Pages);
        }
    }

    private static string ResolveGoldenDirectory()
    {
        var baseDir = AppContext.BaseDirectory;
        return Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "Generation", "Emit", "Typst", "Golden"));
    }
}
