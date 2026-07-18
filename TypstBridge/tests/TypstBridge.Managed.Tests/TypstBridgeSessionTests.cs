using TypstBridge.Managed;
using TypstBridge.Managed.Models;

namespace TypstBridge.Managed.Tests;

/// <summary>
/// ABI v3 persistent-session coverage: lifecycle round-trips, determinism across recompiles
/// and eviction boundaries, edited-source recompiles, and parallel sessions. All tests use
/// the real native library and therefore require the runtime asset to be present.
/// </summary>
public sealed class TypstBridgeSessionTests
{
    private const uint Ok = 0;

    [Fact]
    public void AbiV3IsReportedByManagedWrapperAndNativeBridge()
    {
        TypstBridgeCompiler compiler = CreateAvailableCompiler();

        Assert.Equal(3u, TypstBridgeCompiler.SupportedAbiVersion);
        Assert.Equal(3u, compiler.AbiVersion);
    }

    [Fact]
    public void Session_CreateCompileUpdateRecompileDispose_RoundTrip()
    {
        TypstBridgeCompiler compiler = CreateAvailableCompiler();

        using TypstBridgeCompileSession session = compiler.CreateSession(Environment.CurrentDirectory);
        session.UpdateSource("Hello from a session", rootFileName: "session.typ");

        TypstCompileResult first = session.Compile();

        Assert.Equal(Ok, first.Status);
        TypstOutputFile output = Assert.Single(first.Outputs);
        Assert.Equal("session.pdf", output.FileName);
        Assert.NotEmpty(output.Data);
        Assert.Equal("%PDF"u8.ToArray(), output.Data[..4]);

        // Recompiling the unchanged source must be byte-deterministic.
        TypstCompileResult second = session.Compile();
        Assert.Equal(Ok, second.Status);
        Assert.Equal(output.Data, Assert.Single(second.Outputs).Data);

        // The warm session compile must match the cold single-shot compile byte-for-byte.
        TypstCompileResult cold = compiler.Compile(new TypstCompileRequest(
            source: "Hello from a session",
            workingDirectory: Environment.CurrentDirectory,
            rootFileName: "session.typ"));
        Assert.Equal(Ok, cold.Status);
        Assert.Equal(output.Data, Assert.Single(cold.Outputs).Data);
    }

    [Fact]
    public void Session_WithEditedSource_RecompilesCorrectly()
    {
        TypstBridgeCompiler compiler = CreateAvailableCompiler();

        using TypstBridgeCompileSession session = compiler.CreateSession(Environment.CurrentDirectory);
        session.UpdateSource("Original page");
        TypstCompileResult original = session.Compile();
        Assert.Equal(Ok, original.Status);

        session.UpdateSource("Edited page\n#pagebreak()\nSecond page");
        TypstCompileResult edited = session.Compile(TypstOutputFormat.Svg);

        Assert.Equal(Ok, edited.Status);
        Assert.Equal(2, edited.Outputs.Count);

        // The session compile of the edited source equals a cold compile of the same source.
        TypstCompileResult cold = compiler.Compile(new TypstCompileRequest(
            source: "Edited page\n#pagebreak()\nSecond page",
            workingDirectory: Environment.CurrentDirectory,
            outputFormat: TypstOutputFormat.Svg));
        Assert.Equal(Ok, cold.Status);
        Assert.Equal(cold.Outputs.Count, edited.Outputs.Count);
        for (int i = 0; i < cold.Outputs.Count; i++)
        {
            Assert.Equal(cold.Outputs[i].Data, edited.Outputs[i].Data);
        }
    }

    [Fact]
    public void Session_CompileBeforeUpdateSource_ReturnsInvalidArgument()
    {
        TypstBridgeCompiler compiler = CreateAvailableCompiler();

        using TypstBridgeCompileSession session = compiler.CreateSession(Environment.CurrentDirectory);
        TypstCompileResult result = session.Compile();

        Assert.NotEqual(Ok, result.Status);
        Assert.False(result.Success);
    }

    [Fact]
    public void Session_AfterDispose_ThrowsObjectDisposedException()
    {
        TypstBridgeCompiler compiler = CreateAvailableCompiler();

        TypstBridgeCompileSession session = compiler.CreateSession(Environment.CurrentDirectory);
        session.UpdateSource("Hello");
        session.Dispose();
        session.Dispose(); // double dispose is a no-op

        Assert.Throws<ObjectDisposedException>(() => session.UpdateSource("again"));
        Assert.Throws<ObjectDisposedException>(() => session.Compile());
    }

    [Fact]
    public void ParallelSessions_CompileIndependently()
    {
        TypstBridgeCompiler compiler = CreateAvailableCompiler();
        string[] sources =
        [
            "Single page deck",
            "Two pages\n#pagebreak()\nSecond",
            "Three pages\n#pagebreak()\nSecond\n#pagebreak()\nThird",
            "Another single page deck",
        ];

        var results = new TypstCompileResult[sources.Length];
        Parallel.For(0, sources.Length, i =>
        {
            using TypstBridgeCompileSession session = compiler.CreateSession(Environment.CurrentDirectory);
            session.UpdateSource(sources[i]);
            results[i] = session.Compile(TypstOutputFormat.Svg);
        });

        for (int i = 0; i < sources.Length; i++)
        {
            Assert.Equal(Ok, results[i].Status);
            int expectedPages = sources[i].Split("#pagebreak()").Length;
            Assert.Equal(expectedPages, results[i].Outputs.Count);
        }
    }

    [Fact]
    public void CompilesAcrossEvictionBoundaries_SucceedAndStayDeterministic()
    {
        TypstBridgeCompiler compiler = CreateAvailableCompiler();

        using TypstBridgeCompileSession session = compiler.CreateSession(Environment.CurrentDirectory);
        session.UpdateSource("Eviction boundary determinism");

        TypstCompileResult before = session.Compile();
        Assert.Equal(Ok, before.Status);

        // Force a full memoization eviction, then recompile: same bytes must come out.
        compiler.EvictCache(0);

        TypstCompileResult after = session.Compile();
        Assert.Equal(Ok, after.Status);
        Assert.Equal(Assert.Single(before.Outputs).Data, Assert.Single(after.Outputs).Data);

        // The cold single-shot path is equally unaffected by eviction boundaries.
        compiler.EvictCache(0);
        TypstCompileResult cold = compiler.Compile(new TypstCompileRequest(
            source: "Eviction boundary determinism",
            workingDirectory: Environment.CurrentDirectory));
        Assert.Equal(Ok, cold.Status);
        Assert.Equal(Assert.Single(before.Outputs).Data, Assert.Single(cold.Outputs).Data);
    }

    [Fact]
    public void Session_WithInvalidFontPath_ThrowsBridgeException()
    {
        TypstBridgeCompiler compiler = CreateAvailableCompiler();

        var exception = Assert.Throws<TypstBridgeException>(
            () => compiler.CreateSession(Environment.CurrentDirectory, fontPaths: ["/definitely/missing/font.ttf"]));

        Assert.Contains("font", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    private static TypstBridgeCompiler CreateAvailableCompiler()
    {
        TypstBridgeCompiler compiler = new();
        if (!compiler.Probe())
        {
            throw new InvalidOperationException(
                "TypstBridge native library is not available. Build runtime assets first, e.g. TypstBridge/packaging/build-native.sh linux-x64.");
        }

        return compiler;
    }
}
