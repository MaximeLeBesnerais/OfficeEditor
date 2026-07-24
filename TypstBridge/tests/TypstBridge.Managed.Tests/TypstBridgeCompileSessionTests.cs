using TypstBridge.Managed;
using TypstBridge.Managed.Models;

namespace TypstBridge.Managed.Tests;

/// <summary>
/// Supplementary coverage for <see cref="TypstBridgeCompileSession"/>: PNG round-trip
/// and null-argument validation. The session lifecycle, dispose guards, and parallel
/// sessions are already covered by <see cref="TypstBridgeSessionTests"/>; this file
/// fills the remaining gaps so every code path in the session class is exercised.
/// </summary>
public sealed class TypstBridgeCompileSessionTests
{
    private const uint Ok = 0;

    [Fact]
    public void Session_PngRoundTrip_ReturnsValidPngOutput()
    {
        TypstBridgeCompiler compiler = CreateAvailableCompiler();

        using TypstBridgeCompileSession session = compiler.CreateSession(Environment.CurrentDirectory);
        session.UpdateSource("Hello PNG from session", rootFileName: "session.typ");
        TypstCompileResult result = session.Compile(TypstOutputFormat.Png, ppi: 96.0);

        Assert.Equal(Ok, result.Status);
        Assert.True(result.Success);
        TypstOutputFile output = Assert.Single(result.Outputs);
        Assert.Equal("page-001.png", output.FileName);
        Assert.True(output.Data.Length >= 24,
            "PNG output should include the signature and IHDR dimensions.");
        Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }, output.Data[..8]);
    }

    [Fact]
    public void UpdateSource_NullSource_ThrowsArgumentNullException()
    {
        TypstBridgeCompiler compiler = CreateAvailableCompiler();

        using TypstBridgeCompileSession session = compiler.CreateSession(Environment.CurrentDirectory);
        var exception = Assert.Throws<ArgumentNullException>(() => session.UpdateSource(null!));
        Assert.Equal("source", exception.ParamName);
    }

    [Fact]
    public void UpdateSource_NullRootFileName_ThrowsArgumentNullException()
    {
        TypstBridgeCompiler compiler = CreateAvailableCompiler();

        using TypstBridgeCompileSession session = compiler.CreateSession(Environment.CurrentDirectory);
        var exception = Assert.Throws<ArgumentNullException>(() => session.UpdateSource("hello", null!));
        Assert.Equal("rootFileName", exception.ParamName);
    }

    private static TypstBridgeCompiler CreateAvailableCompiler() =>
        TestCompilerFactory.CreateAvailableCompiler();
}
