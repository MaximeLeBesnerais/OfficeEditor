using TypstBridge.Managed;
using TypstBridge.Managed.Models;

namespace TypstBridge.Managed.Tests;

public sealed class TypstBridgeCompilerTests
{
    private const uint Ok = 0;
    private const uint Compile = 2;
    private const uint Unsupported = 6;

    [Fact]
    public void ProbeAndVersionWorkWhenNativeLibraryIsAvailable()
    {
        TypstBridgeCompiler compiler = CreateAvailableCompiler();

        Assert.True(compiler.Probe());
        Assert.Equal(2u, TypstBridgeCompiler.SupportedAbiVersion);
        Assert.Equal(2u, compiler.AbiVersion);
        Assert.False(string.IsNullOrWhiteSpace(compiler.Version));
    }

    [Fact]
    public void MinimalPdfCompileReturnsSinglePdfOutput()
    {
        TypstBridgeCompiler compiler = CreateAvailableCompiler();
        TypstCompileRequest request = new(
            source: "Hello from TypstBridge",
            workingDirectory: Environment.CurrentDirectory,
            rootFileName: "managed-contract.typ");

        TypstCompileResult result = compiler.Compile(request);

        Assert.Equal(Ok, result.Status);
        TypstOutputFile output = Assert.Single(result.Outputs);
        Assert.Equal(0u, output.PageIndex);
        Assert.Equal("managed-contract.pdf", output.FileName);
        Assert.StartsWith("%PDF", System.Text.Encoding.ASCII.GetString(output.Data, 0, 4));
    }

    [Fact]
    public void InvalidTypstSourceReturnsCompileStatusAndDiagnostic()
    {
        TypstBridgeCompiler compiler = CreateAvailableCompiler();
        TypstCompileRequest request = new(
            source: "#let =",
            workingDirectory: Environment.CurrentDirectory);

        TypstCompileResult result = compiler.Compile(request);

        Assert.Equal(Compile, result.Status);
        Assert.False(result.Success);
        Assert.NotEmpty(result.Message);
        Assert.NotEmpty(result.Diagnostics);
        TypstDiagnostic diagnostic = result.Diagnostics[0];
        Assert.Equal(TypstDiagnosticSeverity.Error, diagnostic.Severity);
        Assert.False(string.IsNullOrWhiteSpace(diagnostic.Message));
    }

    [Fact]
    public void SinglePageSvgCompileReturnsSingleSvgOutput()
    {
        TypstBridgeCompiler compiler = CreateAvailableCompiler();
        TypstCompileRequest request = new(
            source: "Hello from TypstBridge",
            workingDirectory: Environment.CurrentDirectory,
            outputFormat: TypstOutputFormat.Svg);

        TypstCompileResult result = compiler.Compile(request);

        Assert.Equal(Ok, result.Status);
        Assert.True(result.Success);
        TypstOutputFile output = Assert.Single(result.Outputs);
        AssertSvgOutput(output, 0u, "page-001.svg");
    }

    [Fact]
    public void MultiPageSvgCompileReturnsOutputsInPageOrder()
    {
        TypstBridgeCompiler compiler = CreateAvailableCompiler();
        TypstCompileRequest request = new(
            source: "First page\n#pagebreak()\nSecond page",
            workingDirectory: Environment.CurrentDirectory,
            outputFormat: TypstOutputFormat.Svg);

        TypstCompileResult result = compiler.Compile(request);

        Assert.Equal(Ok, result.Status);
        Assert.True(result.Success);
        Assert.Equal(2, result.Outputs.Count);

        for (int index = 0; index < result.Outputs.Count; index++)
        {
            uint pageIndex = (uint)index;
            AssertSvgOutput(result.Outputs[index], pageIndex, $"page-{index + 1:000}.svg");
        }
    }

    [Fact]
    public void PngFormatReturnsUnsupportedAfterSuccessfulCompilation()
    {
        TypstBridgeCompiler compiler = CreateAvailableCompiler();
        TypstCompileRequest request = new(
            source: "Hello from TypstBridge",
            workingDirectory: Environment.CurrentDirectory,
            outputFormat: TypstOutputFormat.Png,
            ppi: 96.0);

        TypstCompileResult result = compiler.Compile(request);

        Assert.Equal(Unsupported, result.Status);
        Assert.False(result.Success);
        Assert.Empty(result.Outputs);
        Assert.Contains("not implemented", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    private static void AssertSvgOutput(TypstOutputFile output, uint pageIndex, string fileName)
    {
        Assert.Equal(pageIndex, output.PageIndex);
        Assert.Equal(fileName, output.FileName);
        Assert.EndsWith(".svg", output.FileName, StringComparison.OrdinalIgnoreCase);

        string svg = System.Text.Encoding.UTF8.GetString(output.Data);
        Assert.True(
            svg.Contains("<svg", StringComparison.OrdinalIgnoreCase) ||
            svg.Contains("<?xml", StringComparison.OrdinalIgnoreCase),
            "SVG output should decode as UTF-8 and contain an SVG marker.");
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
