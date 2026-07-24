using TypstBridge.Managed;

namespace TypstBridge.Managed.Tests;

internal static class TestCompilerFactory
{
    internal static TypstBridgeCompiler CreateAvailableCompiler()
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
