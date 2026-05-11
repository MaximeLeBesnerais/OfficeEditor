using System.Reflection;
using System.Runtime.InteropServices;

namespace TypstBridge.Managed.Native;

internal static class NativeLibraryResolver
{
    private static int registered;

    internal static void Register()
    {
        if (Interlocked.Exchange(ref registered, 1) == 1)
        {
            return;
        }

        NativeLibrary.SetDllImportResolver(typeof(TypstBridgeNative).Assembly, Resolve);
    }

    private static IntPtr Resolve(string libraryName, Assembly assembly, DllImportSearchPath? searchPath)
    {
        if (!string.Equals(libraryName, TypstBridgeNative.LibraryName, StringComparison.Ordinal))
        {
            return IntPtr.Zero;
        }

        foreach (string candidate in GetCandidatePaths(assembly))
        {
            if (NativeLibrary.TryLoad(candidate, out IntPtr handle))
            {
                return handle;
            }
        }

        return NativeLibrary.TryLoad(libraryName, assembly, searchPath, out IntPtr defaultHandle)
            ? defaultHandle
            : IntPtr.Zero;
    }

    private static IEnumerable<string> GetCandidatePaths(Assembly assembly)
    {
        string fileName = GetNativeFileName();
        string rid = GetRuntimeIdentifier();
        string baseDirectory = AppContext.BaseDirectory;

        yield return Path.Combine(baseDirectory, fileName);
        yield return Path.Combine(baseDirectory, "runtimes", rid, "native", fileName);

        string? assemblyDirectory = Path.GetDirectoryName(assembly.Location);
        if (!string.IsNullOrWhiteSpace(assemblyDirectory) && !string.Equals(assemblyDirectory, baseDirectory, StringComparison.Ordinal))
        {
            yield return Path.Combine(assemblyDirectory, fileName);
            yield return Path.Combine(assemblyDirectory, "runtimes", rid, "native", fileName);
        }
    }

    private static string GetNativeFileName()
    {
        if (OperatingSystem.IsWindows())
        {
            return "typst_bridge.dll";
        }

        if (OperatingSystem.IsMacOS())
        {
            return "libtypst_bridge.dylib";
        }

        return "libtypst_bridge.so";
    }

    private static string GetRuntimeIdentifier()
    {
        string os = OperatingSystem.IsWindows() ? "win" : OperatingSystem.IsMacOS() ? "osx" : "linux";
        string arch = RuntimeInformation.ProcessArchitecture switch
        {
            Architecture.X64 => "x64",
            Architecture.Arm64 => "arm64",
            Architecture.X86 => "x86",
            Architecture.Arm => "arm",
            _ => RuntimeInformation.ProcessArchitecture.ToString().ToLowerInvariant()
        };

        return $"{os}-{arch}";
    }
}
