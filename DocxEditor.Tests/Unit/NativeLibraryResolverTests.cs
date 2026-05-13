using System.Reflection;
using System.Runtime.InteropServices;
using TypstBridge.Managed;

namespace DocxEditor.Tests.Unit;

public class NativeLibraryResolverTests
{
    private static readonly Type ResolverType = typeof(TypstBridgeCompiler).Assembly.GetType(
        "TypstBridge.Managed.Native.NativeLibraryResolver",
        throwOnError: true)!;

    [Fact]
    public void Register_CalledMultipleTimes_DoesNotThrow()
    {
        MethodInfo register = GetMethod("Register", BindingFlags.Static | BindingFlags.NonPublic);

        Exception? first = Record.Exception(() => register.Invoke(null, null));
        Exception? second = Record.Exception(() => register.Invoke(null, null));

        Assert.Null(first);
        Assert.Null(second);
    }

    [Fact]
    public void GetCandidatePaths_IncludesBaseDirectoryAndRuntimeNativeCandidates()
    {
        MethodInfo getCandidatePaths = GetMethod("GetCandidatePaths", BindingFlags.Static | BindingFlags.NonPublic);
        string fileName = InvokeString("GetNativeFileName");
        string rid = InvokeString("GetRuntimeIdentifier");

        IEnumerable<string> candidates = InvokeCandidates(getCandidatePaths, typeof(TypstBridgeCompiler).Assembly);

        Assert.Contains(Path.Combine(AppContext.BaseDirectory, fileName), candidates);
        Assert.Contains(Path.Combine(AppContext.BaseDirectory, "runtimes", rid, "native", fileName), candidates);
    }

    [Fact]
    public void GetCandidatePaths_IncludesAssemblyDirectoryCandidates_WhenAssemblyLocationDiffersFromBaseDirectory()
    {
        MethodInfo getCandidatePaths = GetMethod("GetCandidatePaths", BindingFlags.Static | BindingFlags.NonPublic);
        string fileName = InvokeString("GetNativeFileName");
        string rid = InvokeString("GetRuntimeIdentifier");
        Assembly frameworkAssembly = typeof(object).Assembly;
        string assemblyDirectory = Path.GetDirectoryName(frameworkAssembly.Location)!;

        IEnumerable<string> candidates = InvokeCandidates(getCandidatePaths, frameworkAssembly);

        Assert.NotEqual(AppContext.BaseDirectory, assemblyDirectory);
        Assert.Contains(Path.Combine(assemblyDirectory, fileName), candidates);
        Assert.Contains(Path.Combine(assemblyDirectory, "runtimes", rid, "native", fileName), candidates);
    }

    [Fact]
    public void GetNativeFileName_ReturnsCurrentPlatformFileName()
    {
        string fileName = InvokeString("GetNativeFileName");

        Assert.Equal(ExpectedNativeFileName(), fileName);
    }

    [Fact]
    public void GetRuntimeIdentifier_ReturnsCurrentPlatformRuntimeIdentifier()
    {
        string rid = InvokeString("GetRuntimeIdentifier");

        Assert.Equal(ExpectedRuntimeIdentifier(), rid);
    }

    [Fact]
    public void Resolve_WithNonMatchingLibraryName_ReturnsZero()
    {
        MethodInfo resolve = GetMethod("Resolve", BindingFlags.Static | BindingFlags.NonPublic);

        object? result = resolve.Invoke(null, ["not_typst_bridge", typeof(TypstBridgeCompiler).Assembly, null]);

        Assert.Equal(IntPtr.Zero, Assert.IsType<IntPtr>(result));
    }

    private static MethodInfo GetMethod(string name, BindingFlags bindingFlags)
    {
        return ResolverType.GetMethod(name, bindingFlags)
            ?? throw new MissingMethodException(ResolverType.FullName, name);
    }

    private static string InvokeString(string methodName)
    {
        MethodInfo method = GetMethod(methodName, BindingFlags.Static | BindingFlags.NonPublic);

        return Assert.IsType<string>(method.Invoke(null, null));
    }

    private static IEnumerable<string> InvokeCandidates(MethodInfo method, Assembly assembly)
    {
        object? result = method.Invoke(null, [assembly]);

        return Assert.IsAssignableFrom<IEnumerable<string>>(result).ToArray();
    }

    private static string ExpectedNativeFileName()
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

    private static string ExpectedRuntimeIdentifier()
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
