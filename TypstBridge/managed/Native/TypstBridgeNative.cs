using System.Runtime.InteropServices;

namespace TypstBridge.Managed.Native;

internal static class TypstBridgeNative
{
    internal const string LibraryName = "typst_bridge";

    [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl, EntryPoint = "typst_bridge_abi_version")]
    [DefaultDllImportSearchPaths(DllImportSearchPath.AssemblyDirectory | DllImportSearchPath.ApplicationDirectory | DllImportSearchPath.SafeDirectories)]
    internal static extern uint AbiVersion();

    [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl, EntryPoint = "typst_bridge_version")]
    [DefaultDllImportSearchPaths(DllImportSearchPath.AssemblyDirectory | DllImportSearchPath.ApplicationDirectory | DllImportSearchPath.SafeDirectories)]
    internal static extern IntPtr Version();

    [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl, EntryPoint = "typst_bridge_probe")]
    [DefaultDllImportSearchPaths(DllImportSearchPath.AssemblyDirectory | DllImportSearchPath.ApplicationDirectory | DllImportSearchPath.SafeDirectories)]
    internal static extern TypstBridgeStatus Probe();

    [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl, EntryPoint = "typst_bridge_compile")]
    [DefaultDllImportSearchPaths(DllImportSearchPath.AssemblyDirectory | DllImportSearchPath.ApplicationDirectory | DllImportSearchPath.SafeDirectories)]
    internal static extern IntPtr Compile(ref NativeCompileRequest request);

    [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl, EntryPoint = "typst_bridge_free_result")]
    [DefaultDllImportSearchPaths(DllImportSearchPath.AssemblyDirectory | DllImportSearchPath.ApplicationDirectory | DllImportSearchPath.SafeDirectories)]
    internal static extern void FreeResult(IntPtr result);

    [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl, EntryPoint = "typst_bridge_free_string")]
    [DefaultDllImportSearchPaths(DllImportSearchPath.AssemblyDirectory | DllImportSearchPath.ApplicationDirectory | DllImportSearchPath.SafeDirectories)]
    internal static extern void FreeString(IntPtr value);

    [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl, EntryPoint = "typst_bridge_last_error_message")]
    [DefaultDllImportSearchPaths(DllImportSearchPath.AssemblyDirectory | DllImportSearchPath.ApplicationDirectory | DllImportSearchPath.SafeDirectories)]
    internal static extern IntPtr LastErrorMessage();
}

internal enum TypstBridgeStatus : uint
{
    Ok = 0,
    InvalidArgument = 1,
    Compile = 2,
    Render = 3,
    Io = 4,
    Panic = 5,
    Unsupported = 6,
    Internal = 255
}

[StructLayout(LayoutKind.Sequential)]
internal struct NativeCompileRequest
{
    public uint AbiVersion;
    public IntPtr SourceUtf8;
    public UIntPtr SourceLen;
    public IntPtr WorkingDirUtf8;
    public UIntPtr WorkingDirLen;
    public IntPtr RootFileNameUtf8;
    public UIntPtr RootFileNameLen;
    public IntPtr FontPaths;
    public UIntPtr FontPathsCount;
    public uint OutputFormat;
    public double Ppi;
    public uint Flags;
}

[StructLayout(LayoutKind.Sequential)]
internal struct NativeString
{
    public IntPtr ValueUtf8;
    public UIntPtr ValueLen;
}

[StructLayout(LayoutKind.Sequential)]
internal struct NativeCompileResult
{
    public TypstBridgeStatus Status;
    public IntPtr Outputs;
    public UIntPtr OutputsCount;
    public IntPtr Diagnostics;
    public UIntPtr DiagnosticsCount;
    public IntPtr MessageUtf8;
    public UIntPtr MessageLen;
}

[StructLayout(LayoutKind.Sequential)]
internal struct NativeOutputItem
{
    public uint PageIndex;
    public IntPtr FileNameUtf8;
    public UIntPtr FileNameLen;
    public IntPtr Data;
    public UIntPtr DataLen;
}

[StructLayout(LayoutKind.Sequential)]
internal struct NativeDiagnostic
{
    public uint Severity;
    public IntPtr MessageUtf8;
    public UIntPtr MessageLen;
    public IntPtr FileUtf8;
    public UIntPtr FileLen;
    public uint Line;
    public uint Column;
}
