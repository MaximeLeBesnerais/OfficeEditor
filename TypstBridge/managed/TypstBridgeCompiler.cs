using System.Runtime.InteropServices;
using System.Text;
using TypstBridge.Managed.Models;
using TypstBridge.Managed.Native;

namespace TypstBridge.Managed;

/// <summary>
/// Managed entry point for compiling Typst source through the first-party native bridge.
/// </summary>
public sealed class TypstBridgeCompiler
{
    /// <summary>
    /// ABI version supported by this managed wrapper.
    /// </summary>
    public const uint SupportedAbiVersion = 3;

    static TypstBridgeCompiler()
    {
        NativeLibraryResolver.Register();
    }

    /// <summary>
    /// ABI version reported by the native bridge. Throws if the native library cannot be loaded.
    /// </summary>
    public uint AbiVersion => CallNative(TypstBridgeNative.AbiVersion, "read the Typst bridge ABI version");

    /// <summary>
    /// Native bridge version string. Throws if the native library cannot be loaded.
    /// </summary>
    public string Version
    {
        get
        {
            IntPtr value = CallNative(TypstBridgeNative.Version, "read the Typst bridge version");
            return value == IntPtr.Zero ? string.Empty : Marshal.PtrToStringUTF8(value) ?? string.Empty;
        }
    }

    /// <summary>
    /// Returns <see langword="true" /> when the native bridge can be loaded and reports a successful probe.
    /// Missing native libraries are treated as an unavailable bridge, not as an exception.
    /// </summary>
    public bool Probe()
    {
        try
        {
            return TypstBridgeNative.Probe() == TypstBridgeStatus.Ok;
        }
        catch (DllNotFoundException)
        {
            return false;
        }
        catch (EntryPointNotFoundException)
        {
            return false;
        }
        catch (BadImageFormatException)
        {
            return false;
        }
    }

    /// <summary>
    /// Compiles a Typst source string using the native bridge and copies all native-owned output into managed memory.
    /// </summary>
    public TypstCompileResult Compile(TypstCompileRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);

        List<IntPtr> allocations = [];
        IntPtr resultPtr = IntPtr.Zero;

        try
        {
            EnsureSupportedAbiVersion();
            NativeCompileRequest nativeRequest = CreateNativeRequest(request, allocations);
            resultPtr = CallNative(() => TypstBridgeNative.Compile(ref nativeRequest), "compile Typst source");

            if (resultPtr == IntPtr.Zero)
            {
                throw new TypstBridgeException($"The native Typst bridge returned a null compile result. {GetLastErrorMessage()}".Trim());
            }

            return CopyResult(resultPtr);
        }
        finally
        {
            if (resultPtr != IntPtr.Zero)
            {
                TypstBridgeNative.FreeResult(resultPtr);
            }

            foreach (IntPtr allocation in allocations)
            {
                Marshal.FreeHGlobal(allocation);
            }
        }
    }

    private static void EnsureSupportedAbiVersion()
    {
        uint nativeAbiVersion = CallNative(TypstBridgeNative.AbiVersion, "read the Typst bridge ABI version");
        if (nativeAbiVersion != SupportedAbiVersion)
        {
            throw new TypstBridgeException(
                $"Unsupported Typst bridge ABI version {nativeAbiVersion}. This managed wrapper supports ABI version {SupportedAbiVersion}.");
        }
    }

    /// <summary>
    /// Creates a persistent compile session that keeps the native world (library, font set,
    /// working directory) warm across compiles. Use
    /// <see cref="TypstBridgeCompileSession.UpdateSource" /> to swap the source in place and
    /// <see cref="TypstBridgeCompileSession.Compile" /> to recompile; dispose the session to
    /// release the native handle.
    /// </summary>
    public TypstBridgeCompileSession CreateSession(string workingDirectory, IReadOnlyList<string>? fontPaths = null)
    {
        ArgumentNullException.ThrowIfNull(workingDirectory);

        List<IntPtr> allocations = [];
        try
        {
            EnsureSupportedAbiVersion();
            IntPtr workingDirectoryPtr = AllocateUtf8(workingDirectory, nullTerminated: false, allocations, out int workingDirectoryLength);
            IntPtr fontPathsPtr = AllocateFontPathArray(fontPaths ?? [], allocations);

            IntPtr rawHandle = IntPtr.Zero;
            TypstBridgeStatus status = CallNative(
                () => TypstBridgeNative.SessionCreate(
                    workingDirectoryPtr,
                    (UIntPtr)workingDirectoryLength,
                    fontPathsPtr,
                    (UIntPtr)(fontPaths?.Count ?? 0),
                    out rawHandle),
                "create a Typst bridge session");

            if (status != TypstBridgeStatus.Ok || rawHandle == IntPtr.Zero)
            {
                throw new TypstBridgeException($"The native Typst bridge failed to create a session ({status}). {GetLastErrorMessage()}".Trim());
            }

            return new TypstBridgeCompileSession(TypstBridgeSessionHandle.FromRawHandle(rawHandle));
        }
        finally
        {
            foreach (IntPtr allocation in allocations)
            {
                Marshal.FreeHGlobal(allocation);
            }
        }
    }

    /// <summary>
    /// Forces one comemo memoization eviction pass in the native bridge. Automatic eviction
    /// already runs on a cadence (see docs/abi.md); this hook exists for memory-pressure or
    /// idle transitions. <paramref name="maxAge" /> of 0 evicts every unreferenced entry.
    /// </summary>
    public void EvictCache(uint maxAge)
    {
        TypstBridgeStatus status = CallNative(() => TypstBridgeNative.EvictCache(maxAge), "evict the Typst bridge memoization caches");
        if (status != TypstBridgeStatus.Ok)
        {
            throw new TypstBridgeException($"The native Typst bridge failed to evict memoization caches ({status}). {GetLastErrorMessage()}".Trim());
        }
    }

    private static NativeCompileRequest CreateNativeRequest(TypstCompileRequest request, List<IntPtr> allocations)
    {
        IntPtr source = AllocateUtf8(request.Source, nullTerminated: false, allocations, out int sourceLength);
        IntPtr workingDirectory = AllocateUtf8(request.WorkingDirectory, nullTerminated: false, allocations, out int workingDirectoryLength);
        IntPtr rootFileName = AllocateUtf8(request.RootFileName, nullTerminated: false, allocations, out int rootFileNameLength);
        IntPtr fontPaths = AllocateFontPathArray(request.FontPaths, allocations);

        return new NativeCompileRequest
        {
            AbiVersion = SupportedAbiVersion,
            SourceUtf8 = source,
            SourceLen = (UIntPtr)sourceLength,
            WorkingDirUtf8 = workingDirectory,
            WorkingDirLen = (UIntPtr)workingDirectoryLength,
            RootFileNameUtf8 = rootFileName,
            RootFileNameLen = (UIntPtr)rootFileNameLength,
            FontPaths = fontPaths,
            FontPathsCount = (UIntPtr)request.FontPaths.Count,
            OutputFormat = (uint)request.OutputFormat,
            Ppi = request.Ppi,
            Flags = 0
        };
    }

    internal static IntPtr AllocateFontPathArray(IReadOnlyList<string> fontPaths, List<IntPtr> allocations)
    {
        if (fontPaths.Count == 0)
        {
            return IntPtr.Zero;
        }

        int itemSize = Marshal.SizeOf<NativeString>();
        IntPtr array = Marshal.AllocHGlobal(checked(itemSize * fontPaths.Count));
        allocations.Add(array);

        for (int i = 0; i < fontPaths.Count; i++)
        {
            IntPtr item = AllocateUtf8(fontPaths[i], nullTerminated: false, allocations, out int byteLength);
            NativeString nativeString = new()
            {
                ValueUtf8 = item,
                ValueLen = (UIntPtr)byteLength
            };
            Marshal.StructureToPtr(nativeString, IntPtr.Add(array, i * itemSize), fDeleteOld: false);
        }

        return array;
    }

    internal static IntPtr AllocateUtf8(string value, bool nullTerminated, List<IntPtr> allocations, out int byteLength)
    {
        byte[] bytes = Encoding.UTF8.GetBytes(value);
        byteLength = bytes.Length;

        int allocationLength = checked(bytes.Length + (nullTerminated ? 1 : 0));
        IntPtr buffer = Marshal.AllocHGlobal(allocationLength);
        allocations.Add(buffer);

        Marshal.Copy(bytes, 0, buffer, bytes.Length);
        if (nullTerminated)
        {
            Marshal.WriteByte(buffer, bytes.Length, 0);
        }

        return buffer;
    }

    internal static TypstCompileResult CopyResult(IntPtr resultPtr)
    {
        NativeCompileResult nativeResult = Marshal.PtrToStructure<NativeCompileResult>(resultPtr);
        TypstOutputFile[] outputs = CopyOutputs(nativeResult.Outputs, ToInt32(nativeResult.OutputsCount));
        TypstDiagnostic[] diagnostics = CopyDiagnostics(nativeResult.Diagnostics, ToInt32(nativeResult.DiagnosticsCount));
        string message = ReadUtf8(nativeResult.MessageUtf8, nativeResult.MessageLen) ?? string.Empty;

        return new TypstCompileResult((uint)nativeResult.Status, message, outputs, diagnostics);
    }

    private static TypstOutputFile[] CopyOutputs(IntPtr outputsPtr, int count)
    {
        if (outputsPtr == IntPtr.Zero || count == 0)
        {
            return [];
        }

        TypstOutputFile[] outputs = new TypstOutputFile[count];
        int itemSize = Marshal.SizeOf<NativeOutputItem>();

        for (int i = 0; i < count; i++)
        {
            IntPtr itemPtr = IntPtr.Add(outputsPtr, i * itemSize);
            NativeOutputItem item = Marshal.PtrToStructure<NativeOutputItem>(itemPtr);
            byte[] data = CopyBytes(item.Data, item.DataLen);
            string fileName = ReadUtf8(item.FileNameUtf8, item.FileNameLen) ?? string.Empty;
            outputs[i] = new TypstOutputFile(item.PageIndex, fileName, data);
        }

        return outputs;
    }

    private static TypstDiagnostic[] CopyDiagnostics(IntPtr diagnosticsPtr, int count)
    {
        if (diagnosticsPtr == IntPtr.Zero || count == 0)
        {
            return [];
        }

        TypstDiagnostic[] diagnostics = new TypstDiagnostic[count];
        int itemSize = Marshal.SizeOf<NativeDiagnostic>();

        for (int i = 0; i < count; i++)
        {
            IntPtr itemPtr = IntPtr.Add(diagnosticsPtr, i * itemSize);
            NativeDiagnostic item = Marshal.PtrToStructure<NativeDiagnostic>(itemPtr);
            string message = ReadUtf8(item.MessageUtf8, item.MessageLen) ?? string.Empty;
            string? file = ReadUtf8(item.FileUtf8, item.FileLen);
            diagnostics[i] = new TypstDiagnostic((TypstDiagnosticSeverity)item.Severity, message, file, item.Line, item.Column);
        }

        return diagnostics;
    }

    private static byte[] CopyBytes(IntPtr dataPtr, UIntPtr dataLength)
    {
        int length = ToInt32(dataLength);
        if (dataPtr == IntPtr.Zero || length == 0)
        {
            return [];
        }

        byte[] data = new byte[length];
        Marshal.Copy(dataPtr, data, 0, length);
        return data;
    }

    private static string? ReadUtf8(IntPtr value, UIntPtr length)
    {
        int byteLength = ToInt32(length);
        if (value == IntPtr.Zero)
        {
            return null;
        }

        return byteLength == 0 ? string.Empty : Marshal.PtrToStringUTF8(value, byteLength);
    }

    private static int ToInt32(UIntPtr value)
    {
        ulong count = value.ToUInt64();
        if (count > int.MaxValue)
        {
            throw new TypstBridgeException($"The native Typst bridge returned a buffer that is too large to copy into managed memory: {count} bytes.");
        }

        return (int)count;
    }

    internal static T CallNative<T>(Func<T> action, string operation)
    {
        try
        {
            return action();
        }
        catch (DllNotFoundException ex)
        {
            throw new TypstBridgeException($"Unable to {operation}: native library '{TypstBridgeNative.LibraryName}' could not be found.", ex);
        }
        catch (EntryPointNotFoundException ex)
        {
            throw new TypstBridgeException($"Unable to {operation}: native library '{TypstBridgeNative.LibraryName}' does not expose the expected TypstBridge ABI.", ex);
        }
        catch (BadImageFormatException ex)
        {
            throw new TypstBridgeException($"Unable to {operation}: native library '{TypstBridgeNative.LibraryName}' is not compatible with this process architecture.", ex);
        }
    }

    internal static string GetLastErrorMessage()
    {
        try
        {
            IntPtr message = TypstBridgeNative.LastErrorMessage();
            return message == IntPtr.Zero ? string.Empty : Marshal.PtrToStringUTF8(message) ?? string.Empty;
        }
        catch (DllNotFoundException)
        {
            return string.Empty;
        }
        catch (EntryPointNotFoundException)
        {
            return string.Empty;
        }
    }
}
