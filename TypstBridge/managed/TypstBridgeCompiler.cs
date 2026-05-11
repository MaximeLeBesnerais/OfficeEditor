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
            uint abiVersion = CallNative(TypstBridgeNative.AbiVersion, "read the Typst bridge ABI version");
            NativeCompileRequest nativeRequest = CreateNativeRequest(request, abiVersion, allocations);
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

    private static NativeCompileRequest CreateNativeRequest(TypstCompileRequest request, uint abiVersion, List<IntPtr> allocations)
    {
        IntPtr source = AllocateUtf8(request.Source, nullTerminated: false, allocations, out int sourceLength);
        IntPtr workingDirectory = AllocateUtf8(request.WorkingDirectory, nullTerminated: false, allocations, out int workingDirectoryLength);
        IntPtr rootFileName = AllocateUtf8(request.RootFileName, nullTerminated: false, allocations, out int rootFileNameLength);
        IntPtr fontPaths = AllocateFontPathArray(request.FontPaths, allocations);

        return new NativeCompileRequest
        {
            AbiVersion = abiVersion,
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

    private static IntPtr AllocateFontPathArray(IReadOnlyList<string> fontPaths, List<IntPtr> allocations)
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

    private static IntPtr AllocateUtf8(string value, bool nullTerminated, List<IntPtr> allocations, out int byteLength)
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

    private static TypstCompileResult CopyResult(IntPtr resultPtr)
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

    private static T CallNative<T>(Func<T> action, string operation)
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

    private static string GetLastErrorMessage()
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
