using System.Runtime.InteropServices;
using TypstBridge.Managed.Models;
using TypstBridge.Managed.Native;

namespace TypstBridge.Managed;

/// <summary>
/// Persistent native compile session (ABI v3). Keeps the native world — library, parsed
/// font set, working directory — warm across compiles so recompiling an edited source
/// reuses memoized layout work. Create via <see cref="TypstBridgeCompiler.CreateSession" />,
/// set the source with <see cref="UpdateSource" />, compile with <see cref="Compile" />, and
/// dispose to release the native handle (<c>typst_bridge_session_free</c>).
/// </summary>
/// <remarks>
/// A single session serializes its native calls internally, and distinct sessions compile
/// independently, so parallel preview renders should use one session per deck/thread.
/// </remarks>
public sealed class TypstBridgeCompileSession : IDisposable
{
    private readonly TypstBridgeSessionHandle _handle;
    private int _disposed;

    internal TypstBridgeCompileSession(TypstBridgeSessionHandle handle)
    {
        _handle = handle;
    }

    /// <summary>
    /// Replaces the session's main source in place. The world and font set stay hot, so the
    /// next <see cref="Compile" /> reuses memoized work from previous compiles of this session.
    /// </summary>
    public void UpdateSource(string source, string rootFileName = "main.typ")
    {
        ArgumentNullException.ThrowIfNull(source);
        ArgumentNullException.ThrowIfNull(rootFileName);
        ThrowIfDisposed();

        List<IntPtr> allocations = [];
        try
        {
            IntPtr sourcePtr = TypstBridgeCompiler.AllocateUtf8(source, nullTerminated: false, allocations, out int sourceLength);
            IntPtr rootFileNamePtr = TypstBridgeCompiler.AllocateUtf8(rootFileName, nullTerminated: false, allocations, out int rootFileNameLength);

            TypstBridgeStatus status = TypstBridgeCompiler.CallNative(
                () => TypstBridgeNative.SessionUpdateSource(_handle, sourcePtr, (UIntPtr)sourceLength, rootFileNamePtr, (UIntPtr)rootFileNameLength),
                "update the Typst bridge session source");

            if (status != TypstBridgeStatus.Ok)
            {
                throw new TypstBridgeException($"The native Typst bridge failed to update the session source ({status}). {TypstBridgeCompiler.GetLastErrorMessage()}".Trim());
            }
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
    /// Compiles the session's current source. The result layout is identical to
    /// <see cref="TypstBridgeCompiler.Compile" />: one PDF output or one output per page for
    /// SVG/PNG. All native-owned memory is copied into managed memory before returning.
    /// </summary>
    public TypstCompileResult Compile(TypstOutputFormat outputFormat = TypstOutputFormat.Pdf, double ppi = 144.0)
    {
        ThrowIfDisposed();

        IntPtr resultPtr = TypstBridgeCompiler.CallNative(
            () => TypstBridgeNative.SessionCompile(_handle, (uint)outputFormat, ppi),
            "compile Typst source in a session");

        if (resultPtr == IntPtr.Zero)
        {
            throw new TypstBridgeException($"The native Typst bridge returned a null compile result. {TypstBridgeCompiler.GetLastErrorMessage()}".Trim());
        }

        try
        {
            return TypstBridgeCompiler.CopyResult(resultPtr);
        }
        finally
        {
            TypstBridgeNative.FreeResult(resultPtr);
        }
    }

    /// <summary>
    /// Releases the native session handle. Safe to call multiple times; if omitted, the
    /// handle's finalizer releases the native session instead.
    /// </summary>
    public void Dispose()
    {
        if (Interlocked.Exchange(ref _disposed, 1) == 0)
        {
            _handle.Dispose();
        }
    }

    private void ThrowIfDisposed()
    {
        ObjectDisposedException.ThrowIf(_disposed != 0, this);
    }
}
