using System.Runtime.InteropServices;

namespace TypstBridge.Managed.Native;

/// <summary>
/// Owns a native TypstBridge session handle. The native side frees the session
/// exactly once, from <see cref="ReleaseHandle" />, so finalization is safe even
/// if the session was never disposed explicitly.
/// </summary>
internal sealed class TypstBridgeSessionHandle : SafeHandle
{
    private TypstBridgeSessionHandle()
        : base(IntPtr.Zero, ownsHandle: true)
    {
    }

    public override bool IsInvalid => handle == IntPtr.Zero;

    internal static TypstBridgeSessionHandle FromRawHandle(IntPtr rawHandle)
    {
        var safeHandle = new TypstBridgeSessionHandle();
        safeHandle.SetHandle(rawHandle);
        return safeHandle;
    }

    protected override bool ReleaseHandle()
    {
        TypstBridgeNative.SessionFree(handle);
        return true;
    }
}
