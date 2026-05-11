namespace TypstBridge.Managed.Models;

/// <summary>
/// Exception thrown when the managed wrapper cannot call the native Typst bridge.
/// </summary>
public sealed class TypstBridgeException : Exception
{
    public TypstBridgeException(string message)
        : base(message)
    {
    }

    public TypstBridgeException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}
