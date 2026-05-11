namespace TypstBridge.Managed.Models;

/// <summary>
/// Output format requested from the native Typst bridge.
/// </summary>
public enum TypstOutputFormat : uint
{
    Pdf = 1,
    Png = 2,
    Svg = 3
}
