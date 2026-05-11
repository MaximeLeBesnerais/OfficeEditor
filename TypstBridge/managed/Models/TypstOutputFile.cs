namespace TypstBridge.Managed.Models;

/// <summary>
/// One output item produced by a compile request.
/// </summary>
public sealed record TypstOutputFile(uint PageIndex, string FileName, byte[] Data);
