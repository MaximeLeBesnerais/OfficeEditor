namespace TypstBridge.Managed.Models;

/// <summary>
/// Managed copy of a native Typst bridge compile result.
/// </summary>
public sealed record TypstCompileResult(
    uint Status,
    string Message,
    IReadOnlyList<TypstOutputFile> Outputs,
    IReadOnlyList<TypstDiagnostic> Diagnostics)
{
    public bool Success => Status == 0;
}
