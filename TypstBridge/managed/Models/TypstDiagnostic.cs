namespace TypstBridge.Managed.Models;

/// <summary>
/// Severity for a diagnostic returned by the native Typst bridge.
/// </summary>
public enum TypstDiagnosticSeverity : uint
{
    Error = 1,
    Warning = 2,
    Info = 3
}

/// <summary>
/// Structured compile or render diagnostic reported by Typst.
/// </summary>
public sealed record TypstDiagnostic(
    TypstDiagnosticSeverity Severity,
    string Message,
    string? File,
    uint Line,
    uint Column);
