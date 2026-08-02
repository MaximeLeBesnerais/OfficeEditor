namespace DocxEditor.Core.Generation.Design;

/// <summary>
/// One deterministic warning recorded while resolving design tokens. <see cref="Code"/> is a
/// stable identifier emitters can branch on (e.g. <c>UnknownColorToken</c>, <c>OffPaletteColor</c>,
/// <c>UndefinedFontSlot</c>, <c>UnknownStyleReference</c>); <see cref="Message"/> is human
/// readable and <see cref="Context"/> optionally names where the token was referenced. Warnings
/// are recorded in resolution order, so for a given document and resolution walk the sequence is
/// deterministic.
/// </summary>
public sealed record DesignResolutionWarning(string Code, string Message, string? Context)
{
    /// <summary>Path/context-prefixed rendering, mirroring the parser's issue format.</summary>
    public override string ToString() =>
        Context is null ? $"[{Code}] {Message}" : $"[{Code}] {Message} (in {Context})";
}
