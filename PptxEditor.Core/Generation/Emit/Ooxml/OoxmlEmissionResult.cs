namespace PptxEditor.Core.Generation.Emit.Ooxml;

/// <summary>Result of an <see cref="OoxmlEmitter"/> run: the .pptx package plus non-fatal warnings.</summary>
public sealed record OoxmlEmissionResult
{
    /// <summary>The generated .pptx bytes.</summary>
    public required byte[] Bytes { get; init; }

    /// <summary>Non-fatal diagnostics (e.g. an image fit that fell back to stretch).</summary>
    public required IReadOnlyList<string> Warnings { get; init; }
}
