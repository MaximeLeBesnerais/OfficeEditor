namespace DocxEditor.Core.Generation.Emit.Ooxml.Positioned;

/// <summary>
/// Result of a <see cref="PositionedElementEmitter"/> run: how many positioned elements
/// were emitted (anchored into the target), how many were skipped, and any explicit
/// warnings produced while degrading unsupported or approximated combinations. Warnings
/// are always surfaced here — the emitter never silently drops a primitive or a
/// representable property.
/// </summary>
public sealed record PositionedElementEmitResult
{
    /// <summary>Number of elements emitted as OOXML anchors into the target container.</summary>
    public int EmittedCount { get; init; }

    /// <summary>Number of elements that could not be emitted (e.g. an unresolvable image).</summary>
    public int SkippedCount { get; init; }

    /// <summary>Explicit degradation/approximation warnings, in emit order.</summary>
    public IReadOnlyList<PositionedElementEmitWarning> Warnings { get; init; } = [];
}

/// <summary>
/// One explicit warning about an approximated or degraded positioned element. The index
/// is the positioned element's original declaration index, <see cref="ElementType"/>
/// is the JSON type name (textBox/image/rect/line/callout).
/// </summary>
public sealed record PositionedElementEmitWarning(int Index, string ElementType, string Message)
{
    /// <summary>Optional path below the positioned element associated with the warning.</summary>
    public string? PathSuffix { get; init; }

    /// <summary>Formats the warning like the generation validator's path-qualified issues.</summary>
    public override string ToString() =>
        $"positioned[{Index}]{(PathSuffix is null ? string.Empty : $".{PathSuffix}")} ({ElementType}): {Message}";
}
