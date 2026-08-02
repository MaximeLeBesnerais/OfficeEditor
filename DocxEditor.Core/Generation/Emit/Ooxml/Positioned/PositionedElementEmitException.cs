using OfficeEditor.Core.Exceptions;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Positioned;

/// <summary>
/// Thrown by <see cref="PositionedElementEmitter"/> for invalid input that cannot be
/// degraded: a null/empty target container, an unknown <c>PositionedElement</c> subtype,
/// or non-positive box extents on a primitive that requires them. Model-level geometry
/// problems are normally rejected earlier by the generation validator; this exception is
/// the emitter's own loud contract check (never swallow exceptions — AGENTS.md).
/// </summary>
public sealed class PositionedElementEmitException : OfficeEditorException
{
    public PositionedElementEmitException(string message) : base(message) { }

    public PositionedElementEmitException(string message, Exception innerException) : base(message, innerException) { }
}
