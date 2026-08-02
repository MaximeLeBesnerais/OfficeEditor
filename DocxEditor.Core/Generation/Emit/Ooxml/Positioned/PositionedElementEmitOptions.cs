using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Positioned;

/// <summary>
/// Options for the positioned-tier OOXML emitter (<see cref="PositionedElementEmitter"/>).
/// All properties are optional: the emitter runs with built-in fallbacks when they are
/// omitted (blank design, no image parts). The design tokens feed color/font/typography
/// resolution for text, fills and strokes; the image resolver is the seam that lets the
/// integrating pipeline (owned by another branch) supply image relationship ids and
/// natural sizes without this emitter ever touching asset bytes.
/// </summary>
public sealed record PositionedElementEmitOptions
{
    /// <summary>
    /// Resolved design tokens (palette, fonts, typography, spacing, shapes, page). When
    /// null, token/typography references resolve to raw values and shape defaults fall
    /// back to built-ins (no fill, no stroke, square corners).
    /// </summary>
    public DesignTokens? Design { get; init; }

    /// <summary>
    /// Seam to an image-part manager owned by the integrating pipeline. When null,
    /// positioned images degrade to a warning and are skipped (the emitter never loads
    /// asset paths or data URIs itself — see <see cref="IPositionedImageResolver"/>).
    /// </summary>
    public IPositionedImageResolver? ImageResolver { get; init; }

    /// <summary>
    /// Optional explicit base for drawing <c>wp:docPr</c>/<c>a:cNvPr</c> ids. When null
    /// the emitter scans the target part for existing drawing ids and starts above the
    /// highest one, guaranteeing uniqueness within the part.
    /// </summary>
    public uint? DrawingIdBase { get; init; }

    /// <summary>Integrating emitter's document-wide id allocator.</summary>
    internal Func<uint>? DrawingIdAllocator { get; init; }
}
