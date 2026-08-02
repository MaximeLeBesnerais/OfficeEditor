namespace DocxEditor.Core.Generation.Model;

/// <summary>
/// Root of the DOCX generation vocabulary: a from-scratch document described by a
/// vocabulary version, optional metadata, optional design tokens, an optional template
/// path and one-or-more sections. Flow content lives in sections (see <see cref="Section"/>);
/// the positioned tier (anchored text boxes, shapes, floating pictures) is section-scoped.
/// This model is distinct from the edit instruction set in <c>DocxEditor.Core.Serialization</c>:
/// generate = new document from JSON, edit = ops against an existing document.
/// </summary>
public sealed record DocxGenerationDocument
{
    /// <summary>The only vocabulary version supported by the parser.</summary>
    public const string SupportedVersion = "1.0";

    /// <summary>Vocabulary version. Only <see cref="SupportedVersion"/> is valid.</summary>
    public required string Version { get; init; }

    /// <summary>Optional document-level metadata (title, author, subject, …).</summary>
    public DocxMetadata? Metadata { get; init; }

    /// <summary>Optional design tokens (palette, fonts, typography, spacing, shapes, page).</summary>
    public DesignTokens? Design { get; init; }

    /// <summary>
    /// Optional path to a template document the emitter opens and reuses for styles and
    /// defaults (style preservation: existing style definitions are never mutated).
    /// </summary>
    public string? TemplatePath { get; init; }

    /// <summary>Sections in document order; at least one is required.</summary>
    public required IReadOnlyList<Section> Sections { get; init; }
}

/// <summary>Optional document-level metadata stored in the package core properties.</summary>
public sealed record DocxMetadata
{
    /// <summary>Document title.</summary>
    public string? Title { get; init; }

    /// <summary>Author name.</summary>
    public string? Author { get; init; }

    /// <summary>Subject line.</summary>
    public string? Subject { get; init; }

    /// <summary>Comma-separated keywords.</summary>
    public string? Keywords { get; init; }

    /// <summary>Free-form description/abstract.</summary>
    public string? Description { get; init; }

    /// <summary>Document language tag (e.g. "en-US").</summary>
    public string? Language { get; init; }
}
