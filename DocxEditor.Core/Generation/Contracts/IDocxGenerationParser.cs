using DocxEditor.Core.Generation.Model;
using DocxEditor.Core.Generation.Schema;

namespace DocxEditor.Core.Generation.Contracts;

/// <summary>
/// Stable parser contract for the DOCX generation vocabulary. The OOXML and Typst
/// emitters and any CLI/API wiring consume this interface rather than the concrete
/// parser, so the vocabulary surface is versioned and testable in isolation.
/// </summary>
public interface IDocxGenerationParser
{
    /// <summary>The vocabulary version this parser accepts ("1.0").</summary>
    string SupportedVersion { get; }

    /// <summary>
    /// Validates JSON in a single pass, collecting every error and warning instead of
    /// failing on the first problem. Never throws for contract violations.
    /// </summary>
    DocxGenerationValidationResult Validate(string json);

    /// <summary>
    /// Validates and returns the parsed document, throwing
    /// <see cref="Schema.DocxGenerationValidationException"/> listing all errors when the
    /// JSON is invalid.
    /// </summary>
    DocxGenerationDocument Parse(string json);
}
