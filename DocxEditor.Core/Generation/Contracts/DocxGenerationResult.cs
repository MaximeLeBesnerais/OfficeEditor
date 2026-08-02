using DocxEditor.Core.Generation.Model;
using DocxEditor.Core.Generation.Schema;

namespace DocxEditor.Core.Generation.Contracts;

/// <summary>
/// Generation result shell: the parsed/emitted <see cref="DocxGenerationDocument"/> plus
/// every artifact an emitter produced (<see cref="Outputs"/>) and any non-fatal findings
/// raised while emitting (<see cref="Warnings"/>). OOXML and Typst emitters return this;
/// downstream CLI/API wiring reports on <see cref="Outputs"/>.
/// </summary>
public sealed record DocxGenerationResult
{
    /// <summary>The document that was generated. Always present.</summary>
    public required DocxGenerationDocument Document { get; init; }

    /// <summary>Artifacts produced by the emitter, in emit order.</summary>
    public IReadOnlyList<EmittedOutput> Outputs { get; init; } = [];

    /// <summary>
    /// Non-fatal findings produced while emitting (e.g. a style referenced but not present
    /// in the template, an inline image emitted as a placeholder because no image resolver
    /// is wired, or positioned-tier elements skipped until the positioned workstream merges).
    /// Empty when emission was clean.
    /// </summary>
    public IReadOnlyList<DocxGenerationIssue> Warnings { get; init; } = [];
}

/// <summary>One artifact produced by an emitter (e.g. the generated .docx package).</summary>
public sealed record EmittedOutput(DocxOutputKind Kind, string FilePath);

/// <summary>The format of an emitted artifact.</summary>
public enum DocxOutputKind
{
    /// <summary>The generated OpenXML document package (.docx).</summary>
    Document,

    /// <summary>A Typst preview source (.typ) used for rendering.</summary>
    TypstPreview
}
