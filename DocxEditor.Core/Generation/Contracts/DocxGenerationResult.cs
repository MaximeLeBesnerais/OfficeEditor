using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Contracts;

/// <summary>
/// Generation result shell: the parsed/emitted <see cref="DocxGenerationDocument"/> plus
/// every artifact an emitter produced (<see cref="Outputs"/>). OOXML and Typst emitters
/// return this; downstream CLI/API wiring reports on <see cref="Outputs"/>.
/// </summary>
public sealed record DocxGenerationResult
{
    /// <summary>The document that was generated. Always present.</summary>
    public required DocxGenerationDocument Document { get; init; }

    /// <summary>Artifacts produced by the emitter, in emit order.</summary>
    public IReadOnlyList<EmittedOutput> Outputs { get; init; } = [];
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
