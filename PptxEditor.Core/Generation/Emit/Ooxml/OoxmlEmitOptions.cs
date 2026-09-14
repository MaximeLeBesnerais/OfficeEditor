namespace PptxEditor.Core.Generation.Emit.Ooxml;

/// <summary>Options for <see cref="OoxmlEmitter"/> runs.</summary>
public sealed record OoxmlEmitOptions
{
    /// <summary>
    /// The directory of the document being emitted — for generation documents, the
    /// directory of the JSON deck file. Relative image file sources resolve against
    /// this directory only; there is no repository-root or process-CWD fallback, and
    /// containment (lexical plus, for existing paths, canonical/symlink-aware) keeps
    /// every resolved asset inside it.
    ///
    /// When null (string-only surfaces such as the API/MCP that receive raw JSON without
    /// a base directory), a relative image source is rejected with a loud, actionable
    /// error that suggests a data URI, an absolute path, or an explicit document
    /// directory.
    /// </summary>
    public string? DocumentDirectory { get; init; }
}
