using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Contracts;

/// <summary>
/// Contract implemented by downstream emitters (OOXML delivery, Typst preview). The
/// pipeline is <c>parse → model → <see cref="Emit"/> → <see cref="DocxGenerationResult"/></c>.
/// The layout pass is a pure C# walk over the parsed model; emitters translate the
/// resolved tree to their target and record produced artifacts on the result shell.
/// </summary>
public interface IDocxDocumentEmitter
{
    /// <summary>
    /// Emits a fully validated <see cref="DocxGenerationDocument"/> to the target format,
    /// returning the generation result shell carrying the produced outputs. Implementations
    /// must not mutate the input model.
    /// </summary>
    DocxGenerationResult Emit(DocxGenerationDocument document);
}
