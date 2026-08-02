using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Generation.Model;
using DocxEditor.Core.Generation.Schema;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Flow;

/// <summary>
/// Default positioned-tier emitter: warns and skips every anchored element. The positioned
/// workstream owns the anchored-emission implementation (anchored drawings with wrap/rotation/
/// z-order); until it merges, positioned content is reported instead of silently dropped.
/// </summary>
internal sealed class UnsupportedPositionedTierEmitter : IPositionedTierEmitter
{
    public static UnsupportedPositionedTierEmitter Instance { get; } = new();

    private UnsupportedPositionedTierEmitter()
    {
    }

    public void EmitPositionedElement(
        PositionedElement element,
        Body body,
        string path,
        IList<DocxGenerationIssue> warnings)
    {
        warnings.Add(new DocxGenerationIssue(
            path,
            $"positioned element '{element.GetType().Name}' was skipped: anchored emission is owned by the positioned workstream and is not available on this branch.",
            null,
            DocxGenerationIssueSeverity.Warning));
    }
}
