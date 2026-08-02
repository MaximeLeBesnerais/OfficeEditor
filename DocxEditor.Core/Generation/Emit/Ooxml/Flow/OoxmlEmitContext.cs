using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Generation.Design;
using DocxEditor.Core.Generation.Emit.Ooxml.Design;
using DocxEditor.Core.Generation.Emit.Ooxml.Positioned;
using DocxEditor.Core.Generation.Model;
using DocxEditor.Core.Generation.Schema;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Flow;

/// <summary>
/// Shared state for a single emit run: the package, the resolved design tokens, the style
/// cache (template styles read-only, blank-document styles created on demand), the fresh
/// numbering allocator and the warning sink. Emitters receive this context and never reach
/// into the package directly beyond what it exposes, so the flow branch stays replaceable.
/// </summary>
internal sealed class OoxmlEmitContext
{
    private readonly Dictionary<OpenXmlCompositeElement, OpenXmlPart> _containerParts = [];
    private readonly NumberingAllocator _numberingAllocator;

    /// <summary>The package being built.</summary>
    public required WordprocessingDocument Document { get; init; }

    /// <summary>The main document part (styles, headers/footers, numbering, images).</summary>
    public required MainDocumentPart MainPart { get; init; }

    /// <summary>Canonical per-document design resolver.</summary>
    public required DocxDesignResolver DesignResolver { get; init; }

    /// <summary>Style cache/manager; existing definitions remain untouched.</summary>
    public required DocxStyleManager StyleManager { get; init; }

    /// <summary>Asset loader + relationship manager shared by inline and positioned images.</summary>
    public required DocxImagePipeline Images { get; init; }

    /// <summary>Positioned drawing emitter configured with this document's design and assets.</summary>
    public required PositionedElementEmitter PositionedEmitter { get; init; }

    /// <summary>True when the package was opened from a template (styles must never be written).</summary>
    public bool FromTemplate { get; init; }

    /// <summary>Warning sink; findings are appended in emit order.</summary>
    public required List<DocxGenerationIssue> Warnings { get; init; }

    public OoxmlEmitContext()
    {
        _numberingAllocator = new NumberingAllocator();
    }

    /// <summary>Associates a header/footer root with the part that owns its relationships.</summary>
    public void RegisterPartContainer(OpenXmlCompositeElement container, OpenXmlPart part)
    {
        _containerParts[container] = part;
    }

    /// <summary>Returns the relationship-owning part for a flow container.</summary>
    public OpenXmlPart GetOwningPart(OpenXmlCompositeElement container) =>
        _containerParts.GetValueOrDefault(container) ?? MainPart;

    /// <summary>
    /// Resolves a style reference. Existing styles are referenced by ID and never mutated;
    /// unknown references warn and use the requested baseline fallback, which is generated
    /// lazily with a collision-free ID when absent.
    /// </summary>
    public string? ResolveStyle(
        string? styleId,
        StyleValues kind,
        string path,
        BaselineStyleKind? fallbackKind = null)
    {
        if (!string.IsNullOrWhiteSpace(styleId) && StyleManager.HasExistingStyle(styleId))
        {
            return styleId;
        }

        var fallback = fallbackKind;
        if (fallback is null && kind == StyleValues.Paragraph)
        {
            fallback = BaselineStyleKind.Body;
        }
        else if (fallback is null && kind == StyleValues.Table)
        {
            fallback = BaselineStyleKind.Table;
        }

        return StyleManager.ResolveStyleReference(styleId, path, fallback);
    }

    /// <summary>
    /// Allocates a fresh, collision-free numbering instance for one list. Each list gets its
    /// own instance so counters restart (and a custom <paramref name="start"/> takes effect);
    /// existing template numbering is preserved untouched.
    /// </summary>
    public int AllocateNumbering(bool ordered, int start) =>
        _numberingAllocator.Allocate(MainPart, ordered, start);

    /// <summary>Adds a warning for a resolved block/cell path.</summary>
    public void Warn(string path, string message) =>
        Warnings.Add(new DocxGenerationIssue(path, message, null, DocxGenerationIssueSeverity.Warning));
}
