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

    /// <summary>Document-wide allocator for DrawingML non-visual ids.</summary>
    public required DrawingIdAllocator DrawingIds { get; init; }

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
    /// Resolves a style reference for a specific style <paramref name="kind"/>. Existing styles
    /// are validated against the expected kind and referenced by ID (never mutated); an
    /// unknown or wrong-kind reference warns and uses the requested baseline fallback, which is
    /// generated lazily with a collision-free ID when absent.
    /// </summary>
    public string? ResolveStyle(
        string? styleId,
        StyleValues kind,
        string path,
        BaselineStyleKind? fallbackKind = null)
    {
        var fallback = fallbackKind;
        if (fallback is null && kind == StyleValues.Paragraph)
        {
            fallback = BaselineStyleKind.Body;
        }
        else if (fallback is null && kind == StyleValues.Table)
        {
            fallback = BaselineStyleKind.Table;
        }

        return StyleManager.ResolveStyleReference(styleId, kind, path, fallback);
    }

    /// <summary>
    /// Allocates a fresh, collision-free numbering instance for one list. Each list gets its
    /// own instance so counters restart (and a custom <paramref name="start"/> takes effect);
    /// existing template numbering is preserved untouched.
    /// </summary>
    public int AllocateNumbering(bool ordered, int start) =>
        _numberingAllocator.Allocate(MainPart, ordered, start);

    /// <summary>Allocates an id unique across all drawing-bearing package parts.</summary>
    public uint NextDrawingId() => DrawingIds.Next();

    /// <summary>Adds a warning for a resolved block/cell path.</summary>
    public void Warn(string path, string message) =>
        Warnings.Add(new DocxGenerationIssue(path, message, null, DocxGenerationIssueSeverity.Warning));
}

/// <summary>
/// Monotonic DrawingML non-visual id allocator initialized from every existing XML package
/// part. A single instance is shared by body, header/footer and positioned emitters.
/// </summary>
internal sealed class DrawingIdAllocator
{
    private ulong _nextId;

    public DrawingIdAllocator(WordprocessingDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);

        uint maxId = 0;
        HashSet<OpenXmlPart> visited = [];
        foreach (var part in EnumerateParts(document, visited))
        {
            if (part.RootElement is not { } root)
            {
                continue;
            }

            foreach (var element in root.Descendants())
            {
                if (element.LocalName is not ("docPr" or "cNvPr"))
                {
                    continue;
                }

                var id = element.GetAttribute("id", string.Empty).Value;
                if (uint.TryParse(id, System.Globalization.CultureInfo.InvariantCulture, out var parsed))
                {
                    maxId = Math.Max(maxId, parsed);
                }
            }
        }

        _nextId = (ulong)maxId + 1;
    }

    public uint Next()
    {
        if (_nextId > uint.MaxValue)
        {
            throw new InvalidOperationException("No DrawingML non-visual ids remain available in the document.");
        }

        return (uint)_nextId++;
    }

    private static IEnumerable<OpenXmlPart> EnumerateParts(
        OpenXmlPartContainer container,
        HashSet<OpenXmlPart> visited)
    {
        foreach (var pair in container.Parts)
        {
            var part = pair.OpenXmlPart;
            if (!visited.Add(part))
            {
                continue;
            }

            yield return part;
            foreach (var descendant in EnumerateParts(part, visited))
            {
                yield return descendant;
            }
        }
    }
}
