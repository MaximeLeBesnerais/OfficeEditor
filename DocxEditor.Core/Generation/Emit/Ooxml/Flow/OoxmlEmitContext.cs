using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
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
    private readonly Dictionary<string, Style> _styleCache = new(StringComparer.Ordinal);
    private readonly NumberingAllocator _numberingAllocator;
    private int _nextDrawingId;

    /// <summary>The package being built.</summary>
    public required WordprocessingDocument Document { get; init; }

    /// <summary>The main document part (styles, headers/footers, numbering, images).</summary>
    public required MainDocumentPart MainPart { get; init; }

    /// <summary>Resolved design tokens, or null when the document carries none.</summary>
    public DesignTokens? Design { get; init; }

    /// <summary>True when the package was opened from a template (styles must never be written).</summary>
    public bool FromTemplate { get; init; }

    /// <summary>Optional image-content seam; null = placeholder emission for inline images.</summary>
    public IImageContentResolver? ImageResolver { get; init; }

    /// <summary>Optional positioned-tier seam; null = warn-and-skip for positioned elements.</summary>
    public IPositionedTierEmitter? PositionedTierEmitter { get; init; }

    /// <summary>Warning sink; findings are appended in emit order.</summary>
    public required List<DocxGenerationIssue> Warnings { get; init; }

    public OoxmlEmitContext()
    {
        _numberingAllocator = new NumberingAllocator();
    }

    /// <summary>Loads template styles into the cache (called once, only for templates).</summary>
    public void LoadTemplateStyles()
    {
        if (!FromTemplate)
        {
            return;
        }

        var stylesPart = MainPart.StyleDefinitionsPart;
        if (stylesPart?.Styles is null)
        {
            return;
        }

        foreach (var style in stylesPart.Styles.Elements<Style>())
        {
            if (style.StyleId?.Value is { } id)
            {
                _styleCache[id] = style;
            }
        }
    }

    /// <summary>Registers a style (used for the blank-document default Normal style).</summary>
    public void RegisterStyle(Style style)
    {
        if (style.StyleId?.Value is { } id)
        {
            _styleCache[id] = style;
        }
    }

    /// <summary>
    /// Resolves a style reference. Template styles are referenced by ID and never mutated; an
    /// unknown template style warns and is skipped. Blank documents get a default style of the
    /// requested kind created on demand.
    /// </summary>
    public bool TryEnsureStyle(string styleId, StyleValues kind, string path)
    {
        if (_styleCache.ContainsKey(styleId))
        {
            return true;
        }

        if (FromTemplate)
        {
            Warnings.Add(new DocxGenerationIssue(
                path,
                $"style '{styleId}' is not defined in the template; the style reference was skipped (existing template styles are never mutated).",
                null,
                DocxGenerationIssueSeverity.Warning));
            return false;
        }

        var stylesPart = MainPart.StyleDefinitionsPart ?? MainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles ??= new Styles();
        var style = FormattingHelpers.BuildDefaultStyle(styleId, kind);
        stylesPart.Styles.Append(style);
        _styleCache[styleId] = style;
        return true;
    }

    /// <summary>
    /// Allocates a fresh, collision-free numbering instance for one list. Each list gets its
    /// own instance so counters restart (and a custom <paramref name="start"/> takes effect);
    /// existing template numbering is preserved untouched.
    /// </summary>
    public int AllocateNumbering(bool ordered, int start) =>
        _numberingAllocator.Allocate(MainPart, ordered, start);

    /// <summary>Allocates a fresh drawing id (unique within this emit run).</summary>
    public int NextDrawingId() => _nextDrawingId++;

    /// <summary>Adds a warning for a resolved block/cell path.</summary>
    public void Warn(string path, string message) =>
        Warnings.Add(new DocxGenerationIssue(path, message, null, DocxGenerationIssueSeverity.Warning));
}
