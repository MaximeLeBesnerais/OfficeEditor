using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Generation.Design;
using OfficeEditor.Core.Exceptions;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Design;

/// <summary>
/// Owns the <c>styles.xml</c> part for generation emission. It loads the template's existing styles
/// once and caches them, references them by ID and <b>never mutates an existing style definition</b>
/// (AGENTS.md style preservation). Missing baseline styles are generated lazily on first request,
/// each with a deterministic, collision-free ID/name (preferred conventional ID first, then an
/// ordinal suffix), and generated IDs are cached so repeated requests return the same ID without
/// duplicating definitions. When the document has no styles part (blank generation), a fresh part is
/// created with Word-convention document defaults so generated content renders correctly.
/// </summary>
public sealed class DocxStyleManager
{
    private readonly WordprocessingDocument _document;
    private readonly DocxDesignResolver _resolver;
    private readonly ResolvedDesign _design;

    private readonly Dictionary<string, Style> _existingById = new(StringComparer.Ordinal);
    private readonly Dictionary<string, string> _existingByName = new(StringComparer.Ordinal);
    private readonly HashSet<string> _takenIds = new(StringComparer.Ordinal);
    private readonly HashSet<string> _takenNames = new(StringComparer.Ordinal);
    private readonly Dictionary<BaselineStyleKind, string> _generated = new();
    private bool _hasDefaultParagraphStyle;
    private StyleDefinitionsPart? _part;

    /// <summary>Opens the styles part of <paramref name="document"/> and loads existing styles once.</summary>
    public DocxStyleManager(WordprocessingDocument document, DocxDesignResolver resolver)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(resolver);
        _document = document;
        _resolver = resolver;
        _design = resolver.ResolveAll();

        var mainPart = document.MainDocumentPart
            ?? throw new OfficeEditorException("The document has no main document part; open or create it before building styles.");
        _part = mainPart.StyleDefinitionsPart;
        if (_part?.Styles is { } styles)
        {
            foreach (var style in styles.Elements<Style>())
            {
                var id = style.StyleId?.Value;
                var name = style.StyleName?.Val?.Value;
                if (id is not null)
                {
                    _existingById[id] = style;
                    _takenIds.Add(id);
                    if (style.Type?.Value == StyleValues.Paragraph && style.Default?.Value == true)
                    {
                        _hasDefaultParagraphStyle = true;
                    }
                }
                if (name is not null)
                {
                    _existingByName.TryAdd(name, id ?? name);
                    _takenNames.Add(name);
                }
            }
        }
    }

    // ---- existing styles ----

    /// <summary>True when the template already defines a style with the given ID.</summary>
    public bool HasExistingStyle(string styleId) => _existingById.ContainsKey(styleId);

    /// <summary>Returns the template style ID for the given ID (same value) when it exists.</summary>
    public string? GetExistingStyleId(string styleId) =>
        _existingById.TryGetValue(styleId, out var style) ? style.StyleId?.Value : null;

    /// <summary>All style IDs present in the template (never mutated).</summary>
    public IReadOnlyCollection<string> ExistingStyleIds => _existingById.Keys;

    /// <summary>The generated (or referenced) style ID for a baseline kind, when it was requested.</summary>
    public string? GetGeneratedStyleId(BaselineStyleKind kind) =>
        _generated.TryGetValue(kind, out var id) ? id : null;

    // ---- baseline styles ----

    /// <summary>Returns the style ID for a baseline kind, referencing the template's definition when
    /// present or generating a missing one (lazily, cached, collision-free).</summary>
    public string GetOrCreateStyle(BaselineStyleKind kind)
    {
        if (_generated.TryGetValue(kind, out var cached))
        {
            return cached;
        }

        var (preferredId, preferredName) = Conventional(kind);
        if (_existingById.ContainsKey(preferredId))
        {
            _generated[kind] = preferredId;
            return preferredId;
        }
        if (_existingByName.TryGetValue(preferredName, out var namedId) && _existingById.ContainsKey(namedId))
        {
            _generated[kind] = namedId;
            return namedId;
        }

        var bodyStyleId = kind == BaselineStyleKind.Body ? string.Empty : GetOrCreateBodyStyle();
        var styleId = AllocateId(preferredId);
        var styleName = AllocateName(preferredName);
        var style = BaselineStyles.Create(kind, _design, styleId, styleName, bodyStyleId);
        if (kind == BaselineStyleKind.Body)
        {
            style.Default = styleId == "Normal" && !_hasDefaultParagraphStyle;
        }
        AppendStyle(style);
        _generated[kind] = styleId;
        return styleId;
    }

    /// <summary>Body/Normal style.</summary>
    public string GetOrCreateBodyStyle() => GetOrCreateStyle(BaselineStyleKind.Body);

    /// <summary>Document title style.</summary>
    public string GetOrCreateTitleStyle() => GetOrCreateStyle(BaselineStyleKind.Title);

    /// <summary>Heading style for a level (1..6).</summary>
    public string GetOrCreateHeadingStyle(int level) =>
        GetOrCreateStyle(BaselineStyleKindExtensions.ForHeading(level));

    /// <summary>Callout/note paragraph style.</summary>
    public string GetOrCreateCalloutStyle() => GetOrCreateStyle(BaselineStyleKind.Callout);

    /// <summary>Table base style.</summary>
    public string GetOrCreateTableStyle() => GetOrCreateStyle(BaselineStyleKind.Table);

    /// <summary>Table header cell paragraph style.</summary>
    public string GetOrCreateTableHeaderStyle() => GetOrCreateStyle(BaselineStyleKind.TableHeader);

    /// <summary>Code block paragraph style.</summary>
    public string GetOrCreateCodeStyle() => GetOrCreateStyle(BaselineStyleKind.Code);

    // ---- author style references ----

    /// <summary>
    /// Resolves an author-supplied style reference (from a block's <c>style</c> field) against the
    /// template's styles by ID or by name. When the reference is unknown it records a deterministic
    /// warning on the resolver and falls back to the requested baseline style (body by default),
    /// or skips the reference when <paramref name="fallbackKind"/> is null. Never generates a style
    /// named after the reference.
    /// </summary>
    public string? ResolveStyleReference(
        string? styleReference,
        string? context = null,
        BaselineStyleKind? fallbackKind = BaselineStyleKind.Body)
    {
        if (string.IsNullOrWhiteSpace(styleReference))
        {
            return fallbackKind is { } kind ? GetOrCreateStyle(kind) : null;
        }
        if (_existingById.ContainsKey(styleReference))
        {
            return styleReference;
        }
        if (_existingByName.TryGetValue(styleReference, out var namedId) && _existingById.ContainsKey(namedId))
        {
            return namedId;
        }
        _resolver.RecordWarning(
            "UnknownStyleReference",
            fallbackKind is { } fallback
                ? $"style reference '{styleReference}' is not defined in the document; falling back to the '{Conventional(fallback).Name}' style."
                : $"style reference '{styleReference}' is not defined in the document; the style reference was skipped.",
            context);
        return fallbackKind is { } fallbackKindValue ? GetOrCreateStyle(fallbackKindValue) : null;
    }

    // ---- internals ----

    private void AppendStyle(Style style)
    {
        var part = EnsureStylesPart();
        part.Styles ??= new Styles();
        part.Styles.Append(style);
        if (style.StyleId?.Value is { } id)
        {
            _takenIds.Add(id);
            _existingById[id] = style;
        }
        if (style.StyleName?.Val?.Value is { } name)
        {
            _takenNames.Add(name);
        }
    }

    private StyleDefinitionsPart EnsureStylesPart()
    {
        if (_part is not null)
        {
            return _part;
        }
        var mainPart = _document.MainDocumentPart
            ?? throw new OfficeEditorException("The document has no main document part; open or create it before building styles.");
        _part = mainPart.AddNewPart<StyleDefinitionsPart>();
        _part.Styles = new Styles();
        _part.Styles.Append(BuildDocumentDefaults());
        return _part;
    }

    /// <summary>
    /// Word-convention document defaults (body font + 11pt) written only when creating a fresh
    /// styles part. An existing template's defaults are never touched.
    /// </summary>
    private DocDefaults BuildDocumentDefaults()
    {
        var font = _design.BodyFontFamilyOrDefault;
        var size = DocxFormattingHelpers.HalfPoints(DocxDesignDefaults.BodyFontSizePt);
        return new DocDefaults(
            new RunPropertiesDefault(
                new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = font, HighAnsi = font, ComplexScript = font },
                    new FontSize { Val = size },
                    new FontSizeComplexScript { Val = size })));
    }

    private string AllocateId(string preferred)
    {
        if (_takenIds.Add(preferred))
        {
            return preferred;
        }
        for (var i = 1; ; i++)
        {
            var candidate = $"{preferred}{i}";
            if (_takenIds.Add(candidate))
            {
                return candidate;
            }
        }
    }

    private string AllocateName(string preferred)
    {
        if (_takenNames.Add(preferred))
        {
            return preferred;
        }
        for (var i = 1; ; i++)
        {
            var candidate = $"{preferred} {i}";
            if (_takenNames.Add(candidate))
            {
                return candidate;
            }
        }
    }

    private static (string Id, string Name) Conventional(BaselineStyleKind kind) => kind switch
    {
        BaselineStyleKind.Body => ("Normal", "Normal"),
        BaselineStyleKind.Title => ("Title", "Title"),
        BaselineStyleKind.Heading1 => ("Heading1", "heading 1"),
        BaselineStyleKind.Heading2 => ("Heading2", "heading 2"),
        BaselineStyleKind.Heading3 => ("Heading3", "heading 3"),
        BaselineStyleKind.Heading4 => ("Heading4", "heading 4"),
        BaselineStyleKind.Heading5 => ("Heading5", "heading 5"),
        BaselineStyleKind.Heading6 => ("Heading6", "heading 6"),
        BaselineStyleKind.Callout => ("Callout", "Callout"),
        BaselineStyleKind.Table => ("TableGrid", "Table Grid"),
        BaselineStyleKind.TableHeader => ("TableHeader", "TableHeader"),
        BaselineStyleKind.Code => ("Code", "Code"),
        _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, "unknown baseline style kind.")
    };
}
