using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Markdown.Model;

namespace DocxEditor.Core.Markdown;

/// <summary>
/// Resolves a requested style reference (a Word StyleId or a user-defined visible style
/// name, including names with spaces) against an existing document's styles part.
/// Resolution precedence is: exact StyleId, then exact style name, then a unique
/// case-insensitive style name. Ambiguous case-insensitive matches and kind mismatches
/// (paragraph/character/table) are rejected with diagnostics. The resolver is strictly
/// read-only over the template: it never mutates or appends to existing style
/// definitions; unresolved references either yield a generated fallback <see cref="Style"/>
/// of the expected kind (permissive) or an error and no style (strict).
/// </summary>
public sealed class MarkdownStyleResolver
{
    private readonly MarkdownStyleResolverOptions _options;
    private readonly Dictionary<string, Style> _byId;
    private readonly Dictionary<string, List<Style>> _byExactName;
    private readonly Dictionary<string, List<Style>> _byNameCi;
    private readonly HashSet<string> _takenIds;
    private readonly Dictionary<(string Reference, MarkdownStyleKind Kind), Style> _fallbacks = new();

    /// <summary>Builds a resolver over the styles of the given document part.</summary>
    public MarkdownStyleResolver(StyleDefinitionsPart? stylesPart, MarkdownStyleResolverOptions? options = null)
        : this(stylesPart?.Styles, options)
    {
    }

    /// <summary>Builds a resolver over an in-memory <c>Styles</c> root element (tests / detached parts).</summary>
    public MarkdownStyleResolver(Styles? styles, MarkdownStyleResolverOptions? options = null)
    {
        _options = options ?? MarkdownStyleResolverOptions.Default;
        _byId = new Dictionary<string, Style>(StringComparer.Ordinal);
        _byExactName = new Dictionary<string, List<Style>>(StringComparer.Ordinal);
        _byNameCi = new Dictionary<string, List<Style>>(StringComparer.OrdinalIgnoreCase);
        _takenIds = new HashSet<string>(StringComparer.Ordinal);

        if (styles is not null)
        {
            foreach (var style in styles.Elements<Style>())
            {
                var id = style.StyleId?.Value;
                if (id is not null)
                {
                    _byId.TryAdd(id, style);
                    _takenIds.Add(id);
                }

                var name = style.StyleName?.Val?.Value;
                if (name is not null)
                {
                    AddByName(_byExactName, name, style);
                    AddByName(_byNameCi, name, style);
                }
            }
        }
    }

    /// <summary>
    /// Resolves <paramref name="reference"/> for the given expected style kind.
    /// <paramref name="element"/> carries the semantic key (e.g. <c>codeInline</c>) and
    /// <paramref name="path"/> locates the emitting element; both flow into diagnostics.
    /// </summary>
    public MarkdownStyleResolution Resolve(
        string? reference,
        MarkdownStyleKind expectedKind,
        string? element = null,
        string? path = null)
    {
        if (string.IsNullOrWhiteSpace(reference))
        {
            return MarkdownStyleResolution.None(expectedKind);
        }

        var trimmed = reference.Trim();

        // 1. Exact StyleId (style ids are unique by schema, so at most one match).
        if (_byId.TryGetValue(trimmed, out var byId))
        {
            return ResolveMatch(byId, trimmed, expectedKind, element, path);
        }

        // 2. Exact style name. Multiple styles may share a name; that is ambiguous too.
        if (_byExactName.TryGetValue(trimmed, out var exact))
        {
            return exact.Count == 1
                ? ResolveMatch(exact[0], trimmed, expectedKind, element, path)
                : Ambiguous(trimmed, exact.Count, expectedKind, element, path);
        }

        // 3. Unique case-insensitive style name; more than one match is ambiguous.
        if (_byNameCi.TryGetValue(trimmed, out var ci))
        {
            return ci.Count == 1
                ? ResolveMatch(ci[0], trimmed, expectedKind, element, path)
                : Ambiguous(trimmed, ci.Count, expectedKind, element, path);
        }

        return Unresolved(trimmed, expectedKind, element, path);
    }

    /// <summary>
    /// Resolves the style for a semantic markdown key (e.g. <c>codeInline</c>), deriving the
    /// expected kind from <see cref="MarkdownStyleKinds.ForElement"/>.
    /// </summary>
    public MarkdownStyleResolution ResolveElement(string element, string? reference, string? path = null) =>
        Resolve(reference, MarkdownStyleKinds.ForElement(element), element, path);

    private MarkdownStyleResolution ResolveMatch(
        Style style,
        string reference,
        MarkdownStyleKind expectedKind,
        string? element,
        string? path)
    {
        var actualKind = MapKind(style.Type?.Value);
        if (actualKind != expectedKind)
        {
            return Reject(
                reference,
                $"style reference '{reference}' resolves to a {Describe(actualKind)} style, but '{element ?? "element"}' requires a {Describe(expectedKind)} style.",
                element, path, expectedKind);
        }

        return new MarkdownStyleResolution(true, style.StyleId?.Value, expectedKind, null, []);
    }

    private MarkdownStyleResolution Ambiguous(
        string reference,
        int matchCount,
        MarkdownStyleKind expectedKind,
        string? element,
        string? path)
    {
        return Reject(
            reference,
            $"style reference '{reference}' is ambiguous: {matchCount} styles match by name; specify a unique style name or StyleId.",
            element, path, expectedKind);
    }

    private MarkdownStyleResolution Unresolved(
        string reference,
        MarkdownStyleKind expectedKind,
        string? element,
        string? path)
    {
        if (_options.Strict)
        {
            return Reject(
                reference,
                $"style reference '{reference}' is not defined in the document; no style was applied.",
                element, path, expectedKind);
        }

        var fallback = GetOrCreateFallback(reference, expectedKind);
        return new MarkdownStyleResolution(
            false,
            fallback.StyleId?.Value,
            expectedKind,
            fallback,
            [
                new MarkdownStyleDiagnostic(
                    MarkdownDiagnosticSeverity.Warning,
                    $"style reference '{reference}' is not defined in the document; generating a {Describe(expectedKind)} fallback style '{fallback.StyleId?.Value}'.",
                    element, path)
            ]);
    }

    private MarkdownStyleResolution Reject(
        string reference,
        string message,
        string? element,
        string? path,
        MarkdownStyleKind expectedKind)
    {
        var diagnostic = new MarkdownStyleDiagnostic(
            _options.Strict ? MarkdownDiagnosticSeverity.Error : MarkdownDiagnosticSeverity.Warning,
            message,
            element,
            path);
        if (_options.Strict)
        {
            return new MarkdownStyleResolution(false, null, expectedKind, null, [diagnostic]);
        }

        // Permissive: still produce a style so rendering continues, but the diagnostic is kept.
        var fallback = GetOrCreateFallback(reference, expectedKind);
        return new MarkdownStyleResolution(false, fallback.StyleId?.Value, expectedKind, fallback, [diagnostic]);
    }

    private Style GetOrCreateFallback(string reference, MarkdownStyleKind kind)
    {
        if (_fallbacks.TryGetValue((reference, kind), out var cached))
        {
            return cached;
        }

        var id = AllocateId(SanitizeId(reference, kind));
        var name = reference;
        var style = new Style(
            new StyleName { Val = name })
        {
            StyleId = id,
            Type = kind switch
            {
                MarkdownStyleKind.Character => StyleValues.Character,
                MarkdownStyleKind.Table => StyleValues.Table,
                _ => StyleValues.Paragraph
            }
        };

        _fallbacks[(reference, kind)] = style;
        return style;
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

    private static string SanitizeId(string reference, MarkdownStyleKind kind)
    {
        var builder = new StringBuilder(reference.Length);
        foreach (var ch in reference)
        {
            if (char.IsLetterOrDigit(ch))
            {
                builder.Append(ch);
            }
        }

        var id = builder.ToString();
        return id.Length > 0 ? id : $"{kind}Style";
    }

    private static void AddByName(Dictionary<string, List<Style>> map, string name, Style style)
    {
        if (!map.TryGetValue(name, out var list))
        {
            list = [];
            map[name] = list;
        }

        list.Add(style);
    }

    private static MarkdownStyleKind? MapKind(StyleValues? value)
    {
        if (value == StyleValues.Paragraph)
        {
            return MarkdownStyleKind.Paragraph;
        }

        if (value == StyleValues.Character)
        {
            return MarkdownStyleKind.Character;
        }

        if (value == StyleValues.Table)
        {
            return MarkdownStyleKind.Table;
        }

        return null;
    }

    private static string Describe(MarkdownStyleKind? kind) => kind switch
    {
        MarkdownStyleKind.Character => "character",
        MarkdownStyleKind.Table => "table",
        _ => "paragraph"
    };
}
