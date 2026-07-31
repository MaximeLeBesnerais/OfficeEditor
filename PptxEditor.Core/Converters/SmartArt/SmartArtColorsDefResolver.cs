using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace PptxEditor.Core.Converters.SmartArt;

/// <summary>
/// Resolves a diagram node's <c>dgm:colorsDef</c> fill for the cached-vs-relayout
/// conflict case (INV-regressions-b1 §5-6). PowerPoint re-lays SmartArt out from
/// the diagram definition parts; when the cached <c>dsp:sp</c> spPr gradFill
/// disagrees with a flat colorsDef mapping, PowerPoint displays the colorsDef
/// result. This resolver reproduces that outcome ONLY in the clear-conflict case:
///
/// A flat colorsDef fill wins over the cached spPr gradFill when ALL of:
///  1. the shape's modelId resolves to a data-part presentation point whose
///     prSet gives presStyleLbl (default "node0") and presStyleIdx (default 0);
///  2. the colorsDef styleLbl exists and its fillClrLst uses the "repeat"
///     method (the default; "span" interpolates → not a plain solid);
///  3. the selected entry (presStyleIdx modulo entry count) is a plain
///     a:schemeClr or a:srgbClr with NO colour-transform children
///     (tint/shade/satMod/lumMod/lumOff/alpha/…);
///  4. the quickStyle's matching styleLbl carries no a:gradFill, a:effectLst or
///     a:effectDag descendant (fillRef/effectRef matrix references do NOT block:
///     the theme gradient they select is exactly the cached gradFill PowerPoint
///     itself wrote, and its relayout still shows the colorsDef colour family).
///
/// In every other situation the cached spPr keeps precedence (null result):
/// cached solidFill (never consulted — gradient only), transformed or missing
/// colorsDef entries, span method, quickStyle gradients/effects, unknown
/// labels/modelIds, or shapes not loaded from a package (no part context).
/// </summary>
internal static class SmartArtColorsDefResolver
{
    private const string DiagramNs = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
    private const string DrawingmlNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// <summary>
    /// Per-diagram context cache, keyed by the dsp:drawing root element. A null
    /// <see cref="ContextHolder.Context"/> caches the negative (no resolvable
    /// parts) so repeated shape extractions do not re-enumerate the package.
    /// </summary>
    private static readonly ConditionalWeakTable<OpenXmlElement, ContextHolder> ContextCache = new();

    /// <summary>
    /// Resolves the flat colorsDef fill for a cached diagram shape. Returns null
    /// when the clear-conflict rule does not apply (cached spPr keeps precedence).
    /// When <paramref name="explicitContext"/> is null, the context is located
    /// from the shape's drawing part (package-loaded shapes only).
    /// </summary>
    internal static string? TryResolveFlatFill(
        OpenXmlElement dspShape,
        string? modelId,
        IReadOnlyDictionary<string, string>? schemeColors,
        ColorsDefContext? explicitContext = null)
    {
        var context = explicitContext ?? GetOrCreateContext(dspShape);
        return context?.TryResolveFlatFill(modelId, schemeColors);
    }

    /// <summary>
    /// Builds a context from already-parsed part roots (colorsDef + dataModel
    /// required; styleDef optional). Returns null when a required root is
    /// missing or has an unexpected root element.
    /// </summary>
    internal static ColorsDefContext? CreateContext(
        OpenXmlElement? colorsDefRoot, OpenXmlElement? dataModelRoot, OpenXmlElement? styleDefRoot)
    {
        if (colorsDefRoot == null || dataModelRoot == null) return null;
        if (colorsDefRoot.LocalName != "colorsDef" || colorsDefRoot.NamespaceUri != DiagramNs) return null;
        if (dataModelRoot.LocalName != "dataModel" || dataModelRoot.NamespaceUri != DiagramNs) return null;
        if (styleDefRoot != null && (styleDefRoot.LocalName != "styleDef" || styleDefRoot.NamespaceUri != DiagramNs))
            return null;

        return new ColorsDefContext(
            ParseNodeStyles(dataModelRoot),
            ParseFillLists(colorsDefRoot),
            ParseQuickStyleGradientLabels(styleDefRoot));
    }

    private static ColorsDefContext? GetOrCreateContext(OpenXmlElement dspShape)
    {
        var root = dspShape;
        while (root.Parent != null) root = root.Parent;

        var holder = ContextCache.GetValue(root, CreateContextHolder);
        return holder.Context;
    }

    private static ContextHolder CreateContextHolder(OpenXmlElement drawingRoot)
    {
        // The drawing part (ppt/diagrams/drawingN.xml) has no relationships of
        // its own in the corpus — the sibling data/colors/quickStyle parts share
        // its directory and numeric suffix, so locate them by URI. Any failure
        // leaves the cached spPr in charge (status quo).
        if (drawingRoot is not OpenXmlPartRootElement partRoot) return new ContextHolder(null);
        var drawingPart = partRoot.OpenXmlPart;
        if (drawingPart == null) return new ContextHolder(null);

        var uri = drawingPart.Uri.OriginalString;
        var slash = uri.LastIndexOf('/');
        var fileName = slash >= 0 ? uri[(slash + 1)..] : uri;
        if (!fileName.StartsWith("drawing", StringComparison.Ordinal)) return new ContextHolder(null);

        var suffix = fileName["drawing".Length..]; // e.g. "35.xml"
        var prefix = slash >= 0 ? uri[..(slash + 1)] : string.Empty;

        var partsByUri = new Dictionary<string, OpenXmlPart>(StringComparer.Ordinal);
        CollectParts(drawingPart.OpenXmlPackage, partsByUri, new HashSet<OpenXmlPartContainer>());

        partsByUri.TryGetValue(prefix + "colors" + suffix, out var colorsPart);
        partsByUri.TryGetValue(prefix + "data" + suffix, out var dataPart);
        partsByUri.TryGetValue(prefix + "quickStyle" + suffix, out var stylePart);

        return new ContextHolder(CreateContext(
            SafeRoot(colorsPart, "colorsDef"),
            SafeRoot(dataPart, "dataModel"),
            SafeRoot(stylePart, "styleDef")));
    }

    private static void CollectParts(OpenXmlPartContainer container, Dictionary<string, OpenXmlPart> partsByUri,
        HashSet<OpenXmlPartContainer> visited)
    {
        // Part relationships are not a tree (e.g. slideMaster ↔ slideLayout) —
        // guard the walk against cycles.
        if (!visited.Add(container)) return;

        foreach (var pair in container.Parts)
        {
            partsByUri.TryAdd(pair.OpenXmlPart.Uri.OriginalString, pair.OpenXmlPart);
            CollectParts(pair.OpenXmlPart, partsByUri, visited);
        }
    }

    private static OpenXmlElement? SafeRoot(OpenXmlPart? part, string expectedRootName)
    {
        if (part == null) return null;
        try
        {
            var root = part.RootElement;
            return root != null && root.LocalName == expectedRootName && root.NamespaceUri == DiagramNs
                ? root
                : null;
        }
        catch (InvalidOperationException)
        {
            // Part XML not readable — treat as absent; cached spPr keeps precedence.
            return null;
        }
    }

    private static Dictionary<string, NodeStyle> ParseNodeStyles(OpenXmlElement dataModelRoot)
    {
        var styles = new Dictionary<string, NodeStyle>(StringComparer.OrdinalIgnoreCase);
        foreach (var pt in Descendants(dataModelRoot, "pt", DiagramNs))
        {
            var modelId = ReadAttribute(pt, "modelId");
            if (string.IsNullOrEmpty(modelId)) continue;

            var prSet = pt.Elements().FirstOrDefault(
                e => e.LocalName == "prSet" && e.NamespaceUri == DiagramNs);
            var label = prSet != null ? ReadAttribute(prSet, "presStyleLbl") : null;
            var indexText = prSet != null ? ReadAttribute(prSet, "presStyleIdx") : null;
            var index = 0;
            if (!string.IsNullOrEmpty(indexText))
            {
                int.TryParse(indexText, NumberStyles.Integer, CultureInfo.InvariantCulture, out index);
            }

            styles[modelId] = new NodeStyle(
                string.IsNullOrEmpty(label) ? "node0" : label,
                index < 0 ? 0 : index);
        }
        return styles;
    }

    private static Dictionary<string, FillList> ParseFillLists(OpenXmlElement colorsDefRoot)
    {
        var lists = new Dictionary<string, FillList>(StringComparer.Ordinal);
        foreach (var styleLbl in Descendants(colorsDefRoot, "styleLbl", DiagramNs))
        {
            var name = ReadAttribute(styleLbl, "name");
            if (string.IsNullOrEmpty(name)) continue;

            var fillClrLst = styleLbl.Elements().FirstOrDefault(
                e => e.LocalName == "fillClrLst" && e.NamespaceUri == DiagramNs);
            if (fillClrLst == null) continue;

            var entries = new List<FillEntry>();
            foreach (var entry in fillClrLst.Elements())
            {
                if (entry.NamespaceUri != DrawingmlNs) continue;
                if (entry.LocalName == "schemeClr")
                {
                    var val = ReadAttribute(entry, "val");
                    if (!string.IsNullOrEmpty(val))
                        entries.Add(new FillEntry(val, null, entry.HasChildren));
                }
                else if (entry.LocalName == "srgbClr")
                {
                    var val = ReadAttribute(entry, "val");
                    if (!string.IsNullOrEmpty(val))
                        entries.Add(new FillEntry(null, "#" + val.TrimStart('#'), entry.HasChildren));
                }
                else
                {
                    // Resolve the non-scheme DrawingML forms through the canonical
                    // color reader. Keeping the entry in the repeat list preserves
                    // colorsDef indexing while allowing scrgbClr/hslClr/sysClr and
                    // preset colors to participate in the flat-fill conflict rule.
                    var resolved = GradientFillReader.ResolveColor(entry);
                    entries.Add(new FillEntry(null, resolved, entry.HasChildren || resolved == null));
                }
            }

            var method = ReadAttribute(fillClrLst, "meth");
            lists[name] = new FillList(
                string.IsNullOrEmpty(method) ? "repeat" : method,
                entries);
        }
        return lists;
    }

    private static HashSet<string> ParseQuickStyleGradientLabels(OpenXmlElement? styleDefRoot)
    {
        var labels = new HashSet<string>(StringComparer.Ordinal);
        if (styleDefRoot == null) return labels;

        foreach (var styleLbl in Descendants(styleDefRoot, "styleLbl", DiagramNs))
        {
            var name = ReadAttribute(styleLbl, "name");
            if (string.IsNullOrEmpty(name)) continue;

            var hasGradientOrEffect = styleLbl.Descendants().Any(e =>
                e.NamespaceUri == DrawingmlNs &&
                e.LocalName is "gradFill" or "effectLst" or "effectDag");
            if (hasGradientOrEffect) labels.Add(name);
        }
        return labels;
    }

    private static IEnumerable<OpenXmlElement> Descendants(OpenXmlElement root, string localName, string namespaceUri)
    {
        return root.Descendants().Where(e => e.LocalName == localName && e.NamespaceUri == namespaceUri);
    }

    private static string? ReadAttribute(OpenXmlElement element, string attributeName)
    {
        var match = Regex.Match(
            element.OuterXml,
            $@"\b{Regex.Escape(attributeName)}\s*=\s*""([^""]*)""",
            RegexOptions.CultureInvariant);

        return match.Success ? match.Groups[1].Value : null;
    }

    private sealed class ContextHolder
    {
        internal ContextHolder(ColorsDefContext? context) => Context = context;
        internal ColorsDefContext? Context { get; }
    }

    internal readonly struct NodeStyle
    {
        internal NodeStyle(string label, int index) => (Label, Index) = (label, index);
        internal string Label { get; }
        internal int Index { get; }
    }

    internal readonly struct FillEntry
    {
        internal FillEntry(string? scheme, string? hex, bool hasTransforms) =>
            (Scheme, Hex, HasTransforms) = (scheme, hex, hasTransforms);
        internal string? Scheme { get; }
        internal string? Hex { get; }
        internal bool HasTransforms { get; }
    }

    internal readonly struct FillList
    {
        internal FillList(string method, List<FillEntry> entries) => (Method, Entries) = (method, entries);
        internal string Method { get; }
        internal List<FillEntry> Entries { get; }
    }

    /// <summary>
    /// Per-diagram colorsDef context: node style labels (from the data part),
    /// fill colour lists (from the colors part) and quickStyle gradient labels.
    /// </summary>
    internal sealed class ColorsDefContext
    {
        private readonly Dictionary<string, NodeStyle> _nodeStyles;
        private readonly Dictionary<string, FillList> _fillLists;
        private readonly HashSet<string> _quickStyleGradientLabels;

        internal ColorsDefContext(
            Dictionary<string, NodeStyle> nodeStyles,
            Dictionary<string, FillList> fillLists,
            HashSet<string> quickStyleGradientLabels)
        {
            _nodeStyles = nodeStyles;
            _fillLists = fillLists;
            _quickStyleGradientLabels = quickStyleGradientLabels;
        }

        /// <summary>
        /// Returns the resolved flat colorsDef fill ("#RRGGBB") for a node, or
        /// null when the clear-conflict rule does not apply. See the precedence
        /// rules on <see cref="SmartArtColorsDefResolver"/>.
        /// </summary>
        internal string? TryResolveFlatFill(string? modelId, IReadOnlyDictionary<string, string>? schemeColors)
        {
            if (string.IsNullOrEmpty(modelId)) return null;
            if (!_nodeStyles.TryGetValue(modelId, out var nodeStyle)) return null;
            if (_quickStyleGradientLabels.Contains(nodeStyle.Label)) return null;
            if (!_fillLists.TryGetValue(nodeStyle.Label, out var fillList)) return null;
            if (fillList.Method != "repeat" || fillList.Entries.Count == 0) return null;

            var entry = fillList.Entries[nodeStyle.Index % fillList.Entries.Count];
            if (entry.HasTransforms) return null;

            if (entry.Scheme != null)
                return SmartArtDrawingExtractor.ResolveSchemeColor(entry.Scheme, schemeColors);
            return entry.Hex;
        }
    }
}
