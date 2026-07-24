using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace PptxEditor.Core.Converters.SmartArt;

/// <summary>
/// Extracts a <see cref="SmartArtModel"/> from a SmartArt diagram's data part
/// (dgm:data1.xml).  Parses the point list for node text and hierarchy metadata
/// and the connection list for parent-child relationships.
/// </summary>
public static class SmartArtDataExtractor
{
    private const string DiagramNs = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
    private const string DrawingmlNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// <summary>
    /// Extracts the logical diagram model from the data part pointed to by
    /// <c>dgm:relIds</c> inside a diagram graphic frame.
    /// Returns <c>null</c> when the slide has no resolvable data part.
    /// </summary>
    public static SmartArtModel? Extract(SlidePart slidePart, GraphicFrame graphicFrame)
    {
        var graphicData = graphicFrame.Graphic?.GraphicData;
        if (graphicData == null) return null;

        // Try the standard path: resolve data part via dgm:relIds.
        var dataPart = ResolveDataPart(slidePart, graphicData);
        if (dataPart != null) return ExtractFromPart(dataPart);

        // Fallback: scan slide parts for dgm:dataModel. This mirrors the
        // brute-force approach used by PptxToTypstConverter.ConvertDiagramGraphicFrame.
        return ExtractFromSlideParts(slidePart);
    }

    /// <summary>
    /// Scans all slide and diagram parts looking for a <c>dgm:dataModel</c> root element.
    /// </summary>
    internal static SmartArtModel? ExtractFromSlideParts(SlidePart slidePart)
    {
        foreach (var partPair in slidePart.Parts)
        {
            var model = TryExtractFromPart(partPair.OpenXmlPart);
            if (model != null) return model;
        }

        foreach (var diagramPart in slidePart.GetPartsOfType<DiagramPersistLayoutPart>())
        {
            var model = TryExtractFromPart(diagramPart);
            if (model != null) return model;

            foreach (var childPartPair in diagramPart.Parts)
            {
                var childModel = TryExtractFromPart(childPartPair.OpenXmlPart);
                if (childModel != null) return childModel;
            }
        }

        return null;
    }

    private static SmartArtModel? TryExtractFromPart(OpenXmlPart part)
    {
        try
        {
            var root = part.RootElement;
            if (root != null &&
                root.LocalName == "dataModel" &&
                root.NamespaceUri == DiagramNs)
            {
                return ExtractFromDataModelElement(root);
            }
        }
        catch
        {
            // Part not readable — skip.
        }

        return null;
    }

    /// <summary>
    /// Extracts the model from an already-resolved data part root element.
    /// </summary>
    public static SmartArtModel? ExtractFromPart(OpenXmlPart dataPart)
    {
        var root = dataPart.RootElement;
        return root == null ? null : ExtractFromDataModelElement(root);
    }

    /// <summary>
    /// Extracts the model from a raw <c>dgm:dataModel</c> XML element, e.g.
    /// from an already-parsed part or an in-memory test fixture.
    /// </summary>
    internal static SmartArtModel? ExtractFromDataModelElement(OpenXmlElement dataModelElement)
    {
        var model = new SmartArtModel();

        var ptLst = dataModelElement.Elements()
            .FirstOrDefault(e => e.LocalName == "ptLst" && e.NamespaceUri == DiagramNs);
        if (ptLst == null) return null;

        var points = ptLst.Elements()
            .Where(e => e.LocalName == "pt" && e.NamespaceUri == DiagramNs)
            .ToList();

        var cxnLst = dataModelElement.Elements()
            .FirstOrDefault(e => e.LocalName == "cxnLst" && e.NamespaceUri == DiagramNs);

        var connections = cxnLst != null
            ? cxnLst.Elements()
                .Where(e => e.LocalName == "cxn" && e.NamespaceUri == DiagramNs)
                .ToList()
            : new List<OpenXmlElement>();

        var pointsById = new Dictionary<string, OpenXmlElement>(StringComparer.OrdinalIgnoreCase);
        var pointTypes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var docId = string.Empty;

        foreach (var pt in points)
        {
            var modelId = ReadAttribute(pt, "modelId") ?? string.Empty;
            if (string.IsNullOrEmpty(modelId)) continue;

            pointsById[modelId] = pt;

            var type = ReadAttribute(pt, "type") ?? "node";
            pointTypes[modelId] = type;

            var prSet = GetChild(pt, "prSet", DiagramNs);

            if (type == "doc")
            {
                docId = modelId;
                if (prSet != null)
                {
                    model.LayoutTypeId = ReadAttribute(prSet, "loTypeId") ?? string.Empty;
                    model.LayoutCategory = ReadAttribute(prSet, "loCatId") ?? string.Empty;
                    model.QuickStyleId = ReadAttribute(prSet, "qsTypeId") ?? string.Empty;
                    model.ColorStyleId = ReadAttribute(prSet, "csTypeId") ?? string.Empty;
                }
            }

            if (type is "node" or "doc" or "parTrans" or "sibTrans" || string.IsNullOrEmpty(type))
            {
                var text = ExtractText(pt);
                if (!string.IsNullOrEmpty(text) || type == "node" || type == "doc" || string.IsNullOrEmpty(type))
                {
                    var plText = prSet != null ? ReadAttribute(prSet, "phldrT") ?? string.Empty : string.Empty;
                    // phldr is xsd:boolean ("1"/"true"), not an integer index.
                    var phldrStr = prSet != null ? ReadAttribute(prSet, "phldr") : null;
                    var isPlaceholder = string.Equals(phldrStr, "1", StringComparison.Ordinal) ||
                                        string.Equals(phldrStr, "true", StringComparison.OrdinalIgnoreCase);

                    if (type != "doc")
                    {
                        model.Nodes.Add(new SmartArtNode
                        {
                            ModelId = modelId,
                            Text = text ?? string.Empty,
                            PlaceholderText = plText,
                            IsPlaceholder = isPlaceholder,
                            HierarchyLevel = 0
                        });
                    }
                }
            }
        }

        foreach (var cxn in connections)
        {
            var cxnModelId = ReadAttribute(cxn, "modelId") ?? string.Empty;
            var srcId = ReadAttribute(cxn, "srcId") ?? string.Empty;
            var destId = ReadAttribute(cxn, "destId") ?? string.Empty;
            var type = ReadAttribute(cxn, "type") ?? "connection";
            var parTransId = ReadAttribute(cxn, "parTransId");
            var sibTransId = ReadAttribute(cxn, "sibTransId");
            var presId = ReadAttribute(cxn, "presId");

            model.Connections.Add(new SmartArtConnection
            {
                ModelId = cxnModelId,
                SourceId = srcId,
                DestId = destId,
                Type = type,
                ParTransId = parTransId,
                SibTransId = sibTransId,
                PresId = presId
            });
        }

        ComputeHierarchyLevels(model, pointTypes, docId);

        return model;
    }

    private static void ComputeHierarchyLevels(SmartArtModel model,
        Dictionary<string, string> pointTypes, string docId)
    {
        if (string.IsNullOrEmpty(docId) || model.Connections.Count == 0)
            return;

        var levelByNode = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        var nodeIds = new HashSet<string>(model.Nodes.Select(n => n.ModelId), StringComparer.OrdinalIgnoreCase);

        var edgesFromSource = model.Connections
            .Where(c => !string.IsNullOrEmpty(c.SourceId) && !string.IsNullOrEmpty(c.DestId))
            .GroupBy(c => c.SourceId, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase);

        // BFS from doc root to assign levels
        var queue = new Queue<(string id, int level)>();
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        if (!string.IsNullOrEmpty(docId))
        {
            queue.Enqueue((docId, -1));
            visited.Add(docId);
        }

        while (queue.Count > 0)
        {
            var (currentId, currentLevel) = queue.Dequeue();

            if (!edgesFromSource.TryGetValue(currentId, out var outgoing))
                continue;

            foreach (var cxn in outgoing)
            {
                var destId = cxn.DestId;
                if (visited.Contains(destId)) continue;

                // Skip presentation overrides — they don't define hierarchy structure.
                // Only "connection" edges (with parTransId/sibTransId) and "presParOf" edges
                // carry topology.
                var isHierarchyEdge = cxn.Type == "connection" ||
                                      cxn.Type == "presParOf" ||
                                      string.IsNullOrEmpty(cxn.Type);

                if (isHierarchyEdge)
                {
                    // Use intermediate transition nodes to extend the graph
                    if (!string.IsNullOrEmpty(cxn.ParTransId) && visited.Add(cxn.ParTransId))
                    {
                        queue.Enqueue((cxn.ParTransId, currentLevel + 1));
                    }

                    if (!string.IsNullOrEmpty(cxn.SibTransId) && visited.Add(cxn.SibTransId))
                    {
                        queue.Enqueue((cxn.SibTransId, currentLevel + 1));
                    }

                    // Mark + enqueue the destination exactly once: a previous
                    // implementation added destId to visited before this guard,
                    // so the guard never passed and no level was ever assigned.
                    if (visited.Add(destId))
                    {
                        // Determine if dest is a content node
                        var isContentNode = nodeIds.Contains(destId);
                        if (isContentNode || pointTypes.TryGetValue(destId, out var ptType) &&
                            (ptType == "node" || string.IsNullOrEmpty(ptType)))
                        {
                            levelByNode[destId] = currentLevel + 1;
                        }

                        queue.Enqueue((destId, currentLevel + 1));
                    }
                }
                else if (cxn.Type == "presOf")
                {
                    // "presOf" connects a real node to its presentation override.
                    // The override is a child in the presentation hierarchy but
                    // not a content node — just mark it visited and enqueue.
                    if (visited.Add(destId))
                    {
                        queue.Enqueue((destId, currentLevel));
                    }
                }
            }
        }

        // Apply computed levels to nodes
        foreach (var node in model.Nodes)
        {
            if (levelByNode.TryGetValue(node.ModelId, out var level))
            {
                node.HierarchyLevel = level;
            }
        }
    }

    private static string? ExtractText(OpenXmlElement pt)
    {
        var t = GetChild(pt, "t", DiagramNs);
        if (t == null) return null;

        // Join paragraphs with a newline separator instead of concatenating them.
        var paragraphTexts = new List<string>();
        var paragraphs = t.Elements()
            .Where(e => e.LocalName == "p" && e.NamespaceUri == DrawingmlNs);

        foreach (var p in paragraphs)
        {
            var paragraphText = new System.Text.StringBuilder();

            // Text lives in runs (a:r) and fields (a:fld); both carry a:t children.
            foreach (var child in p.Elements())
            {
                if (child.LocalName is not ("r" or "fld") || child.NamespaceUri != DrawingmlNs)
                    continue;

                foreach (var tPart in child.Elements()
                             .Where(e => e.LocalName == "t" && e.NamespaceUri == DrawingmlNs)
                             .Select(e => e.InnerText))
                {
                    paragraphText.Append(tPart);
                }
            }

            paragraphTexts.Add(paragraphText.ToString());
        }

        var text = string.Join("\n", paragraphTexts);
        return text.Length > 0 ? text : null;
    }

    private static OpenXmlPart? ResolveDataPart(SlidePart slidePart, OpenXmlElement graphicData)
    {
        var relIdsElement = graphicData.Elements()
            .FirstOrDefault(e => e.LocalName == "relIds" && e.NamespaceUri == DiagramNs);

        if (relIdsElement == null) return null;

        var attrs = relIdsElement.GetAttributes();
        // Pattern: r:dm=<data rId>, r:lo=<layout rId>, r:qs=<quickStyle rId>, r:cs=<colors rId>
        var rIds = attrs
            .Where(a => !string.IsNullOrEmpty(a.Value) &&
                        a.Value.StartsWith("rId", StringComparison.Ordinal))
            .Select(a => a.Value)
            .Distinct()
            .ToList();

        foreach (var rId in rIds)
        {
            try
            {
                var part = slidePart.GetPartById(rId!);
                if (part != null)
                {
                    // The data part is the one whose root element is dgm:dataModel
                    var root = part.RootElement;
                    if (root != null &&
                        root.LocalName == "dataModel" &&
                        root.NamespaceUri == DiagramNs)
                    {
                        return part;
                    }
                }
            }
            catch
            {
                // Part not resolvable — skip.
            }
        }

        return null;
    }

    private static string? ReadAttribute(OpenXmlElement element, string attributeName)
    {
        var match = Regex.Match(
            element.OuterXml,
            $@"\b{Regex.Escape(attributeName)}\s*=\s*""([^""]*)""",
            RegexOptions.CultureInvariant);

        return match.Success ? match.Groups[1].Value : null;
    }

    private static OpenXmlElement? GetChild(OpenXmlElement parent, string localName, string namespaceUri)
    {
        return parent.Elements()
            .FirstOrDefault(e => e.LocalName == localName && e.NamespaceUri == namespaceUri);
    }
}
