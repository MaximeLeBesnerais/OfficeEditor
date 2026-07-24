namespace PptxEditor.Core.Converters.SmartArt;

/// <summary>
/// Logical diagram model extracted from a SmartArt data part (dgm:dataModel).
/// Represents layout metadata, node hierarchy, and inter-node connections —
/// the READ half of a SmartArt &hArr; JSON round-trip.
/// </summary>
public sealed class SmartArtModel
{
    /// <summary>URN layout type identifier, e.g. "urn:microsoft.com/office/officeart/2005/8/layout/process5".</summary>
    public string LayoutTypeId { get; set; } = string.Empty;

    /// <summary>Layout category, e.g. "process", "hierarchy", "list".</summary>
    public string LayoutCategory { get; set; } = string.Empty;

    /// <summary>URN quick-style identifier.</summary>
    public string QuickStyleId { get; set; } = string.Empty;

    /// <summary>URN colour-style identifier.</summary>
    public string ColorStyleId { get; set; } = string.Empty;

    /// <summary>Visual content nodes (type "node").</summary>
    public List<SmartArtNode> Nodes { get; set; } = new();

    /// <summary>Named connections between points in the data model.</summary>
    public List<SmartArtConnection> Connections { get; set; } = new();
}

/// <summary>
/// A single visual node in the diagram's logical model.
/// </summary>
public sealed class SmartArtNode
{
    /// <summary>Unique OOXML model identifier (GUID).</summary>
    public string ModelId { get; set; } = string.Empty;

    /// <summary>Extracted text content from the node's <c>dgm:t</c> element.</summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>Placeholder text hint from <c>dgm:prSet phldrT</c>, e.g. "[Text]".</summary>
    public string PlaceholderText { get; set; } = string.Empty;

    /// <summary>Placeholder index from <c>phldr</c>, or -1 when absent.</summary>
    public int PlaceholderIndex { get; set; } = -1;

    /// <summary>Nesting depth: 0 = top-level child of doc root, 1 = grandchild, etc.</summary>
    public int HierarchyLevel { get; set; }
}

/// <summary>
/// A connection (edge) between two points in the data model.
/// </summary>
public sealed class SmartArtConnection
{
    /// <summary>Unique OOXML model identifier for this connection.</summary>
    public string ModelId { get; set; } = string.Empty;

    /// <summary>Source point modelId.</summary>
    public string SourceId { get; set; } = string.Empty;

    /// <summary>Destination point modelId.</summary>
    public string DestId { get; set; } = string.Empty;

    /// <summary>Connection type: "connection" (layout edge), "presOf" (presentation override), or "presParOf" (presentation parent).</summary>
    public string Type { get; set; } = string.Empty;

    /// <summary>Parent transition modelId, or null.</summary>
    public string? ParTransId { get; set; }

    /// <summary>Sibling transition modelId, or null.</summary>
    public string? SibTransId { get; set; }

    /// <summary>Presentation layout identifier, or null.</summary>
    public string? PresId { get; set; }
}
