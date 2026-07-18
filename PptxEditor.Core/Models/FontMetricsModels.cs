namespace PptxEditor.Core.Models;

/// <summary>
/// Construction options for <see cref="PptxEditor.Core.Services.FontMetricsCatalog"/>.
/// </summary>
public sealed record FontMetricsCatalogOptions
{
    /// <summary>Extra directories scanned for *.ttf/*.otf fonts (family names read from name tables).</summary>
    public IReadOnlyList<string> AdditionalFontDirectories { get; init; } = [];

    /// <summary>
    /// Explicit family→file registrations, checked before directory/system scans.
    /// Useful for hosts that manage their own font files and for hermetic tests.
    /// </summary>
    public IReadOnlyList<(string Family, string Path)> AdditionalFontPaths { get; init; } = [];

    /// <summary>
    /// When true (default), well-known system font directories and fontconfig
    /// (<c>fc-list</c>) are scanned, mirroring the converter's chain. Set false
    /// for hermetic behavior (e.g. tests) so only embedded + additional fonts resolve.
    /// </summary>
    public bool IncludeSystemFonts { get; init; } = true;
}

/// <summary>
/// Result of resolving one (family, bold, italic) triple through the
/// embedded-first font metrics chain.
/// </summary>
public sealed record ResolvedFontMetrics
{
    public string RequestedFamily { get; init; } = string.Empty;

    /// <summary>Family actually resolved (differs from requested when substitution applied).</summary>
    public string ResolvedFamily { get; init; } = string.Empty;

    public bool Bold { get; init; }

    public bool Italic { get; init; }

    /// <summary>Parsed metrics; null when the face could not be measured.</summary>
    public TypstFontMetrics? Metrics { get; init; }

    /// <summary>True when a substitution face was used (e.g. Aptos/Calibri → Carlito).</summary>
    public bool WasSubstituted { get; init; }

    /// <summary>True when the face resolved to a TrueType Collection, which cannot be parsed.</summary>
    public bool IsTrueTypeCollection { get; init; }

    /// <summary>Non-fatal accuracy notes produced during resolution.</summary>
    public IReadOnlyList<string> Warnings { get; init; } = [];
}
