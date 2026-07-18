using DocumentFormat.OpenXml.Packaging;
using PptxEditor.Core.Converters;
using PptxEditor.Core.Models;

namespace PptxEditor.Core.Services;

/// <summary>
/// Cached font-metrics resolution for text measurement, keyed by (family, bold, italic).
///
/// The system-font scan delegates to <see cref="PptxToTypstConverter.SharedSystemFonts"/> —
/// the converter's embedded-first chain backing store (same well-known directories +
/// fontconfig), scanned once per process behind its static cache. What intentionally
/// stays here (the converter's public measurement surface does not cover it):
/// per-variant embedded ppt/fonts keying (regular/bold/italic/boldItalic), italic
/// variant names, TTC detection with warnings, Aptos/Calibri substitution, and the
/// catalog options (additional paths/directories, IncludeSystemFonts).
///
/// TrueType Collections (.ttc) are detected and reported as unparseable
/// (warning) — they never crash resolution.
/// </summary>
public sealed class FontMetricsCatalog
{
    private readonly record struct FontKey(string Family, bool Bold, bool Italic);

    private static readonly byte[] TrueTypeCollectionMagic = [(byte)'t', (byte)'t', (byte)'c', (byte)'f'];

    private readonly PresentationDocument? _document;
    private readonly FontMetricsCatalogOptions _options;
    private readonly Dictionary<FontKey, ResolvedFontMetrics> _resolutionCache = new();
    private readonly Dictionary<FontKey, TypstFontMetrics> _embeddedMetrics = new();
    private readonly Dictionary<string, string> _fontPaths = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, (TypstFontMetrics? Metrics, bool IsTtc)> _pathCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _availableFamilies = new(StringComparer.OrdinalIgnoreCase);
    private bool _initialized;

    /// <param name="document">Deck whose embedded ppt/fonts should be preferred. Null = no embedded fonts.</param>
    /// <param name="options">Scan options; defaults mirror the converter (system dirs + fontconfig).</param>
    public FontMetricsCatalog(PresentationDocument? document = null, FontMetricsCatalogOptions? options = null)
    {
        _document = document;
        _options = options ?? new FontMetricsCatalogOptions();
    }

    /// <summary>
    /// Resolves metrics for one face. Never throws for missing/unparseable fonts;
    /// degradation is reported via <see cref="ResolvedFontMetrics.Warnings"/>.
    /// </summary>
    public ResolvedFontMetrics Resolve(string family, bool bold, bool italic)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(family);
        EnsureInitialized();

        var key = new FontKey(family, bold, italic);
        if (_resolutionCache.TryGetValue(key, out var cached))
            return cached;

        var resolved = ResolveCore(family, bold, italic, allowSubstitution: true);
        _resolutionCache[key] = resolved;
        return resolved;
    }

    private ResolvedFontMetrics ResolveCore(string family, bool bold, bool italic, bool allowSubstitution)
    {
        // 1. Embedded ppt/fonts (declared via embeddedFont elements with per-variant refs).
        if (_embeddedMetrics.TryGetValue(new FontKey(family, bold, italic), out var embeddedVariant))
            return Found(family, family, bold, italic, embeddedVariant);
        if ((bold || italic) && _embeddedMetrics.TryGetValue(new FontKey(family, false, false), out var embeddedRegular))
            return Found(family, family, bold, italic, embeddedRegular);

        // 2. Named variant + plain family through the path map
        //    (explicit registrations, then system/fontconfig scan results) —
        //    mirrors the converter's "{family} Bold" lookup then plain family.
        var warnings = new List<string>();
        foreach (var candidate in VariantNames(family, bold, italic))
        {
            if (!_fontPaths.TryGetValue(candidate, out var path))
                continue;

            var (metrics, isTtc) = ReadPathMetrics(path);
            if (isTtc)
            {
                warnings.Add($"Font '{candidate}' resolved to a TrueType Collection ('{path}') which cannot be parsed; metrics unavailable.");
                continue;
            }
            if (metrics != null)
                return Found(family, family, bold, italic, metrics, warnings);
        }

        // 3. Substitution map — mirrors SubstituteUnavailableFont in the converter.
        //    Carlito is the metric stand-in for Calibri; for Aptos it is NOT
        //    metric-compatible, so always warn.
        if (allowSubstitution && !IsAvailable(family))
        {
            var substitute = PickSubstitute(family);
            if (substitute != null)
            {
                var inner = ResolveCore(substitute, bold, italic, allowSubstitution: false);
                var substitutionWarning =
                    $"Font '{family}' is not available; substituted '{substitute}' for measurement. " +
                    "The substitute is not metric-compatible (notably Aptos→Carlito); width estimates are approximate.";
                return inner with
                {
                    RequestedFamily = family,
                    WasSubstituted = true,
                    Warnings = [substitutionWarning, .. inner.Warnings, .. warnings]
                };
            }
        }

        // 4. Unmeasurable — caller decides how to degrade; never silent.
        warnings.Add($"Font '{family}' (bold={bold}, italic={italic}) could not be resolved to font metrics; runs using it are unmeasurable.");
        return new ResolvedFontMetrics
        {
            RequestedFamily = family,
            ResolvedFamily = family,
            Bold = bold,
            Italic = italic,
            Metrics = null,
            Warnings = warnings
        };
    }

    private static ResolvedFontMetrics Found(string requested, string resolved, bool bold, bool italic, TypstFontMetrics metrics, List<string>? warnings = null)
        => new()
        {
            RequestedFamily = requested,
            ResolvedFamily = resolved,
            Bold = bold,
            Italic = italic,
            Metrics = metrics,
            Warnings = warnings ?? (IReadOnlyList<string>)[]
        };

    private static IEnumerable<string> VariantNames(string family, bool bold, bool italic)
    {
        if (bold && italic) yield return $"{family} Bold Italic";
        if (bold) yield return $"{family} Bold";
        if (italic) yield return $"{family} Italic";
        yield return family;
    }

    private bool IsAvailable(string family)
        => _availableFamilies.Contains(family) || VariantNames(family, true, true).Any(_fontPaths.ContainsKey);

    private string? PickSubstitute(string family)
    {
        if (family.StartsWith("Aptos", StringComparison.OrdinalIgnoreCase))
        {
            if (IsAvailable("Aptos")) return "Aptos";
            if (IsAvailable("Carlito")) return "Carlito";
        }
        if (family.Equals("Calibri", StringComparison.OrdinalIgnoreCase))
        {
            if (IsAvailable("Aptos")) return "Aptos";
            if (IsAvailable("Carlito")) return "Carlito";
        }
        return null;
    }

    private (TypstFontMetrics? Metrics, bool IsTtc) ReadPathMetrics(string path)
    {
        if (_pathCache.TryGetValue(path, out var cached))
            return cached;

        (TypstFontMetrics?, bool) result;
        if (IsTrueTypeCollection(path))
        {
            result = (null, true);
        }
        else
        {
            result = (OpenTypeFontMetricsReader.ReadMetrics(path), false);
        }

        _pathCache[path] = result;
        return result;
    }

    private static bool IsTrueTypeCollection(string path)
    {
        if (path.EndsWith(".ttc", StringComparison.OrdinalIgnoreCase))
            return true;
        try
        {
            Span<byte> magic = stackalloc byte[4];
            using var stream = File.OpenRead(path);
            return stream.Read(magic) == 4 && magic.SequenceEqual(TrueTypeCollectionMagic);
        }
        catch
        {
            return false; // unreadable files fail later in ReadMetrics with a null result
        }
    }

    private void EnsureInitialized()
    {
        if (_initialized) return;
        _initialized = true;

        ExtractEmbeddedFonts();
        RegisterAdditionalPaths();
        ScanDirectories(_options.AdditionalFontDirectories);
        if (_options.IncludeSystemFonts)
        {
            // Converter-owned chain: same well-known directories + fontconfig scan,
            // performed at most once per process behind the converter's static cache
            // (never a per-catalog rescan). TTC paths from fontconfig land here too and
            // are flagged at read time by ReadPathMetrics.
            var (families, paths) = PptxToTypstConverter.SharedSystemFonts;
            _availableFamilies.UnionWith(families);
            foreach (var (family, path) in paths)
            {
                _fontPaths.TryAdd(family, path);
            }
        }
    }

    // Embedded-first chain, step 1: ppt/fonts declared via embeddedFont elements.
    // Mirrors PptxToTypstConverter.ExtractFonts/ExtractFontFile, including the
    // EOT-wrapper unwrap, but keys metrics by the declared variant.
    private void ExtractEmbeddedFonts()
    {
        var presentationPart = _document?.PresentationPart;
        if (presentationPart?.Presentation == null)
            return;

        var embeddedFonts = presentationPart.Presentation
            .Descendants()
            .Where(e => e.LocalName == "embeddedFont")
            .ToList();

        foreach (var embeddedFont in embeddedFonts)
        {
            var family = embeddedFont.ChildElements.FirstOrDefault(e => e.LocalName == "font")
                ?.GetAttribute("typeface", "").Value;
            if (string.IsNullOrWhiteSpace(family))
                continue;

            foreach (var fontReference in embeddedFont.ChildElements.Where(e => e.LocalName is "regular" or "bold" or "italic" or "boldItalic"))
            {
                var relationshipId = fontReference.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships").Value;
                if (string.IsNullOrEmpty(relationshipId))
                    continue;
                if (presentationPart.GetPartById(relationshipId) is not FontPart fontPart)
                    continue;

                var data = ReadPartBytes(fontPart);
                if (data == null)
                    continue;

                var unwrapped = UnwrapFontData(data);
                var metrics = unwrapped != null ? OpenTypeFontMetricsReader.TryRead(unwrapped) : null;
                if (metrics == null)
                    continue; // unparseable (incl. TTC inside a FontPart): skip, system fallback still possible

                var (bold, italic) = fontReference.LocalName switch
                {
                    "bold" => (true, false),
                    "italic" => (false, true),
                    "boldItalic" => (true, true),
                    _ => (false, false)
                };
                _embeddedMetrics.TryAdd(new FontKey(family, bold, italic), metrics);
                _availableFamilies.Add(family);
            }
        }
    }

    private static byte[]? ReadPartBytes(FontPart fontPart)
    {
        try
        {
            using var stream = fontPart.GetStream();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            return ms.ToArray();
        }
        catch
        {
            return null;
        }
    }

    // EOT unwrap mirrored from the converter: locate the OTTO/0x00010000 sfnt
    // signature inside the payload and slice to the declared font-data length.
    private static byte[]? UnwrapFontData(byte[] data)
    {
        var otfPos = data.AsSpan().IndexOf(System.Text.Encoding.ASCII.GetBytes("OTTO"));
        var ttfPos = data.AsSpan().IndexOf(new byte[] { 0x00, 0x01, 0x00, 0x00 });

        var pos = otfPos >= 0 ? otfPos : ttfPos;
        if (pos < 0)
            return null;

        var length = TryReadEotFontDataLength(data, pos);
        return data[pos..(pos + length)];
    }

    private static int TryReadEotFontDataLength(byte[] data, int fontDataOffset)
    {
        if (data.Length >= 8)
        {
            var fontDataLength = BitConverter.ToUInt32(data, 4);
            if (fontDataLength > 0 && fontDataLength <= int.MaxValue && fontDataOffset + fontDataLength <= data.Length)
                return (int)fontDataLength;
        }
        return data.Length - fontDataOffset;
    }

    private void RegisterAdditionalPaths()
    {
        foreach (var (family, path) in _options.AdditionalFontPaths)
        {
            if (string.IsNullOrWhiteSpace(family) || string.IsNullOrWhiteSpace(path))
                continue;
            _fontPaths.TryAdd(family, path);
            _availableFamilies.Add(family);
        }
    }

    private void ScanDirectories(IEnumerable<string> directories)
    {
        foreach (var dir in directories.Where(Directory.Exists))
        {
            IEnumerable<string> files;
            try
            {
                files = Directory.GetFiles(dir, "*.ttf", SearchOption.AllDirectories)
                    .Concat(Directory.GetFiles(dir, "*.otf", SearchOption.AllDirectories));
            }
            catch
            {
                continue; // skip inaccessible directories
            }

            foreach (var fontFile in files)
            {
                try
                {
                    var familyName = OpenTypeFontMetricsReader.ReadFontFamilyName(fontFile);
                    if (!string.IsNullOrEmpty(familyName))
                    {
                        _availableFamilies.Add(familyName);
                        _fontPaths.TryAdd(familyName, fontFile);
                    }
                }
                catch { /* skip unreadable fonts */ }
            }
        }
    }
}
