using System.Globalization;
using System.Diagnostics;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PptxEditor.Core.Converters.SmartArt;
using PptxEditor.Core.Models;

namespace PptxEditor.Core.Converters;

public sealed partial class PptxToTypstConverter : IDisposable
{
    private readonly PresentationDocument _document;
    private readonly string _tempDirectory;
    private readonly string _assetsDirectory;
    private readonly string _fontsDirectory;
    private readonly Dictionary<string, TypstFontMetrics> _fontMetrics = new(StringComparer.OrdinalIgnoreCase);
    private Dictionary<string, string> _themeFonts = new(StringComparer.OrdinalIgnoreCase);
    private Dictionary<string, TableStyleDefinition> _tableStyles = new(StringComparer.OrdinalIgnoreCase);
    private int _imageCounter;

    /// <summary>Warnings for the slide currently being converted; null outside slide conversion.</summary>
    private List<string>? _activeSlideWarnings;

    /// <summary>Target PPI for upscale detection. Used to determine if a native image
    /// is smaller than the display size and would be blurred by Typst upscaling.</summary>
    public float Ppi { get; init; } = 150;

    private readonly HashSet<string> _knownSystemFonts = new(StringComparer.OrdinalIgnoreCase)
    {
        "Arial", "Helvetica", "Liberation Sans", "Liberation Serif", "DejaVu Sans", "DejaVu Serif",
        "Times New Roman", "Courier New", "Georgia", "Verdana", "Trebuchet MS", "Palatino",
        "Garamond", "Bookman", "Comic Sans MS", "Impact", "Candara", "Calibri", "Carlito", "Cambria",
        "Constantia", "Corbel", "Franklin Gothic", "Gabriola", "Segoe UI"
    };
    private HashSet<string> _availableSystemFonts = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _systemFontPaths = new(StringComparer.OrdinalIgnoreCase);

    public PptxToTypstConverter(PresentationDocument document)
    {
        _document = document ?? throw new ArgumentNullException(nameof(document));
        _tempDirectory = Path.Combine(Path.GetTempPath(), "OfficeEditor", Guid.NewGuid().ToString("N"));
        _assetsDirectory = Path.Combine(_tempDirectory, "assets");
        _fontsDirectory = Path.Combine(_tempDirectory, "fonts");
        Directory.CreateDirectory(_assetsDirectory);
        Directory.CreateDirectory(_fontsDirectory);
    }

    public TypstPresentation Convert()
    {
        var presentation = BeginConversion();

        var slideIdList = _document.PresentationPart!.Presentation!.SlideIdList;
        if (slideIdList == null) return presentation;

        int slideIndex = 1;
        foreach (var slideId in slideIdList.ChildElements.OfType<SlideId>())
        {
            var slidePart = (SlidePart)_document.PresentationPart.GetPartById(slideId.RelationshipId!);
            var slide = slidePart.Slide;
            if (slide == null) continue;

            var typstSlide = ConvertSlide(slidePart, slide, slideIndex);
            presentation.Slides.Add(typstSlide);
            slideIndex++;
        }

        return presentation;
    }

    /// <summary>
    /// Converts a single slide and returns a <see cref="TypstPresentation"/> containing
    /// exactly that one slide.
    /// </summary>
    /// <param name="slideIndex">0-based index into the presentation's slide order.</param>
    /// <remarks>
    /// Pair with <see cref="GenerateTypstSource"/> to emit a self-contained one-page
    /// document: the global <c>#set text</c> header plus this slide's <c>#set page</c>
    /// block, with no cross-slide state (per-slide pages in the whole-deck emission are
    /// delimited only by <c>#pagebreak()</c>). Compiling the emitted source therefore
    /// yields exactly one page (<c>Pages.Length == 1</c>) matching the corresponding page
    /// of the whole-deck render.
    /// </remarks>
    public TypstPresentation ConvertSingleSlide(int slideIndex)
    {
        var slideIds = (_document.PresentationPart!.Presentation!.SlideIdList?.ChildElements
            .OfType<SlideId>().ToList()) ?? new List<SlideId>();

        if (slideIndex < 0 || slideIndex >= slideIds.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(slideIndex), slideIndex,
                $"Slide index must be in the range [0, {slideIds.Count}).");
        }

        var presentation = BeginConversion();

        var slidePart = (SlidePart)_document.PresentationPart.GetPartById(slideIds[slideIndex].RelationshipId!);
        var slide = slidePart.Slide;
        if (slide != null)
        {
            // TypstSlide.SlideIndex stays 1-based (used e.g. for slide-number placeholders).
            presentation.Slides.Add(ConvertSlide(slidePart, slide, slideIndex + 1));
        }

        return presentation;
    }

    /// <summary>
    /// Performs the conversion setup shared by <see cref="Convert"/> and
    /// <see cref="ConvertSingleSlide"/>: per-deck embedded-font extraction, theme fonts,
    /// system fonts (process-wide cache), and table styles.
    /// </summary>
    private TypstPresentation BeginConversion()
    {
        _fontMetrics.Clear();
        ExtractFonts();
        var fontFiles = Directory.GetFiles(_fontsDirectory).ToList();
        _themeFonts = ExtractThemeFonts();
        _availableSystemFonts = DiscoverSystemFonts();
        _fontsInitialized = true;

        // Load table styles using first slide's theme for scheme color resolution
        var firstSlideId = _document.PresentationPart!.Presentation!.SlideIdList?.ChildElements.OfType<SlideId>().FirstOrDefault();
        if (firstSlideId != null)
        {
            var firstSlidePart = (SlidePart)_document.PresentationPart.GetPartById(firstSlideId.RelationshipId!);
            var globalStyleResolver = new StyleResolver(_document, firstSlidePart);
            LoadTableStyles(globalStyleResolver);
        }
        else
        {
            LoadTableStyles(null);
        }

        return new TypstPresentation
        {
            TempDirectory = _tempDirectory,
            FontFiles = fontFiles,
            FontMetrics = new Dictionary<string, TypstFontMetrics>(_fontMetrics, StringComparer.OrdinalIgnoreCase),
            ThemeFonts = _themeFonts
        };
    }

    private TypstFontMetrics? GetFontMetrics(string fontFamily)
    {
        // 1. Check embedded font cache
        if (_fontMetrics.TryGetValue(fontFamily, out var metrics))
            return metrics;

        // 2. Try to find and read system font
        var systemFontPath = FindSystemFontPath(fontFamily);
        if (systemFontPath != null)
        {
            try
            {
                var m = OpenTypeFontMetricsReader.ReadMetrics(systemFontPath);
                if (m != null)
                {
                    _fontMetrics[fontFamily] = m;
                    return m;
                }
            }
            catch { /* skip unreadable font metrics */ }
        }

        return null;
    }

    private string? FindSystemFontPath(string fontFamily)
    {
        if (_systemFontPaths.TryGetValue(fontFamily, out var path))
            return path;
        return null;
    }

    private bool TryGetFontMetrics(TypstTextElement text, out TypstFontMetrics metrics)
    {
        if (text.Formatting.Bold)
        {
            var boldMetrics = GetFontMetrics($"{text.Formatting.FontFamily} Bold");
            if (boldMetrics != null)
            {
                metrics = boldMetrics;
                return true;
            }
        }

        var familyMetrics = GetFontMetrics(text.Formatting.FontFamily);
        if (familyMetrics != null)
        {
            metrics = familyMetrics;
            return true;
        }

        metrics = null!;
        return false;
    }

    private bool _fontsInitialized;

    /// <summary>
    /// Lazily performs the font half of <see cref="BeginConversion"/> (per-deck embedded
    /// fonts + system font paths) for the public measurement surface, without running a
    /// slide conversion. The system scan is never repeated per call: it routes through
    /// the process-wide cache (<see cref="GetOrScanSystemFonts"/>). A prior
    /// <see cref="Convert"/>/<see cref="ConvertSingleSlide"/> on this instance counts as
    /// initialization.
    /// </summary>
    private void EnsureFontsInitialized()
    {
        if (_fontsInitialized)
        {
            return;
        }

        _fontMetrics.Clear();
        ExtractFonts();
        _availableSystemFonts = DiscoverSystemFonts();
        _fontsInitialized = true;
    }

    /// <summary>
    /// W5 DEDUP-INTEGRATION: public read access to the converter's embedded-first
    /// font-metrics resolution chain (embedded ppt/fonts cache → system font paths via
    /// the process-wide scan → "{family} Bold" variant fallback), so measurement code
    /// (PptxEditor.Core/Services/FontMetricsCatalog and friends) can delegate to it
    /// instead of duplicating the chain.
    /// </summary>
    public TypstFontMetrics? ResolveFontMetricsForMeasurement(string fontFamily, bool bold)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(fontFamily);

        EnsureFontsInitialized();

        if (bold)
        {
            var boldMetrics = GetFontMetrics($"{fontFamily} Bold");
            if (boldMetrics != null)
            {
                return boldMetrics;
            }
        }

        return GetFontMetrics(fontFamily);
    }

    /// <summary>
    /// W5 DEDUP-INTEGRATION: resolved family→file map from the system/fontconfig scan
    /// (includes .ttc paths reported by fontconfig, so measurement can flag TrueType
    /// Collections as unparseable instead of silently failing metric reads).
    /// </summary>
    public IReadOnlyDictionary<string, string> ResolvedSystemFontPaths
    {
        get
        {
            EnsureFontsInitialized();
            return _systemFontPaths;
        }
    }

    /// <summary>
    /// Shared process-wide system-font snapshot (families + family→file paths) backing
    /// <see cref="ResolvedSystemFontPaths"/>, exposed internally so FontMetricsCatalog
    /// consumes the same single scan instead of duplicating it per instance.
    /// </summary>
    internal static (IReadOnlySet<string> Families, IReadOnlyDictionary<string, string> Paths) SharedSystemFonts
    {
        get
        {
            var discovery = GetOrScanSystemFonts();
            return (discovery.Families, discovery.Paths);
        }
    }

    /// <summary>Process-wide snapshot of a system-font scan (families + file paths).</summary>
    private sealed class SystemFontDiscovery
    {
        public required HashSet<string> Families { get; init; }
        public required Dictionary<string, string> Paths { get; init; }
    }

    private static readonly object s_systemFontCacheLock = new();
    private static SystemFontDiscovery? s_systemFontCache;

    /// <summary>
    /// Clears the process-wide system-font cache. The next conversion re-scans system font
    /// directories and fontconfig. Call this after installing or removing fonts; the cache
    /// otherwise lives for the process lifetime (system fonts are stable within a run).
    /// Embedded PPTX fonts are extracted per deck and are unaffected by this cache.
    /// </summary>
    public static void InvalidateSystemFontCache()
    {
        lock (s_systemFontCacheLock)
        {
            s_systemFontCache = null;
        }
    }

    private static SystemFontDiscovery GetOrScanSystemFonts()
    {
        lock (s_systemFontCacheLock)
        {
            // The scan (recursive directory walk + per-file name-table parse + fc-list)
            // costs 100-1000 ms, so it is performed at most once per process; the lock is
            // held across it to prevent concurrent duplicate scans on first use.
            s_systemFontCache ??= ScanSystemFonts();
            return s_systemFontCache;
        }
    }

    private HashSet<string> DiscoverSystemFonts()
    {
        var discovery = GetOrScanSystemFonts();

        _systemFontPaths.Clear();
        foreach (var kv in discovery.Paths)
        {
            _systemFontPaths[kv.Key] = kv.Value;
        }

        return new HashSet<string>(discovery.Families, StringComparer.OrdinalIgnoreCase);
    }

    private static SystemFontDiscovery ScanSystemFonts()
    {
        var systemFontDirs = new[]
        {
            "/usr/share/fonts",
            "/usr/local/share/fonts",
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts)),
            "/System/Library/Fonts",
            "/Library/Fonts",
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".fonts"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".local", "share", "fonts"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts"),
        };

        var fontFamilies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var fontPaths = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var dir in systemFontDirs.Where(Directory.Exists))
        {
            try
            {
                foreach (var fontFile in Directory.GetFiles(dir, "*.ttf", SearchOption.AllDirectories)
                    .Concat(Directory.GetFiles(dir, "*.otf", SearchOption.AllDirectories)))
                {
                    try
                    {
                        var familyName = OpenTypeFontMetricsReader.ReadFontFamilyName(fontFile);
                        if (!string.IsNullOrEmpty(familyName))
                        {
                            fontFamilies.Add(familyName);
                            fontPaths.TryAdd(familyName, fontFile);
                        }
                    }
                    catch { /* skip unreadable fonts */ }
                }
            }
            catch { /* skip inaccessible directories */ }
        }

        DiscoverFontConfigFonts(fontFamilies, fontPaths);

        return new SystemFontDiscovery { Families = fontFamilies, Paths = fontPaths };
    }

    private static void DiscoverFontConfigFonts(HashSet<string> fontFamilies, Dictionary<string, string> fontPaths)
    {
        try
        {
            using var process = Process.Start(new ProcessStartInfo
            {
                FileName = "fc-list",
                ArgumentList = { "--format=%{file}\t%{family}\n" },
                RedirectStandardOutput = true,
                RedirectStandardError = false,
                UseShellExecute = false,
                CreateNoWindow = true
            });

            if (process == null)
                return;

            var outputTask = process.StandardOutput.ReadToEndAsync();
            if (!process.WaitForExit(2000))
            {
                try { process.Kill(entireProcessTree: true); }
                catch { /* best effort cleanup */ }

                return;
            }

            if (!process.HasExited || process.ExitCode != 0)
                return;

            if (!outputTask.Wait(500))
                return;

            var output = outputTask.Result;

            foreach (var line in output.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                var parts = line.Split('\t', 2);
                if (parts.Length != 2 || string.IsNullOrWhiteSpace(parts[1]))
                    continue;

                var fontPath = parts[0].Trim();
                foreach (var family in parts[1].Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
                {
                    fontFamilies.Add(family);
                    if (!string.IsNullOrEmpty(fontPath))
                        fontPaths.TryAdd(family, fontPath);
                }
            }
        }
        catch
        {
            // Fontconfig is optional; keep directory scanning as the portable fallback.
        }
    }

    private static List<string> BuildGlobalFontFamilies(IReadOnlyDictionary<string, string> themeFonts, HashSet<string> availableFonts)
    {
        var fontFamilies = new List<string>();

        if (themeFonts.TryGetValue("+mn-lt", out var minorFont) || themeFonts.TryGetValue("minor-latin", out minorFont))
        {
            var resolvedMinorFont = SubstituteUnavailableFont(minorFont, availableFonts);
            AddFontFamily(fontFamilies, resolvedMinorFont, availableFonts);
        }

        AddFontFamily(fontFamilies, "Carlito", availableFonts);
        AddFontFamily(fontFamilies, "Arial", availableFonts);
        AddFontFamily(fontFamilies, "Helvetica", availableFonts);
        AddFontFamily(fontFamilies, "Liberation Sans", availableFonts);

        return fontFamilies;
    }

    private static void AddFontFamily(List<string> fontFamilies, string fontFamily, HashSet<string> availableFonts)
    {
        if (string.IsNullOrWhiteSpace(fontFamily)
            || !availableFonts.Contains(fontFamily)
            || fontFamilies.Contains(fontFamily, StringComparer.OrdinalIgnoreCase))
        {
            return;
        }

        fontFamilies.Add(fontFamily);
    }


    private TypstSlide ConvertSlide(SlidePart slidePart, Slide slide, int slideIndex)
    {
        var styleResolver = new StyleResolver(_document, slidePart);
        var layout = ExtractSlideLayout(slidePart, slide, styleResolver);
        var typstSlide = new TypstSlide
        {
            SlideIndex = slideIndex,
            Layout = layout
        };

        var shapeTree = slide.CommonSlideData?.ShapeTree;
        if (shapeTree == null) return typstSlide;

        // Warnings raised while converting this slide's elements (e.g. unsupported
        // graphic frames) are collected here and exposed on TypstSlide.Warnings.
        var warnings = new List<string>();
        _activeSlideWarnings = warnings;
        try
        {
            foreach (var element in shapeTree.ChildElements)
            {
                foreach (var typstElement in ConvertElement(slidePart, element, styleResolver, slideIndex))
                {
                    typstSlide.Elements.Add(typstElement);
                }
            }
        }
        finally
        {
            _activeSlideWarnings = null;
        }

        foreach (var warning in warnings)
        {
            typstSlide.AddWarning(warning);
        }

        return typstSlide;
    }

    private void AddSlideWarning(string warning)
    {
        _activeSlideWarnings?.Add(warning);
    }

    private Models.SlideLayout ExtractSlideLayout(SlidePart slidePart, Slide slide, StyleResolver styleResolver)
    {
        // Default to 16:9 (960pt x 540pt at 96 DPI, or 10in x 5.625in = 720pt x 405pt)
        double width = 720;
        double height = 405;
        string? bgColor = null;

        // Try to get slide dimensions from presentation
        var presentation = _document.PresentationPart!.Presentation!;
        var sldSz = presentation.SlideSize;
        if (sldSz != null)
        {
            width = EmuToPt(sldSz.Cx?.Value ?? 9144000);
            height = EmuToPt(sldSz.Cy?.Value ?? 6858000);
        }

        // Background - use style resolver for cascade: slide -> layout -> master
        bgColor = styleResolver.ResolveBackgroundColor();

        return new Models.SlideLayout
        {
            Width = width,
            Height = height,
            BackgroundColor = bgColor
        };
    }

    private IEnumerable<TypstElement> ConvertElement(SlidePart slidePart, OpenXmlElement element, StyleResolver styleResolver, int slideIndex, double offX = 0, double offY = 0, double scaleX = 1, double scaleY = 1)
    {
        return element switch
        {
            P.Shape shape => ConvertShape(slidePart, shape, styleResolver, slideIndex, offX, offY, scaleX, scaleY),
            P.Picture picture => ConvertPicture(slidePart, picture, offX, offY, scaleX, scaleY),
            P.GraphicFrame graphicFrame => ConvertGraphicFrame(slidePart, graphicFrame, styleResolver, offX, offY, scaleX, scaleY),
            P.GroupShape groupShape => ConvertGroupShape(slidePart, groupShape, styleResolver, slideIndex, offX, offY, scaleX, scaleY),
            _ => Array.Empty<TypstElement>()
        };
    }

    private IEnumerable<TypstElement> ConvertShape(SlidePart slidePart, P.Shape shape, StyleResolver styleResolver, int slideIndex, double offX, double offY, double scaleX, double scaleY)
    {
        var (id, name) = GetElementIdAndName(shape.NonVisualShapeProperties);
        var position = GetElementPosition(shape.ShapeProperties);
        
        // If no transform on slide, inherit from layout placeholder
        if (position.X == 0 && position.Y == 0 && position.Width == 100 && position.Height == 50)
        {
            var layoutPosition = GetLayoutPlaceholderPosition(slidePart, shape);
            if (layoutPosition != null)
            {
                position = layoutPosition.Value;
            }
        }
        
        var text = ExtractTextFromShape(shape, styleResolver, slidePart);
        
        // Handle slide number placeholder
        var placeholderType = GetPlaceholderType(shape);
        if (placeholderType == PlaceholderValues.SlideNumber)
        {
            text = new TypstTextElement
            {
                Paragraphs = new()
                {
                    new TypstParagraph
                    {
                        Content = slideIndex.ToString(),
                        Runs = new() { new TypstTextRun { Content = slideIndex.ToString(), Formatting = text.Formatting } },
                        Formatting = text.Formatting
                    }
                },
                AutoFit = text.AutoFit,
                VerticalAlign = text.VerticalAlign,
                PaddingLeft = text.PaddingLeft,
                PaddingTop = text.PaddingTop,
                PaddingRight = text.PaddingRight,
                PaddingBottom = text.PaddingBottom,
                LineSpacing = text.LineSpacing,
                ParagraphCount = text.ParagraphCount,
                HasExplicitLineBreaks = text.HasExplicitLineBreaks
            };
        }

        // Apply group transform
        var finalX = offX + position.X * scaleX;
        var finalY = offY + position.Y * scaleY;
        var finalW = position.Width * scaleX;
        var finalH = position.Height * scaleY;
        var finalRot = position.Rotation;

        // Check for image fill
        var blipFill = shape.ShapeProperties?.Elements<Drawing.BlipFill>().FirstOrDefault();
        if (blipFill != null)
        {
            var imageElement = ExtractImageFromBlipFill(slidePart, blipFill);
            if (imageElement != null)
            {
                // Extract corner radius from shape geometry for rounded image clipping
                if (shape.ShapeProperties != null)
                {
                    imageElement.CornerRadius = ExtractShapeCornerRadius(shape.ShapeProperties, finalW, finalH);
                }

                // If shape also has text, we should ideally overlay it
                // For now, return image if no text, or prioritize text if present
                if (string.IsNullOrWhiteSpace(text.Content))
                {
                yield return new TypstElement
                {
                    Type = "Image",
                    Id = id,
                    Name = name,
                    X = finalX,
                    Y = finalY,
                    Width = finalW,
                    Height = finalH,
                    Rotation = finalRot,
                    Image = imageElement
                };
                    yield break;
                }
            }
        }

        // Check for shape geometry with fill or stroke
        var shapeElement = ExtractShapeGeometry(shape.ShapeProperties, styleResolver, finalW, finalH);
        if (shapeElement != null && (!string.IsNullOrEmpty(shapeElement.FillColor)
            || shapeElement.FillGradient != null
            || (!string.IsNullOrEmpty(shapeElement.StrokeColor) && shapeElement.StrokeWidth > 0)))
        {
            // Return shape element
            yield return new TypstElement
            {
                Type = "Shape",
                Id = id,
                Name = name,
                X = finalX,
                Y = finalY,
                Width = finalW,
                Height = finalH,
                Rotation = finalRot,
                Shape = shapeElement
            };
        }

        // Set original text box height (before padding)
        text.TextBoxHeight = finalH;

        // Return text if present
        if (!string.IsNullOrWhiteSpace(text.Content))
        {
            yield return new TypstElement
            {
                Type = "Text",
                Id = id,
                Name = name,
                X = finalX,
                Y = finalY,
                Width = finalW,
                Height = finalH,
                Rotation = finalRot,
                Text = text
            };
        }
    }

    private TypstShapeElement? ExtractShapeGeometry(ShapeProperties? shapeProperties, StyleResolver? styleResolver = null, double shapeWidth = 0, double shapeHeight = 0)
    {
        if (shapeProperties == null) return null;

        // Extract fill color (solid), falling back to a linear gradient fill when present
        var fillColor = ExtractShapeFillColor(shapeProperties, styleResolver);
        var fillGradient = string.IsNullOrEmpty(fillColor)
            ? ExtractShapeFillGradient(shapeProperties, styleResolver)
            : null;

        // Extract stroke (outline) properties
        var (strokeColor, strokeWidth) = ExtractShapeStroke(shapeProperties, styleResolver);

        // Helper to build TypstShapeElement with common fill+stroke properties
        TypstShapeElement CreateElement(string shapeType) => new()
        {
            ShapeType = shapeType,
            FillColor = fillColor,
            FillGradient = fillGradient,
            StrokeColor = strokeColor,
            StrokeWidth = strokeWidth
        };

        // Check for preset geometry
        var prstGeom = shapeProperties.Elements<Drawing.PresetGeometry>().FirstOrDefault();
        if (prstGeom != null)
        {
            var prst = prstGeom.Preset?.Value;
            if (prst == Drawing.ShapeTypeValues.Rectangle || prst == Drawing.ShapeTypeValues.RoundRectangle)
            {
                var cornerRadius = prst == Drawing.ShapeTypeValues.RoundRectangle
                    ? ExtractShapeCornerRadius(shapeProperties, shapeWidth, shapeHeight)
                    : 0;
                return new TypstShapeElement
                {
                    ShapeType = "rect",
                    FillColor = fillColor,
                    FillGradient = fillGradient,
                    StrokeColor = strokeColor,
                    StrokeWidth = strokeWidth,
                    CornerRadius = cornerRadius
                };
            }
            if (prst == Drawing.ShapeTypeValues.Ellipse)
            {
                return CreateElement("ellipse");
            }
        }

        // Check for custom geometry (path-based shapes)
        var custGeom = shapeProperties.Elements<Drawing.CustomGeometry>().FirstOrDefault();
        if (custGeom != null)
        {
            var pathList = custGeom.Elements<Drawing.PathList>().FirstOrDefault();
            if (pathList != null)
            {
                var path = pathList.Elements<Drawing.Path>().FirstOrDefault();
                if (path != null)
                {
                    var points = ExtractPathPoints(path);
                    if (points.Count > 2)
                    {
                        // Check if it's a simple rectangle (4 points + close)
                        if (IsRectanglePath(points))
                        {
                            return CreateElement("rect");
                        }

                        // Otherwise treat as polygon
                        return new TypstShapeElement
                        {
                            ShapeType = "polygon",
                            FillColor = fillColor,
                            FillGradient = fillGradient,
                            StrokeColor = strokeColor,
                            StrokeWidth = strokeWidth,
                            Points = points
                        };
                    }
                }
            }
        }

        // No recognizable geometry — return shape if it has fill or stroke
        if (!string.IsNullOrEmpty(fillColor) || fillGradient != null || (!string.IsNullOrEmpty(strokeColor) && strokeWidth > 0))
        {
            // Fallback: treat as rectangle
            return CreateElement("rect");
        }

        return null;
    }

    private string ExtractShapeFillColor(ShapeProperties shapeProperties, StyleResolver? styleResolver = null)
    {
        var solidFill = shapeProperties.Elements<Drawing.SolidFill>().FirstOrDefault();
        if (solidFill != null)
        {
            var color = ExtractColor(solidFill, styleResolver);
            if (!string.IsNullOrEmpty(color))
            {
                return color;
            }
        }
        return string.Empty;
    }

    /// <summary>
    /// Extracts a linear gradient fill (&lt;a:gradFill&gt;) from shape properties, including
    /// per-stop &lt;a:alpha&gt; opacities. Returns null when there is no gradient, when it has
    /// fewer than two resolvable stops, or when a stop color cannot be resolved.
    /// Decks like AetherLink use full-bleed gradient rectangles as slide backgrounds.
    /// </summary>
    private TypstGradientFill? ExtractShapeFillGradient(ShapeProperties shapeProperties, StyleResolver? styleResolver)
    {
        var gradFill = shapeProperties.Elements<Drawing.GradientFill>().FirstOrDefault();
        if (gradFill?.GradientStopList is not { } stopList)
            return null;

        var stops = new List<TypstGradientStop>();
        foreach (var gs in stopList.Elements<Drawing.GradientStop>())
        {
            var color = ExtractGradientStopColor(gs, styleResolver);
            if (color == null)
                return null;

            var pos = gs.Position?.Value ?? 0;
            stops.Add(new TypstGradientStop(color, Math.Clamp(pos / 100000.0, 0.0, 1.0)));
        }

        if (stops.Count < 2)
            return null;

        // OOXML a:lin ang is in 60000ths of a degree; OOXML and Typst gradient.linear both
        // measure the axis clockwise from the left→right direction, so degrees pass through
        // verbatim (same convention as the generation pipeline's emitters).
        var angle = (gradFill.Elements<Drawing.LinearGradientFill>().FirstOrDefault()?.Angle?.Value ?? 0) / 60000.0;
        return new TypstGradientFill(angle, stops);
    }

    private static string? ExtractGradientStopColor(Drawing.GradientStop stop, StyleResolver? styleResolver)
    {
        var rgb = stop.Elements<Drawing.RgbColorModelHex>().FirstOrDefault();
        if (rgb?.Val?.Value is { } hex)
        {
            // Unlike a solid fill, a fully transparent gradient stop is meaningful
            // (fade-out) — keep it as #RRGGBB00 instead of treating it as noFill.
            return ApplyColorModifiers(ParseHexColor(hex), rgb.ChildElements)
                ?? FormatHexColor(ParseHexColor(hex), 0);
        }

        var schemeColor = stop.Elements<Drawing.SchemeColor>().FirstOrDefault();
        if (schemeColor != null)
        {
            // Raw XML attribute read per AGENTS.pptx.md rule 1 — SDK enum parsing is unreliable.
            var match = Regex.Match(schemeColor.OuterXml, @"val=""([^""]+)""");
            if (match.Success)
            {
                var resolved = styleResolver?.ResolveSchemeColor(match.Groups[1].Value);
                if (!string.IsNullOrEmpty(resolved))
                {
                    return ApplyColorModifiers(ParseHexColor(resolved), schemeColor.ChildElements)
                        ?? FormatHexColor(ParseHexColor(resolved), 0);
                }
            }
        }

        return null;
    }

    /// <summary>
    /// Extracts stroke (outline) properties from a shape's &lt;a:ln&gt; element.
    /// Returns the stroke color (hex with # prefix) and width in points (EMU / 12700).
    /// </summary>
    private (string StrokeColor, double StrokeWidth) ExtractShapeStroke(ShapeProperties shapeProperties, StyleResolver? styleResolver = null)
    {
        var outline = shapeProperties.Elements<Drawing.Outline>().FirstOrDefault();
        if (outline == null)
            return (string.Empty, 0);

        // Check for noFill — outline exists but no color fill
        var noFill = outline.Elements<Drawing.NoFill>().FirstOrDefault();
        if (noFill != null)
            return (string.Empty, 0);

        var strokeWidth = outline.Width?.Value / 12700.0 ?? 0;
        if (strokeWidth <= 0)
            return (string.Empty, 0);

        var solidFill = outline.Elements<Drawing.SolidFill>().FirstOrDefault();
        if (solidFill != null)
        {
            var color = ExtractColor(solidFill, styleResolver);
            if (!string.IsNullOrEmpty(color))
                return (color, strokeWidth);
        }

        return (string.Empty, strokeWidth);
    }

    /// <summary>
    /// Extracts the corner radius from a RoundRectangle preset geometry's adjustment value list.
    /// The adjustment value (0–100000) represents a percentage of the shape's smaller dimension.
    /// Falls back to 5% of the smaller dimension if no adjustment value is found.
    /// </summary>
    private double ExtractShapeCornerRadius(ShapeProperties shapeProperties, double width, double height)
    {
        if (width <= 0 || height <= 0) return 0;

        var prstGeom = shapeProperties.Elements<Drawing.PresetGeometry>().FirstOrDefault();
        if (prstGeom == null) return 0;

        // Parse the adjustment value from <a:gd name="adj" fmla="val XXXX"/>
        var match = Regex.Match(prstGeom.OuterXml, @"\bfmla\s*=\s*""val\s+(\d+)""");
        if (match.Success && double.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var adjValue))
        {
            return (adjValue / 100000.0) * Math.Min(width, height);
        }

        // Fallback: 5% of the shape's smaller dimension
        return Math.Min(width, height) * 0.05;
    }

    private List<(double X, double Y)> ExtractPathPoints(Drawing.Path path)
    {
        var points = new List<(double X, double Y)>();
        var width = (double)(path.Width?.Value ?? 1);
        var height = (double)(path.Height?.Value ?? 1);
        var currentX = 0.0;
        var currentY = 0.0;

        foreach (var cmd in path.ChildElements)
        {
            switch (cmd)
            {
                case Drawing.MoveTo moveTo:
                    (currentX, currentY) = GetPoint(moveTo.Point, width, height);
                    points.Add((currentX, currentY));
                    break;

                case Drawing.LineTo lineTo:
                    (currentX, currentY) = GetPoint(lineTo.Point, width, height);
                    points.Add((currentX, currentY));
                    break;

                case Drawing.CubicBezierCurveTo cubicBez:
                    // cubicBezTo contains 3 points: control1, control2, end
                    var bezPoints = cubicBez.Elements<Drawing.Point>().ToList();
                    if (bezPoints.Count >= 3)
                    {
                        var (c1x, c1y) = GetPoint(bezPoints[0], width, height);
                        var (c2x, c2y) = GetPoint(bezPoints[1], width, height);
                        var (ex, ey) = GetPoint(bezPoints[2], width, height);

                        // Sample 8 points along the bezier curve
                        for (int i = 1; i <= 8; i++)
                        {
                            double t = i / 9.0;
                            double mt = 1 - t;
                            double mt2 = mt * mt;
                            double mt3 = mt2 * mt;
                            double t2 = t * t;
                            double t3 = t2 * t;

                            double bx = mt3 * currentX + 3 * mt2 * t * c1x + 3 * mt * t2 * c2x + t3 * ex;
                            double by = mt3 * currentY + 3 * mt2 * t * c1y + 3 * mt * t2 * c2y + t3 * ey;
                            points.Add((bx, by));
                        }

                        currentX = ex;
                        currentY = ey;
                    }
                    break;

                case Drawing.CloseShapePath:
                    // Don't add duplicate close point
                    break;
            }
        }

        return points;
    }

    private (double X, double Y) GetPoint(Drawing.Point? point, double width, double height)
    {
        if (point == null) return (0, 0);
        var xStr = point.X?.Value ?? "0";
        var yStr = point.Y?.Value ?? "0";
        var x = double.TryParse(xStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var xp) ? xp : 0;
        var y = double.TryParse(yStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var yp) ? yp : 0;
        // Normalize to 0-1 range
        return (x / width, y / height);
    }

    private bool IsRectanglePath(List<(double X, double Y)> points)
    {
        if (points.Count < 4) return false;
        // Check if points form axis-aligned rectangle (with some tolerance)
        var distinctXs = points.Select(p => Math.Round(p.X, 3)).Distinct().ToList();
        var distinctYs = points.Select(p => Math.Round(p.Y, 3)).Distinct().ToList();
        return distinctXs.Count == 2 && distinctYs.Count == 2;
    }

    private IEnumerable<TypstElement> ConvertPicture(SlidePart slidePart, P.Picture picture, double offX, double offY, double scaleX, double scaleY)
    {
        var (id, name) = GetElementIdAndName(picture.NonVisualPictureProperties);
        var position = GetElementPosition(picture.ShapeProperties);

        var imageElement = ExtractImage(slidePart, picture);
        if (imageElement == null) yield break;

        yield return new TypstElement
        {
            Type = "Image",
            Id = id,
            Name = name,
            X = offX + position.X * scaleX,
            Y = offY + position.Y * scaleY,
            Width = position.Width * scaleX,
            Height = position.Height * scaleY,
            Rotation = position.Rotation,
            Image = imageElement
        };
    }

    private IEnumerable<TypstElement> ConvertGraphicFrame(SlidePart slidePart, P.GraphicFrame graphicFrame, StyleResolver styleResolver, double offX, double offY, double scaleX, double scaleY)
    {
        var (id, name) = GetElementIdAndName(graphicFrame.NonVisualGraphicFrameProperties);
        var position = GetGraphicFramePosition(graphicFrame);

        var graphicData = graphicFrame.Graphic?.GraphicData;
        if (graphicData == null) yield break;

        var table = graphicData.Elements<Drawing.Table>().FirstOrDefault();
        if (table != null)
        {
            var tableElement = ExtractTable(table, styleResolver);
            yield return new TypstElement
            {
                Type = "Table",
                Id = id,
                Name = name,
                X = offX + position.X * scaleX,
                Y = offY + position.Y * scaleY,
                Width = position.Width * scaleX,
                Height = position.Height * scaleY,
                Table = tableElement
            };
            yield break;
        }

        var uri = graphicData.Uri?.Value ?? "";
        if (uri.Contains("/drawingml/2006/diagram", StringComparison.Ordinal))
        {
            // SmartArt: extract positioned shapes and text from the diagram drawing part.
            var hadShapes = false;
            var hadText = false;
            foreach (var element in ConvertDiagramGraphicFrame(slidePart, graphicFrame, position, offX, offY, scaleX, scaleY, styleResolver))
            {
                hadShapes = hadShapes || element.Type == "Shape";
                hadText = hadText || element.Type == "Text";
                yield return element;
            }

            if (hadShapes && hadText)
            {
                AddSlideWarning($"SmartArt diagram '{name}' was rendered from pre-rendered shapes; theme colours may differ from the original.");
            }
            else if (hadShapes)
            {
                AddSlideWarning($"SmartArt diagram '{name}' was rendered from pre-rendered shapes (text not found); theme colours may differ from the original.");
            }
            else if (hadText)
            {
                AddSlideWarning($"SmartArt diagram '{name}' was approximated as positioned text; diagram layout and styling may differ from the original.");
            }
            else
            {
                AddSlideWarning($"SmartArt diagram '{name}' could not be approximated and was replaced by a placeholder.");
                foreach (var placeholder in CreateUnsupportedFramePlaceholders(position, offX, offY, scaleX, scaleY, "SmartArt diagram"))
                    yield return placeholder;
            }

            yield break;
        }

        // Charts, OLE objects, media, and any other graphic frames have no conversion
        // path: render a visible placeholder (never drop content silently) and warn.
        var kind = ClassifyGraphicFrameKind(uri);
        AddSlideWarning($"{kind} '{name}' is not supported and was replaced by a placeholder.");
        foreach (var placeholder in CreateUnsupportedFramePlaceholders(position, offX, offY, scaleX, scaleY, kind))
            yield return placeholder;
    }

    private static string ClassifyGraphicFrameKind(string uri)
    {
        if (uri.Contains("/drawingml/2006/chart", StringComparison.Ordinal))
            return "Chart";
        if (uri.Contains("oleObject", StringComparison.OrdinalIgnoreCase))
            return "Embedded object";
        if (uri.Contains("video", StringComparison.OrdinalIgnoreCase)
            || uri.Contains("audio", StringComparison.OrdinalIgnoreCase)
            || uri.Contains("media", StringComparison.OrdinalIgnoreCase))
            return "Media";
        return "Unsupported content";
    }

    /// <summary>
    /// Builds the visible placeholder for a graphic frame that cannot be rendered: a
    /// framed box with a centered label. Two elements (box + label) at the frame's
    /// position; empty when the frame has no usable extent.
    /// </summary>
    private static IEnumerable<TypstElement> CreateUnsupportedFramePlaceholders(
        (double X, double Y, double Width, double Height) position,
        double offX, double offY, double scaleX, double scaleY,
        string kind)
    {
        var x = offX + position.X * scaleX;
        var y = offY + position.Y * scaleY;
        var width = position.Width * scaleX;
        var height = position.Height * scaleY;
        if (width <= 0 || height <= 0)
            yield break;

        yield return new TypstElement
        {
            Type = "Shape",
            X = x,
            Y = y,
            Width = width,
            Height = height,
            Shape = new TypstShapeElement
            {
                ShapeType = "rect",
                FillColor = "#F5F5F5",
                StrokeColor = "#9E9E9E",
                StrokeWidth = 1.0
            }
        };

        var label = $"{kind} (not supported)";
        var formatting = new TypstTextFormatting
        {
            FontSize = 10,
            Color = "#757575",
            Align = "center"
        };
        yield return new TypstElement
        {
            Type = "Text",
            X = x,
            Y = y,
            Width = width,
            Height = height,
            Text = new TypstTextElement
            {
                Paragraphs = new List<TypstParagraph>
                {
                    new TypstParagraph
                    {
                        Content = label,
                        Runs = new List<TypstTextRun> { new TypstTextRun { Content = label, Formatting = formatting } },
                        Formatting = formatting
                    }
                },
                VerticalAlign = "center",
                ParagraphCount = 1
            }
        };
    }

    private IEnumerable<TypstElement> ConvertDiagramGraphicFrame(SlidePart slidePart, P.GraphicFrame graphicFrame,
        (double X, double Y, double Width, double Height) framePosition,
        double offX, double offY, double scaleX, double scaleY, StyleResolver styleResolver)
    {
        var graphicData = graphicFrame.Graphic?.GraphicData;
        if (graphicData == null) yield break;

        // Resolve the drawing part associated with THIS graphic frame. The drawing
        // part is not referenced from dgm:relIds (which carries only dm/lo/qs/cs),
        // so a brute-force "first part containing dsp:sp" scan cross-renders when a
        // slide hosts 2+ diagrams.
        var drawingPart = ResolveDiagramDrawingPart(slidePart, graphicData);
        if (drawingPart == null) yield break;

        OpenXmlElement? root;
        try
        {
            root = drawingPart.RootElement;
        }
        catch
        {
            yield break;
        }

        if (root == null) yield break;

        var diagramShapes = root.Descendants()
            .Where(e => e.LocalName == "sp" && _diagramNamespaces.Contains(e.NamespaceUri))
            .ToList();

        if (diagramShapes.Count == 0) yield break;

        // Drawing-space → frame normalisation: drawing shape coordinates do not
        // necessarily span the frame extents (e.g. the REF deck's diagram content
        // covers ~83% of the frame height), so scale/offset the drawing's bounding
        // box onto the frame instead of anchoring it top-left at native size.
        //
        // TryExtractShape computes  final = off + (frame + shapeOff) * scale,  so the
        // normalisation is folded into an adjusted frame origin + scale such that
        //   (frameX/ds − min + shapeOff) · ds·scale = (frameX + (shapeOff − min)·ds) · scale
        // i.e. the drawing bbox is mapped onto the frame before the slide transform.
        double drawScaleX = 1.0, drawScaleY = 1.0;
        double shapeFrameX = framePosition.X, shapeFrameY = framePosition.Y;
        var bounds = SmartArtDrawingExtractor.ComputeBoundingBox(diagramShapes);
        if (bounds is { } b && b.Width > 0 && b.Height > 0 &&
            framePosition.Width > 0 && framePosition.Height > 0)
        {
            drawScaleX = framePosition.Width / b.Width;
            drawScaleY = framePosition.Height / b.Height;
            shapeFrameX = framePosition.X / drawScaleX - b.MinX;
            shapeFrameY = framePosition.Y / drawScaleY - b.MinY;
        }

        var shapeScaleX = scaleX * drawScaleX;
        var shapeScaleY = scaleY * drawScaleY;

        foreach (var shape in diagramShapes)
        {
            // Geometry + dimensions always come from spPr/xfrm; txXfrm is the text
            // placement box and is only used for the text element below.
            var geometry = GetDiagramShapeGeometry(shape);
            if (geometry == null) continue;

            var (_, _, shapeW, shapeH) = geometry.Value;
            var modelId = ReadDiagramShapeModelId(shape);

            var diagramShape = SmartArtDrawingExtractor.TryExtractShape(
                shape, offX, offY, shapeScaleX, shapeScaleY,
                shapeFrameX, shapeFrameY,
                shapeW, shapeH,
                styleResolver.SchemeColors,
                modelId);

            if (diagramShape != null)
                yield return diagramShape;

            var textBounds = GetDiagramTextBounds(shape);
            if (textBounds == null) continue;

            var textElement = ExtractTextFromDiagramShape(shape, styleResolver.SchemeColors);
            if (textElement == null || string.IsNullOrWhiteSpace(textElement.Content))
                continue;

            var (tx, ty, tw, th) = textBounds.Value;
            yield return new TypstElement
            {
                Type = "Text",
                X = offX + (shapeFrameX + tx) * shapeScaleX,
                Y = offY + (shapeFrameY + ty) * shapeScaleY,
                Width = tw * shapeScaleX,
                Height = th * shapeScaleY,
                ModelId = modelId,
                Text = textElement
            };
        }
    }

    /// <summary>
    /// Resolves the <see cref="DiagramPersistLayoutPart"/> that belongs to the given
    /// diagram graphic frame: r:dm → data part, then the data part's
    /// <c>dsp:dataModelExt relId</c> → drawing part. Falls back to a drawing part
    /// that declares a relationship to the data part, and finally to the slide's
    /// single drawing part when unambiguous. Returns null when no association can be
    /// established (never guesses among multiple candidates).
    /// </summary>
    private static DiagramPersistLayoutPart? ResolveDiagramDrawingPart(SlidePart slidePart, OpenXmlElement graphicData)
    {
        var dataPart = ResolveDiagramDataPart(slidePart, graphicData);

        if (dataPart != null)
        {
            var drawingRelId = ReadDataModelExtRelId(dataPart);
            if (!string.IsNullOrEmpty(drawingRelId))
            {
                var resolved = TryGetPartById(dataPart, drawingRelId!)
                               ?? TryGetPartById(slidePart, drawingRelId!);
                if (resolved is DiagramPersistLayoutPart persistPart)
                    return persistPart;
            }

            // Fallback: a drawing part whose related parts include this data part.
            foreach (var candidate in slidePart.GetPartsOfType<DiagramPersistLayoutPart>())
            {
                if (candidate.Parts.Any(p => p.OpenXmlPart == dataPart))
                    return candidate;
            }
        }

        // Last resort: a single drawing part on the slide is unambiguous.
        var drawingParts = slidePart.GetPartsOfType<DiagramPersistLayoutPart>().ToList();
        return drawingParts.Count == 1 ? drawingParts[0] : null;
    }

    // dgm:relIds lives in the drawingml diagram namespace (the older
    // "drawing/2006/diagram" variant seen in some files is also accepted).
    private static readonly HashSet<string> _diagramRelNamespaces = new(StringComparer.Ordinal)
    {
        "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        "http://schemas.openxmlformats.org/drawing/2006/diagram"
    };

    private static OpenXmlPart? ResolveDiagramDataPart(SlidePart slidePart, OpenXmlElement graphicData)
    {
        var relIdsElement = graphicData.Elements()
            .FirstOrDefault(e => e.LocalName == "relIds" && _diagramRelNamespaces.Contains(e.NamespaceUri));
        if (relIdsElement == null) return null;

        var dmRelId = relIdsElement.GetAttributes()
            .FirstOrDefault(a => a.LocalName == "dm" && !string.IsNullOrEmpty(a.Value))
            .Value;
        if (string.IsNullOrEmpty(dmRelId)) return null;

        var part = TryGetPartById(slidePart, dmRelId);
        if (part == null) return null;

        // Verify it really is the data part (dgm:dataModel root).
        try
        {
            var root = part.RootElement;
            if (root != null && root.LocalName == "dataModel" &&
                root.NamespaceUri == "http://schemas.openxmlformats.org/drawingml/2006/diagram")
            {
                return part;
            }
        }
        catch
        {
            // Part not readable — treat as unresolved.
        }

        return null;
    }

    private static string? ReadDataModelExtRelId(OpenXmlPart dataPart)
    {
        try
        {
            var ext = dataPart.RootElement?.Descendants()
                .FirstOrDefault(e => e.LocalName == "dataModelExt");
            if (ext == null) return null;

            var relId = ext.GetAttributes()
                .FirstOrDefault(a => a.LocalName == "relId" && !string.IsNullOrEmpty(a.Value))
                .Value;
            return string.IsNullOrEmpty(relId) ? null : relId;
        }
        catch
        {
            return null;
        }
    }

    private static OpenXmlPart? TryGetPartById(OpenXmlPart container, string relationshipId)
    {
        try
        {
            return container.GetPartById(relationshipId);
        }
        catch
        {
            return null;
        }
    }

    private static string? ReadDiagramShapeModelId(OpenXmlElement diagramShape)
    {
        var match = System.Text.RegularExpressions.Regex.Match(
            diagramShape.OuterXml,
            @"\bmodelId\s*=\s*""([^""]*)""");
        return match.Success && !string.IsNullOrEmpty(match.Groups[1].Value)
            ? match.Groups[1].Value
            : null;
    }

    private static readonly HashSet<string> _diagramNamespaces = new(StringComparer.Ordinal)
    {
        "http://schemas.openxmlformats.org/drawing/2006/diagram",
        "http://schemas.microsoft.com/office/drawing/2008/diagram"
    };

    private TypstTextElement? ExtractTextFromDiagramShape(OpenXmlElement diagramShape,
        IReadOnlyDictionary<string, string>? schemeColors)
    {
        var txBody = diagramShape.Elements()
            .FirstOrDefault(e => e.LocalName == "txBody" && _diagramNamespaces.Contains(e.NamespaceUri));

        if (txBody == null) return null;

        var text = ExtractTextFromTextBody(txBody);

        // dsp:style/a:fontRef supplies the default text colour for diagram shapes
        // (e.g. <a:fontRef idx="minor"><a:schemeClr val="lt1"/> = white text on
        // accent-filled boxes). Apply it wherever a paragraph/run still carries the
        // unresolved "#000000" cascade default — explicit run colours win.
        var fontRefColor = ResolveDiagramFontRefColor(diagramShape, schemeColors);
        if (fontRefColor == null) return text;

        foreach (var paragraph in text.Paragraphs)
        {
            if (paragraph.Formatting.Color == "#000000")
                paragraph.Formatting = paragraph.Formatting with { Color = fontRefColor };

            foreach (var run in paragraph.Runs)
            {
                if (run.Formatting.Color == "#000000")
                    run.Formatting = run.Formatting with { Color = fontRefColor };
            }
        }

        return text;
    }

    private static string? ResolveDiagramFontRefColor(OpenXmlElement diagramShape,
        IReadOnlyDictionary<string, string>? schemeColors)
    {
        const string drawingmlNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

        var style = diagramShape.Elements()
            .FirstOrDefault(e => e.LocalName == "style" && _diagramNamespaces.Contains(e.NamespaceUri));
        var fontRef = style?.Elements()
            .FirstOrDefault(e => e.LocalName == "fontRef" && e.NamespaceUri == drawingmlNs);
        if (fontRef == null) return null;

        var schemeClr = fontRef.Elements()
            .FirstOrDefault(e => e.LocalName == "schemeClr" && e.NamespaceUri == drawingmlNs);
        if (schemeClr != null)
        {
            var match = System.Text.RegularExpressions.Regex.Match(schemeClr.OuterXml, @"\bval\s*=\s*""([^""]*)""");
            if (match.Success && !string.IsNullOrEmpty(match.Groups[1].Value))
                return SmartArtDrawingExtractor.ResolveSchemeColor(match.Groups[1].Value, schemeColors);
        }

        var srgbClr = fontRef.Elements()
            .FirstOrDefault(e => e.LocalName == "srgbClr" && e.NamespaceUri == drawingmlNs);
        if (srgbClr != null)
        {
            var match = System.Text.RegularExpressions.Regex.Match(srgbClr.OuterXml, @"\bval\s*=\s*""([^""]*)""");
            if (match.Success && !string.IsNullOrEmpty(match.Groups[1].Value))
                return "#" + match.Groups[1].Value;
        }

        return null;
    }

    private static Drawing.BodyProperties? GetCascadedBodyPr(
        OpenXmlElement? slideTextBody,
        StyleResolver? styleResolver,
        int? placeholderIdx,
        PlaceholderValues? placeholderType)
    {
        var slideBodyPr = slideTextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
        var layoutBodyPr = styleResolver?.GetLayoutPlaceholderBodyPr(placeholderIdx, placeholderType);
        var masterBodyPr = styleResolver?.GetMasterPlaceholderBodyPr(placeholderIdx, placeholderType);

        if (slideBodyPr == null && layoutBodyPr == null && masterBodyPr == null)
            return null;

        var result = new Drawing.BodyProperties();

        // Copy attributes from master -> layout -> slide
        var sources = new[] { masterBodyPr, layoutBodyPr, slideBodyPr };

        foreach (var source in sources)
        {
            if (source == null) continue;

            // Anchor - SDK doesn't parse this attribute, use regex
            var anchorMatch = System.Text.RegularExpressions.Regex.Match(source.OuterXml, @"anchor\s*=\s*""([^""]*)""");
            if (anchorMatch.Success)
            {
                var anchorStr = anchorMatch.Groups[1].Value;
                var anchorVal = anchorStr switch
                {
                    "t" => Drawing.TextAnchoringTypeValues.Top,
                    "ctr" => Drawing.TextAnchoringTypeValues.Center,
                    "b" => Drawing.TextAnchoringTypeValues.Bottom,
                    _ => (Drawing.TextAnchoringTypeValues?)null
                };
                if (anchorVal.HasValue)
                    result.Anchor = anchorVal.Value;
            }

            if (source.LeftInset?.Value != null)
                result.LeftInset = new Int32Value(source.LeftInset.Value);
            if (source.TopInset?.Value != null)
                result.TopInset = new Int32Value(source.TopInset.Value);
            if (source.RightInset?.Value != null)
                result.RightInset = new Int32Value(source.RightInset.Value);
            if (source.BottomInset?.Value != null)
                result.BottomInset = new Int32Value(source.BottomInset.Value);

            if (source.Rotation?.Value != null)
                result.Rotation = new Int32Value(source.Rotation.Value);

            if (source.Wrap?.Value != null)
                result.Wrap = source.Wrap.Value;
        }

        // Copy autofit child elements from slide -> layout -> master
        var noAutoFit = slideBodyPr?.Elements<Drawing.NoAutoFit>().FirstOrDefault()
            ?? layoutBodyPr?.Elements<Drawing.NoAutoFit>().FirstOrDefault()
            ?? masterBodyPr?.Elements<Drawing.NoAutoFit>().FirstOrDefault();
        var normalAutoFit = slideBodyPr?.Elements<Drawing.NormalAutoFit>().FirstOrDefault()
            ?? layoutBodyPr?.Elements<Drawing.NormalAutoFit>().FirstOrDefault()
            ?? masterBodyPr?.Elements<Drawing.NormalAutoFit>().FirstOrDefault();
        var shapeAutoFit = slideBodyPr?.Elements<Drawing.ShapeAutoFit>().FirstOrDefault()
            ?? layoutBodyPr?.Elements<Drawing.ShapeAutoFit>().FirstOrDefault()
            ?? masterBodyPr?.Elements<Drawing.ShapeAutoFit>().FirstOrDefault();

        if (noAutoFit != null)
            result.Append(noAutoFit.CloneNode(true));
        else if (normalAutoFit != null)
            result.Append(normalAutoFit.CloneNode(true));
        else if (shapeAutoFit != null)
            result.Append(shapeAutoFit.CloneNode(true));

        return result;
    }

    /// <summary>
    /// Reads shape geometry (position + size) from <c>dsp:spPr/a:xfrm</c> — the only
    /// authoritative source for the shape's drawing-space rectangle.
    /// </summary>
    private static (double X, double Y, double Width, double Height)? GetDiagramShapeGeometry(OpenXmlElement diagramShape)
    {
        var spPr = diagramShape.Elements()
            .FirstOrDefault(e => e.LocalName == "spPr" && _diagramNamespaces.Contains(e.NamespaceUri));
        if (spPr == null) return null;

        var xfrm = spPr.Elements()
            .FirstOrDefault(e => e.LocalName == "xfrm" &&
                e.NamespaceUri == "http://schemas.openxmlformats.org/drawingml/2006/main");

        return ReadDiagramXfrm(xfrm);
    }

    /// <summary>
    /// Reads the text placement box: <c>dsp:txXfrm</c> when present (Microsoft 2008
    /// diagram namespace), falling back to the shape geometry.
    /// </summary>
    private static (double X, double Y, double Width, double Height)? GetDiagramTextBounds(OpenXmlElement diagramShape)
    {
        var txXfrm = diagramShape.Elements()
            .FirstOrDefault(e => e.LocalName == "txXfrm" && _diagramNamespaces.Contains(e.NamespaceUri));

        return ReadDiagramXfrm(txXfrm) ?? GetDiagramShapeGeometry(diagramShape);
    }

    private static (double X, double Y, double Width, double Height)? ReadDiagramXfrm(OpenXmlElement? xfrm)
    {
        if (xfrm == null) return null;

        var off = xfrm.Elements()
            .FirstOrDefault(e => e.LocalName == "off" &&
                e.NamespaceUri == "http://schemas.openxmlformats.org/drawingml/2006/main");
        var ext = xfrm.Elements()
            .FirstOrDefault(e => e.LocalName == "ext" &&
                e.NamespaceUri == "http://schemas.openxmlformats.org/drawingml/2006/main");

        if (off == null || ext == null) return null;

        var xAttr = off.GetAttribute("x", "");
        var yAttr = off.GetAttribute("y", "");
        var cxAttr = ext.GetAttribute("cx", "");
        var cyAttr = ext.GetAttribute("cy", "");

        if (string.IsNullOrEmpty(xAttr.Value) || string.IsNullOrEmpty(yAttr.Value) ||
            string.IsNullOrEmpty(cxAttr.Value) || string.IsNullOrEmpty(cyAttr.Value))
            return null;

        if (!long.TryParse(xAttr.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var xEmu) ||
            !long.TryParse(yAttr.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var yEmu) ||
            !long.TryParse(cxAttr.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var cxEmu) ||
            !long.TryParse(cyAttr.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var cyEmu))
            return null;

        return (EmuToPt(xEmu), EmuToPt(yEmu), EmuToPt(cxEmu), EmuToPt(cyEmu));
    }

    private IEnumerable<TypstElement> ConvertGroupShape(SlidePart slidePart, P.GroupShape groupShape, StyleResolver styleResolver, int slideIndex, double parentOffX, double parentOffY, double parentScaleX, double parentScaleY)
    {
        var grpXfrm = groupShape.GroupShapeProperties?.TransformGroup;

        double newOffX = parentOffX;
        double newOffY = parentOffY;
        double newScaleX = parentScaleX;
        double newScaleY = parentScaleY;

        if (grpXfrm != null)
        {
            var grpOffX = EmuToPt((long)(grpXfrm.Offset?.X?.Value ?? 0));
            var grpOffY = EmuToPt((long)(grpXfrm.Offset?.Y?.Value ?? 0));
            var grpExtX = (double)(grpXfrm.Extents?.Cx?.Value ?? 1);
            var grpExtY = (double)(grpXfrm.Extents?.Cy?.Value ?? 1);
            var chOffX = EmuToPt((long)(grpXfrm.ChildOffset?.X?.Value ?? 0));
            var chOffY = EmuToPt((long)(grpXfrm.ChildOffset?.Y?.Value ?? 0));
            var chExtX = (double)(grpXfrm.ChildExtents?.Cx?.Value ?? 1);
            var chExtY = (double)(grpXfrm.ChildExtents?.Cy?.Value ?? 1);

            var localScaleX = grpExtX / chExtX;
            var localScaleY = grpExtY / chExtY;

            newOffX = parentOffX + (grpOffX - chOffX) * parentScaleX;
            newOffY = parentOffY + (grpOffY - chOffY) * parentScaleY;
            newScaleX = parentScaleX * localScaleX;
            newScaleY = parentScaleY * localScaleY;
        }

        foreach (var child in groupShape.ChildElements)
        {
            foreach (var element in ConvertElement(slidePart, child, styleResolver, slideIndex, newOffX, newOffY, newScaleX, newScaleY))
                yield return element;
        }
    }

    private TypstTextElement ExtractTextFromShape(P.Shape shape, StyleResolver styleResolver, SlidePart? slidePart = null)
    {
        var textBody = shape.TextBody;
        if (textBody == null)
        {
            return new TypstTextElement();
        }

        var placeholderInfo = GetPlaceholderInfo(shape);
        var placeholderType = GetPlaceholderType(shape);
        
        // If slide shape has no explicit type, look up from layout/master by idx
        if (placeholderType == null && placeholderInfo?.Index.HasValue == true && slidePart != null)
        {
            placeholderType = GetPlaceholderTypeFromLayout(slidePart, placeholderInfo.Value.Index!.Value);
        }
        
        var result = ExtractTextFromTextBody(textBody, styleResolver, placeholderInfo?.Index, placeholderType);

        // Get default text style from master based on placeholder type
        var defaultStyle = styleResolver.GetDefaultTextStyle(placeholderType);

        // Apply defaults for missing values per-paragraph
        var updatedParagraphs = new List<TypstParagraph>();
        foreach (var paragraph in result.Paragraphs)
        {
            var oldFmt = paragraph.Formatting;
            var newFmt = oldFmt;

            if (newFmt.FontSize == 18.0 && defaultStyle.FontSize.HasValue)
                newFmt = newFmt with { FontSize = defaultStyle.FontSize.Value };
            if (!newFmt.Bold && defaultStyle.Bold.HasValue)
                newFmt = newFmt with { Bold = defaultStyle.Bold.Value };
            if (!newFmt.Italic && defaultStyle.Italic.HasValue)
                newFmt = newFmt with { Italic = defaultStyle.Italic.Value };
            if (newFmt.Color == "#000000" && !string.IsNullOrEmpty(defaultStyle.Color))
                newFmt = newFmt with { Color = defaultStyle.Color };
            if (newFmt.FontFamily == "Arial" && !string.IsNullOrEmpty(defaultStyle.FontFamily))
            {
                var fontName = defaultStyle.FontFamily;
                if (fontName.EndsWith(" Bold", StringComparison.OrdinalIgnoreCase))
                {
                    fontName = fontName.Substring(0, fontName.Length - 5);
                    newFmt = newFmt with { Bold = true };
                }
                newFmt = newFmt with { FontFamily = fontName };
            }

            // Propagate paragraph formatting changes to runs that inherited them
            var updatedRuns = new List<TypstTextRun>();
            foreach (var run in paragraph.Runs)
            {
                var runFmt = run.Formatting;
                if (runFmt.FontSize == oldFmt.FontSize)
                    runFmt = runFmt with { FontSize = newFmt.FontSize };
                if (runFmt.Bold == oldFmt.Bold)
                    runFmt = runFmt with { Bold = newFmt.Bold };
                if (runFmt.Italic == oldFmt.Italic)
                    runFmt = runFmt with { Italic = newFmt.Italic };
                if (runFmt.Color == oldFmt.Color)
                    runFmt = runFmt with { Color = newFmt.Color };
                if (runFmt.FontFamily == oldFmt.FontFamily)
                    runFmt = runFmt with { FontFamily = newFmt.FontFamily };

                updatedRuns.Add(new TypstTextRun { Content = run.Content, Formatting = runFmt });
            }

            updatedParagraphs.Add(new TypstParagraph
            {
                Content = paragraph.Content,
                Runs = updatedRuns,
                Formatting = newFmt,
                Level = paragraph.Level,
                BulletChar = paragraph.BulletChar,
                AutoNumberType = paragraph.AutoNumberType,
                HasBullet = paragraph.HasBullet,
                LineSpacing = paragraph.LineSpacing,
                SpaceBefore = paragraph.SpaceBefore,
                SpaceAfter = paragraph.SpaceAfter,
                MarginLeft = paragraph.MarginLeft,
                Indent = paragraph.Indent
            });
        }

        return new TypstTextElement
        {
            Paragraphs = updatedParagraphs,
            AutoFit = result.AutoFit,
            VerticalAlign = result.VerticalAlign,
            PaddingLeft = result.PaddingLeft,
            PaddingTop = result.PaddingTop,
            PaddingRight = result.PaddingRight,
            PaddingBottom = result.PaddingBottom,
            LineSpacing = result.LineSpacing,
            ParagraphCount = result.ParagraphCount,
            HasExplicitLineBreaks = result.HasExplicitLineBreaks
        };
    }

    private TypstTextElement ExtractTextFromTextBody(OpenXmlElement textBody)
    {
        return ExtractTextFromTextBody(textBody, null, null, null);
    }

    private TypstTextElement ExtractTextFromTextBody(OpenXmlElement textBody, StyleResolver? styleResolver, int? placeholderIdx, PlaceholderValues? placeholderType = null)
    {
        var paragraphs = new List<TypstParagraph>();
        string? align = null;
        var paragraphCount = 0;
        var hasExplicitLineBreaks = false;

        // Extract body properties (padding, auto-fit, anchor) with cascade
        var bodyPr = GetCascadedBodyPr(textBody, styleResolver, placeholderIdx, placeholderType);
        var autoFit = bodyPr?.Elements<Drawing.ShapeAutoFit>().FirstOrDefault() != null
            || bodyPr?.Elements<Drawing.NormalAutoFit>().FirstOrDefault() != null;
        var vertAlign = "top";
        if (bodyPr != null)
        {
            var anchorMatch = System.Text.RegularExpressions.Regex.Match(bodyPr.OuterXml, @"anchor\s*=\s*""([^""]*)""");
            if (anchorMatch.Success)
            {
                var anchorStr = anchorMatch.Groups[1].Value;
                if (anchorStr == "ctr")
                    vertAlign = "center";
                else if (anchorStr == "b")
                    vertAlign = "bottom";
            }
        }
        // OOXML bodyPr inset defaults (ECMA-376: lIns/rIns = 91440 EMU = 0.1",
        // tIns/bIns = 45720 EMU = 0.05") apply whenever an attribute is absent —
        // including when bodyPr itself is missing. Read via regex on OuterXml per
        // AGENTS.pptx.md rule 1.
        var padLeft = GetEmuAttributeAsPt(bodyPr, "lIns") ?? EmuToPt(DefaultHorizontalInsetEmu);
        var padTop = GetEmuAttributeAsPt(bodyPr, "tIns") ?? EmuToPt(DefaultVerticalInsetEmu);
        var padRight = GetEmuAttributeAsPt(bodyPr, "rIns") ?? EmuToPt(DefaultHorizontalInsetEmu);
        var padBottom = GetEmuAttributeAsPt(bodyPr, "bIns") ?? EmuToPt(DefaultVerticalInsetEmu);

        // Get text body list style for cascade level 3
        var bodyLstStyle = textBody.ChildElements.FirstOrDefault(e => e.LocalName == "lstStyle");

        foreach (var paragraph in textBody.Elements<Drawing.Paragraph>())
        {
            paragraphCount++;

            // Extract alignment from first paragraph
            if (align == null)
            {
                align = ExtractParagraphAlignment(paragraph);
            }

            var paragraphText = new StringBuilder();
            Drawing.Run? firstRun = null;
            var textRuns = new List<(string Text, Drawing.Run Run, TypstTextFormatting RawFmt)>();
            var elements = new List<(bool IsRun, Drawing.Run? Run, string Text)>();

            foreach (var child in paragraph.ChildElements)
            {
                if (child is Drawing.Run run)
                {
                    var runText = run.Text?.Text;
                    if (!string.IsNullOrEmpty(runText))
                    {
                        paragraphText.Append(runText);
                        textRuns.Add((runText, run, ExtractTextFormatting(run)));
                        elements.Add((true, run, runText));
                    }

                    if (firstRun == null)
                    {
                        firstRun = run;
                    }
                }

                if (child is Drawing.Break)
                {
                    paragraphText.Append('\n');
                    hasExplicitLineBreaks = true;
                    elements.Add((false, null, "\n"));
                }
            }

            var content = paragraphText.ToString();
            var level = 0;
            var pPr = paragraph.Elements<Drawing.ParagraphProperties>().FirstOrDefault();
            if (pPr?.Level?.Value != null)
            {
                level = pPr.Level.Value;
            }

            // Detect mixed formatting within the paragraph
            bool hasMixedFormatting = false;
            if (textRuns.Count > 1)
            {
                var firstRawFmt = textRuns[0].RawFmt;
                for (int i = 1; i < textRuns.Count; i++)
                {
                    if (!AreFormattingEqual(firstRawFmt, textRuns[i].RawFmt))
                    {
                        hasMixedFormatting = true;
                        break;
                    }
                }
            }

            // Resolve formatting through full cascade
            // When mixed formatting is present, don't let firstRun dictate paragraph defaults
            var formatting = ResolveParagraphFormatting(hasMixedFormatting ? null : firstRun, pPr, bodyLstStyle, level, styleResolver, placeholderIdx, placeholderType);
            if (align != null)
            {
                formatting = formatting with { Align = align };
            }

            // Build per-run formatting merged with paragraph defaults
            var runs = new List<TypstTextRun>();
            foreach (var (isRun, run, text) in elements)
            {
                if (isRun && run != null)
                {
                    var runFormatting = MergeRunWithParagraphDefaults(formatting, run, styleResolver);
                    runs.Add(new TypstTextRun { Content = text, Formatting = runFormatting });
                }
                else
                {
                    runs.Add(new TypstTextRun { Content = text, Formatting = formatting, IsLineBreak = text == "\n" });
                }
            }

            // Resolve bullet properties through full cascade
            var (bulletChar, autoNumberType, hasBullet) = ResolveBulletProperties(pPr, bodyLstStyle, level, styleResolver, placeholderIdx, placeholderType);

            // Resolve line spacing through full cascade
            var paragraphLineSpacing = ResolveLineSpacing(pPr, bodyLstStyle, level, styleResolver, placeholderIdx, placeholderType);

            // Resolve paragraph spacing (spcBef / spcAft) through full cascade
            var (spaceBefore, spaceAfter) = ResolveParagraphSpacing(pPr, bodyLstStyle, level, styleResolver, placeholderIdx, placeholderType);

            // Extract explicit margin left and indent from paragraph properties
            double? marginLeft = null;
            double? indent = null;
            if (pPr != null)
            {
                var marLAttr = GetAttributeValue(pPr, "marL");
                if (!string.IsNullOrEmpty(marLAttr) && int.TryParse(marLAttr, out var marL))
                    marginLeft = EmuToPt(marL);

                var indentAttr = GetAttributeValue(pPr, "indent");
                if (!string.IsNullOrEmpty(indentAttr) && int.TryParse(indentAttr, out var ind))
                    indent = EmuToPt(ind);
            }

            paragraphs.Add(new TypstParagraph
            {
                Content = content,
                Runs = runs,
                Formatting = formatting,
                Level = level,
                BulletChar = bulletChar,
                AutoNumberType = autoNumberType,
                HasBullet = hasBullet,
                LineSpacing = paragraphLineSpacing,
                SpaceBefore = spaceBefore,
                SpaceAfter = spaceAfter,
                MarginLeft = marginLeft,
                Indent = indent
            });
        }

        return new TypstTextElement
        {
            Paragraphs = paragraphs,
            AutoFit = autoFit,
            VerticalAlign = vertAlign,
            PaddingLeft = padLeft,
            PaddingTop = padTop,
            PaddingRight = padRight,
            PaddingBottom = padBottom,
            LineSpacing = paragraphs.FirstOrDefault()?.LineSpacing,
            ParagraphCount = Math.Max(1, paragraphCount),
            HasExplicitLineBreaks = hasExplicitLineBreaks
        };
    }

    private static PlaceholderValues? GetPlaceholderType(P.Shape shape)
    {
        var nvSpPr = shape.NonVisualShapeProperties;
        if (nvSpPr == null) return null;

        // The placeholder shape can be in different locations depending on the shape type
        PlaceholderShape? ph = null;

        // Try NonVisualShapeProperties first (for shapes)
        ph = nvSpPr.Elements<PlaceholderShape>().FirstOrDefault();

        // For some shapes, the placeholder might be in the ApplicationNonVisualDrawingProperties
        if (ph == null)
        {
            var appProps = nvSpPr.ApplicationNonVisualDrawingProperties;
            ph = appProps?.Elements<PlaceholderShape>().FirstOrDefault();
        }

        if (ph == null) return null;

        // The OpenXML SDK often fails to parse enum values from attributes
        // Use regex parsing from OuterXml as the primary method
        var outerXml = ph.OuterXml;
        if (!string.IsNullOrEmpty(outerXml))
        {
            var match = System.Text.RegularExpressions.Regex.Match(outerXml, @"type\s*=\s*""([^""]*)""");
            if (match.Success)
            {
                var typeStr = match.Groups[1].Value;
                return typeStr switch
                {
                    "title" => PlaceholderValues.Title,
                    "ctrTitle" => PlaceholderValues.CenteredTitle,
                    "subTitle" => PlaceholderValues.SubTitle,
                    "body" => PlaceholderValues.Body,
                    "pic" => PlaceholderValues.Picture,
                    "chart" => PlaceholderValues.Chart,
                    "tbl" => PlaceholderValues.Table,
                    "sldNum" => PlaceholderValues.SlideNumber,
                    "ftr" => PlaceholderValues.Footer,
                    "hdr" => PlaceholderValues.Header,
                    "obj" => PlaceholderValues.Object,
                    "dt" => PlaceholderValues.DateAndTime,
                    _ => (PlaceholderValues?)null
                };
            }
        }

        return null;
    }

    private static PlaceholderValues? GetPlaceholderTypeFromLayout(SlidePart slidePart, int idx)
    {
        var layoutPart = slidePart.SlideLayoutPart;
        if (layoutPart?.SlideLayout?.CommonSlideData?.ShapeTree == null)
            return null;

        // Find matching placeholder in layout by idx
        foreach (var layoutShape in layoutPart.SlideLayout.CommonSlideData.ShapeTree.ChildElements.OfType<P.Shape>())
        {
            var layoutPh = layoutShape.NonVisualShapeProperties?.Elements<PlaceholderShape>().FirstOrDefault();
            if (layoutPh == null)
            {
                var appProps = layoutShape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties;
                layoutPh = appProps?.Elements<PlaceholderShape>().FirstOrDefault();
            }

            if (layoutPh != null)
            {
                var outerXml = layoutPh.OuterXml;
                if (!string.IsNullOrEmpty(outerXml))
                {
                    var idxMatch = System.Text.RegularExpressions.Regex.Match(outerXml, @"idx\s*=\s*""([^""]*)""");
                    if (idxMatch.Success && int.TryParse(idxMatch.Groups[1].Value, out var layoutIdx) && layoutIdx == idx)
                    {
                        // Try explicit type first
                        var typeMatch = System.Text.RegularExpressions.Regex.Match(outerXml, @"type\s*=\s*""([^""]*)""");
                        if (typeMatch.Success && !string.IsNullOrEmpty(typeMatch.Groups[1].Value))
                        {
                            var typeStr = typeMatch.Groups[1].Value;
                            return typeStr switch
                            {
                                "title" => PlaceholderValues.Title,
                                "ctrTitle" => PlaceholderValues.CenteredTitle,
                                "subTitle" => PlaceholderValues.SubTitle,
                                "body" => PlaceholderValues.Body,
                                "pic" => PlaceholderValues.Picture,
                                "chart" => PlaceholderValues.Chart,
                                "tbl" => PlaceholderValues.Table,
                                "sldNum" => PlaceholderValues.SlideNumber,
                                "ftr" => PlaceholderValues.Footer,
                                "hdr" => PlaceholderValues.Header,
                                "obj" => PlaceholderValues.Object,
                                _ => (PlaceholderValues?)null
                            };
                        }

                        // If no explicit type, infer from list style characteristics
                        var txBody = layoutShape.TextBody;
                        if (txBody != null)
                        {
                            var lstStyle = txBody.ChildElements.FirstOrDefault(e => e.LocalName == "lstStyle");
                            if (lstStyle != null)
                            {
                                var lstStyleXml = lstStyle.OuterXml;
                                if (lstStyleXml.Contains("buAutoNum") || lstStyleXml.Contains("buChar") || lstStyleXml.Contains("buNone"))
                                {
                                    // Placeholder with list styles is typically a body/content placeholder
                                    return PlaceholderValues.Body;
                                }
                            }
                        }
                    }
                }
            }
        }

        return null;
    }

    private double? ExtractParagraphLineSpacing(Drawing.Paragraph paragraph)
    {
        var pPr = paragraph.Elements<Drawing.ParagraphProperties>().FirstOrDefault();
        return ExtractLineSpacingFromElement(pPr);
    }

    private static double? ExtractLineSpacingFromElement(OpenXmlElement? element)
    {
        if (element == null) return null;

        var lnSpc = element.ChildElements.FirstOrDefault(e => e.LocalName == "lnSpc");
        if (lnSpc == null) return null;

        // Try spcPts first (absolute points)
        var spcPts = lnSpc.ChildElements.FirstOrDefault(e => e.LocalName == "spcPts");
        if (spcPts != null)
        {
            var valAttr = spcPts.GetAttribute("val", "");
            if (int.TryParse(valAttr.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value))
                return value / 100.0;
        }

        // Try spcPct (percentage of line height)
        var spcPct = lnSpc.ChildElements.FirstOrDefault(e => e.LocalName == "spcPct");
        if (spcPct != null)
        {
            var valAttr = spcPct.GetAttribute("val", "");
            if (int.TryParse(valAttr.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value))
                return value / 100000.0; // spcPct is in 1/1000ths of a percent (100000 = 100% = 1.0)
        }

        return null;
    }

    private static double? ExtractLineSpacingFromLstStyle(OpenXmlElement? lstStyle, int level)
    {
        if (lstStyle == null) return null;

        var levelName = $"lvl{level + 1}pPr";
        var lvlPpr = lstStyle.ChildElements.FirstOrDefault(e => e.LocalName == levelName);
        if (lvlPpr == null) return null;

        return ExtractLineSpacingFromElement(lvlPpr);
    }

    private string? ExtractParagraphAlignment(Drawing.Paragraph paragraph)
    {
        // ParagraphProperties is an element, not a property, in Drawing namespace
        var pPr = paragraph.Elements<Drawing.ParagraphProperties>().FirstOrDefault();
        if (pPr == null) return null;

        // The OpenXML SDK doesn't always parse the algn attribute into Alignment property
        // Read the raw XML attribute directly
        try
        {
            var algnAttr = pPr.GetAttribute("algn", "");
            if (!string.IsNullOrEmpty(algnAttr.Value))
            {
                return algnAttr.Value switch
                {
                    "ctr" => "center",
                    "r" => "right",
                    "just" => "left",  // Typst doesn't have 'justify', use left as fallback
                    _ => "left"
                };
            }
        }
        catch
        {
            // Attribute doesn't exist in schema
        }

        return null;
    }

    private TypstTextFormatting ExtractTextFormatting(Drawing.Run run)
    {
        var runProps = run.RunProperties;
        if (runProps == null) return new TypstTextFormatting();

        var fmt = new TypstTextFormatting();

        // Font size (in hundredths of a point)
        if (runProps.FontSize?.Value != null)
        {
            fmt = fmt with { FontSize = runProps.FontSize.Value / 100.0 };
        }

        // Bold
        if (runProps.Bold?.Value != null)
        {
            fmt = fmt with { Bold = runProps.Bold.Value };
        }

        // Italic
        if (runProps.Italic?.Value != null)
        {
            fmt = fmt with { Italic = runProps.Italic.Value };
        }

        // Underline
        if (runProps.Underline?.Value != null && runProps.Underline.Value != Drawing.TextUnderlineValues.None)
        {
            fmt = fmt with { Underline = true };
        }

        // Color
        var color = ExtractRunColor(runProps);
        if (!string.IsNullOrEmpty(color))
        {
            fmt = fmt with { Color = color };
        }

        // Font family - map common "Bold" suffix fonts to base family
        var latinFont = runProps.Elements<Drawing.LatinFont>().FirstOrDefault();
        if (latinFont?.Typeface != null)
        {
            var fontName = latinFont.Typeface.Value ?? "";
            // If font name ends with " Bold" and bold isn't set, set it
            if (fontName.EndsWith(" Bold", StringComparison.OrdinalIgnoreCase))
            {
                fontName = fontName.Substring(0, fontName.Length - 5);
                fmt = fmt with { Bold = true };
            }
            fmt = fmt with { FontFamily = fontName };
        }

        return fmt;
    }

    private TypstTextFormatting ResolveParagraphFormatting(
        Drawing.Run? firstRun,
        Drawing.ParagraphProperties? pPr,
        OpenXmlElement? bodyLstStyle,
        int level,
        StyleResolver? styleResolver,
        int? placeholderIdx,
        PlaceholderValues? placeholderType)
    {
        var fmt = new TypstTextFormatting();

        // 1. Run level - from first run's rPr
        if (firstRun != null)
        {
            var runProps = firstRun.RunProperties;
            if (runProps != null)
            {
                if (runProps.FontSize?.Value != null)
                    fmt = fmt with { FontSize = runProps.FontSize.Value / 100.0 };
                if (runProps.Bold?.Value != null)
                    fmt = fmt with { Bold = runProps.Bold.Value };
                if (runProps.Italic?.Value != null)
                    fmt = fmt with { Italic = runProps.Italic.Value };

                var color = ExtractRunColor(runProps, styleResolver);
                if (!string.IsNullOrEmpty(color))
                    fmt = fmt with { Color = color };

                var latinFont = runProps.Elements<Drawing.LatinFont>().FirstOrDefault();
                if (latinFont?.Typeface != null)
                {
                    var fontName = latinFont.Typeface.Value ?? "";
                    if (fontName.EndsWith(" Bold", StringComparison.OrdinalIgnoreCase))
                    {
                        fontName = fontName.Substring(0, fontName.Length - 5);
                        fmt = fmt with { Bold = true };
                    }
                    fmt = fmt with { FontFamily = fontName };
                }
            }
        }

        // 2. Paragraph default: pPr/defRPr
        var defRPr = pPr?.Elements<Drawing.DefaultRunProperties>().FirstOrDefault();
        if (defRPr != null)
        {
            if (fmt.FontSize == 18.0 && defRPr.FontSize?.Value != null)
                fmt = fmt with { FontSize = defRPr.FontSize.Value / 100.0 };
            if (!fmt.Bold && defRPr.Bold?.Value != null)
                fmt = fmt with { Bold = defRPr.Bold.Value };
            if (!fmt.Italic && defRPr.Italic?.Value != null)
                fmt = fmt with { Italic = defRPr.Italic.Value };

            var defColor = ExtractDefRPrColor(defRPr);
            if (fmt.Color == "#000000" && !string.IsNullOrEmpty(defColor))
                fmt = fmt with { Color = defColor };

            var defLatinFont = defRPr.Elements<Drawing.LatinFont>().FirstOrDefault();
            if (fmt.FontFamily == "Arial" && defLatinFont?.Typeface != null)
            {
                var fontName = defLatinFont.Typeface.Value ?? "";
                if (fontName.EndsWith(" Bold", StringComparison.OrdinalIgnoreCase))
                {
                    fontName = fontName.Substring(0, fontName.Length - 5);
                    fmt = fmt with { Bold = true };
                }
                fmt = fmt with { FontFamily = fontName };
            }
        }

        // 3. Text body list style
        if (bodyLstStyle != null)
        {
            var bodyStyle = ExtractLstStyleDefRPr(bodyLstStyle, level);
            fmt = MergeDefaultStyle(fmt, bodyStyle);
        }

        // 4. Layout placeholder list style
        if (styleResolver != null)
        {
            var layoutStyle = styleResolver.GetLayoutPlaceholderLstStyle(placeholderIdx, placeholderType, level);
            fmt = MergeDefaultStyle(fmt, layoutStyle);

            // 5. Master placeholder list style
            var masterStyle = styleResolver.GetMasterPlaceholderLstStyle(placeholderIdx, placeholderType, level);
            fmt = MergeDefaultStyle(fmt, masterStyle);
        }

        return fmt;
    }

    private static TypstTextFormatting MergeDefaultStyle(TypstTextFormatting fmt, StyleResolver.DefaultTextStyle style)
    {
        if (fmt.FontSize == 18.0 && style.FontSize.HasValue)
            fmt = fmt with { FontSize = style.FontSize.Value };
        if (!fmt.Bold && style.Bold.HasValue)
            fmt = fmt with { Bold = style.Bold.Value };
        if (!fmt.Italic && style.Italic.HasValue)
            fmt = fmt with { Italic = style.Italic.Value };
        if (fmt.Color == "#000000" && !string.IsNullOrEmpty(style.Color))
            fmt = fmt with { Color = style.Color };
        if (fmt.FontFamily == "Arial" && !string.IsNullOrEmpty(style.FontFamily))
        {
            var fontName = style.FontFamily;
            if (fontName.EndsWith(" Bold", StringComparison.OrdinalIgnoreCase))
            {
                fontName = fontName.Substring(0, fontName.Length - 5);
                fmt = fmt with { Bold = true };
            }
            fmt = fmt with { FontFamily = fontName };
        }
        return fmt;
    }

    private static (string? BulletChar, string? AutoNumberType, bool HasBullet) ResolveBulletProperties(
        Drawing.ParagraphProperties? pPr,
        OpenXmlElement? bodyLstStyle,
        int level,
        StyleResolver? styleResolver,
        int? placeholderIdx,
        PlaceholderValues? placeholderType)
    {
        // 1. Paragraph level
        var info = ExtractBulletInfo(pPr);
        if (info.HasBullet || info.HasBulletNone)
            return (info.BulletChar, info.AutoNumberType, info.HasBullet);

        // 2. Text body list style
        info = ExtractBulletInfoFromLstStyle(bodyLstStyle, level);
        if (info.HasBullet || info.HasBulletNone)
            return (info.BulletChar, info.AutoNumberType, info.HasBullet);

        if (styleResolver != null)
        {
            // 3. Layout placeholder list style
            info = styleResolver.GetLayoutPlaceholderBulletInfo(placeholderIdx, placeholderType, level);
            if (info.HasBullet || info.HasBulletNone)
                return (info.BulletChar, info.AutoNumberType, info.HasBullet);

            // 4. Master placeholder list style
            info = styleResolver.GetMasterPlaceholderBulletInfo(placeholderIdx, placeholderType, level);
            if (info.HasBullet || info.HasBulletNone)
                return (info.BulletChar, info.AutoNumberType, info.HasBullet);

            // 5. Master txStyles
            info = styleResolver.GetMasterTxStyleBulletInfo(placeholderType, level);
            if (info.HasBullet || info.HasBulletNone)
                return (info.BulletChar, info.AutoNumberType, info.HasBullet);
        }

        return (null, null, false);
    }

    private static double? ResolveLineSpacing(
        Drawing.ParagraphProperties? pPr,
        OpenXmlElement? bodyLstStyle,
        int level,
        StyleResolver? styleResolver,
        int? placeholderIdx,
        PlaceholderValues? placeholderType)
    {
        // 1. Paragraph level
        var spacing = ExtractLineSpacingFromElement(pPr);
        if (spacing.HasValue)
            return spacing.Value;

        // 2. Text body list style
        spacing = ExtractLineSpacingFromLstStyle(bodyLstStyle, level);
        if (spacing.HasValue)
            return spacing.Value;

        if (styleResolver != null)
        {
            // 3. Layout placeholder list style
            spacing = styleResolver.GetLayoutPlaceholderLineSpacing(placeholderIdx, placeholderType, level);
            if (spacing.HasValue)
                return spacing.Value;

            // 4. Master placeholder list style
            spacing = styleResolver.GetMasterPlaceholderLineSpacing(placeholderIdx, placeholderType, level);
            if (spacing.HasValue)
                return spacing.Value;

            // 5. Master txStyles
            spacing = styleResolver.GetMasterTxStyleLineSpacing(placeholderType, level);
            if (spacing.HasValue)
                return spacing.Value;
        }

        return null;
    }

    private static (double? SpaceBefore, double? SpaceAfter) ResolveParagraphSpacing(
        Drawing.ParagraphProperties? pPr,
        OpenXmlElement? bodyLstStyle,
        int level,
        StyleResolver? styleResolver,
        int? placeholderIdx,
        PlaceholderValues? placeholderType)
    {
        // 1. Paragraph level
        var spacing = ExtractParagraphSpacing(pPr);
        if (spacing.SpaceBefore.HasValue || spacing.SpaceAfter.HasValue)
            return spacing;

        // 2. Text body list style
        spacing = ExtractParagraphSpacingFromLstStyle(bodyLstStyle, level);
        if (spacing.SpaceBefore.HasValue || spacing.SpaceAfter.HasValue)
            return spacing;

        if (styleResolver != null)
        {
            // 3. Layout placeholder list style
            spacing = styleResolver.GetLayoutPlaceholderSpacing(placeholderIdx, placeholderType, level);
            if (spacing.SpaceBefore.HasValue || spacing.SpaceAfter.HasValue)
                return spacing;

            // 4. Master placeholder list style
            spacing = styleResolver.GetMasterPlaceholderSpacing(placeholderIdx, placeholderType, level);
            if (spacing.SpaceBefore.HasValue || spacing.SpaceAfter.HasValue)
                return spacing;

            // 5. Master txStyles
            spacing = styleResolver.GetMasterTxStyleSpacing(placeholderType, level);
            if (spacing.SpaceBefore.HasValue || spacing.SpaceAfter.HasValue)
                return spacing;
        }

        return (null, null);
    }

    private static (double? SpaceBefore, double? SpaceAfter) ExtractParagraphSpacing(OpenXmlElement? pPr)
    {
        if (pPr == null) return (null, null);

        double? spcBef = null;
        double? spcAft = null;

        var spcBefEl = pPr.ChildElements.FirstOrDefault(e => e.LocalName == "spcBef");
        if (spcBefEl != null)
        {
            var spcPts = spcBefEl.ChildElements.FirstOrDefault(e => e.LocalName == "spcPts");
            if (spcPts != null)
            {
                var valAttr = spcPts.GetAttribute("val", "");
                if (int.TryParse(valAttr.Value, out var ptsHundredths))
                    spcBef = ptsHundredths / 100.0;
            }
            var spcPct = spcBefEl.ChildElements.FirstOrDefault(e => e.LocalName == "spcPct");
            if (spcPct != null)
            {
                var valAttr = spcPct.GetAttribute("val", "");
                if (int.TryParse(valAttr.Value, out var pct))
                    spcBef = pct / 100000.0;
            }
        }

        var spcAftEl = pPr.ChildElements.FirstOrDefault(e => e.LocalName == "spcAft");
        if (spcAftEl != null)
        {
            var spcPts = spcAftEl.ChildElements.FirstOrDefault(e => e.LocalName == "spcPts");
            if (spcPts != null)
            {
                var valAttr = spcPts.GetAttribute("val", "");
                if (int.TryParse(valAttr.Value, out var ptsHundredths))
                    spcAft = ptsHundredths / 100.0;
            }
            var spcPct = spcAftEl.ChildElements.FirstOrDefault(e => e.LocalName == "spcPct");
            if (spcPct != null)
            {
                var valAttr = spcPct.GetAttribute("val", "");
                if (int.TryParse(valAttr.Value, out var pct))
                    spcAft = pct / 100000.0;
            }
        }

        return (spcBef, spcAft);
    }

    private static (double? SpaceBefore, double? SpaceAfter) ExtractParagraphSpacingFromLstStyle(OpenXmlElement? lstStyle, int level)
    {
        if (lstStyle == null) return (null, null);

        var levelName = $"lvl{level + 1}pPr";
        var lvlPpr = lstStyle.ChildElements.FirstOrDefault(e => e.LocalName == levelName);
        if (lvlPpr == null) return (null, null);

        return ExtractParagraphSpacing(lvlPpr);
    }

    private static (string? BulletChar, string? AutoNumberType, bool HasBullet, bool HasBulletNone) ExtractBulletInfo(OpenXmlElement? element)
    {
        if (element == null) return (null, null, false, false);

        var buChar = element.ChildElements.FirstOrDefault(e => e.LocalName == "buChar");
        if (buChar != null)
        {
            var charAttr = buChar.GetAttribute("char", "");
            if (!string.IsNullOrEmpty(charAttr.Value))
                return (charAttr.Value, null, true, false);
        }

        var buAutoNum = element.ChildElements.FirstOrDefault(e => e.LocalName == "buAutoNum");
        if (buAutoNum != null)
        {
            var typeAttr = buAutoNum.GetAttribute("type", "");
            if (!string.IsNullOrEmpty(typeAttr.Value))
                return (null, typeAttr.Value, true, false);
        }

        var buNone = element.ChildElements.FirstOrDefault(e => e.LocalName == "buNone");
        if (buNone != null)
            return (null, null, false, true);

        return (null, null, false, false);
    }

    private static (string? BulletChar, string? AutoNumberType, bool HasBullet, bool HasBulletNone) ExtractBulletInfoFromLstStyle(OpenXmlElement? lstStyle, int level)
    {
        if (lstStyle == null) return (null, null, false, false);

        var levelName = $"lvl{level + 1}pPr";
        var lvlPpr = lstStyle.ChildElements.FirstOrDefault(e => e.LocalName == levelName);
        if (lvlPpr == null) return (null, null, false, false);

        return ExtractBulletInfo(lvlPpr);
    }

    private string? ExtractDefRPrColor(Drawing.DefaultRunProperties defRPr)
    {
        var solidFill = defRPr.Elements<Drawing.SolidFill>().FirstOrDefault();
        if (solidFill != null)
        {
            return ExtractColor(solidFill);
        }
        return null;
    }

    private static StyleResolver.DefaultTextStyle ExtractLstStyleDefRPr(OpenXmlElement? lstStyle, int level)
    {
        if (lstStyle == null) return new StyleResolver.DefaultTextStyle();

        var levelName = $"lvl{level + 1}pPr";
        var lvlPpr = lstStyle.ChildElements.FirstOrDefault(e => e.LocalName == levelName);
        if (lvlPpr == null) return new StyleResolver.DefaultTextStyle();

        var defRPr = lvlPpr.Elements<Drawing.DefaultRunProperties>().FirstOrDefault();
        if (defRPr == null) return new StyleResolver.DefaultTextStyle();

        var style = new StyleResolver.DefaultTextStyle
        {
            FontSize = defRPr.FontSize?.Value != null ? defRPr.FontSize.Value / 100.0 : null,
            Bold = defRPr.Bold?.Value,
            Italic = defRPr.Italic?.Value,
            Underline = defRPr.Underline?.Value != null && defRPr.Underline.Value != Drawing.TextUnderlineValues.None,
            Color = ExtractDefRPrColorStatic(defRPr)
        };

        var latinFont = defRPr.Elements<Drawing.LatinFont>().FirstOrDefault();
        if (latinFont?.Typeface != null)
            style.FontFamily = latinFont.Typeface.Value;

        return style;
    }

    private static string? ExtractDefRPrColorStatic(Drawing.DefaultRunProperties defRPr)
    {
        var solidFill = defRPr.Elements<Drawing.SolidFill>().FirstOrDefault();
        return solidFill != null ? ExtractSolidFillColorStatic(solidFill, null) : null;
    }

    private string? ExtractRunColor(Drawing.RunProperties runProps, StyleResolver? styleResolver = null)
    {
        var solidFill = runProps.Elements<Drawing.SolidFill>().FirstOrDefault();
        if (solidFill != null)
        {
            return ExtractColor(solidFill, styleResolver);
        }

        return null;
    }

    private string? ExtractColor(Drawing.SolidFill solidFill, StyleResolver? styleResolver = null)
        => ExtractSolidFillColorStatic(solidFill, styleResolver);

    /// <summary>
    /// Reads the native pixel dimensions from a PNG or JPEG byte array header.
    /// Returns (width, height) or null for unsupported/unparseable formats.
    /// </summary>
    private static (int Width, int Height)? GetNativeImageDimensions(byte[] imageData)
    {
        if (imageData.Length < 24)
            return null;

        // PNG: signature \x89PNG\r\n\x1a\n followed by IHDR chunk at offset 8
        // IHDR length (4 bytes) + "IHDR" (4 bytes) + width (4 bytes BE) + height (4 bytes BE)
        if (imageData[0] == 0x89 && imageData[1] == 0x50 && imageData[2] == 0x4E && imageData[3] == 0x47
            && imageData[4] == 0x0D && imageData[5] == 0x0A && imageData[6] == 0x1A && imageData[7] == 0x0A)
        {
            if (imageData.Length < 24)
                return null;
            int width = (imageData[16] << 24) | (imageData[17] << 16) | (imageData[18] << 8) | imageData[19];
            int height = (imageData[20] << 24) | (imageData[21] << 16) | (imageData[22] << 8) | imageData[23];
            return (width, height);
        }

        // JPEG: SOI marker 0xFF 0xD8
        if (imageData[0] == 0xFF && imageData[1] == 0xD8)
        {
            int offset = 2;
            while (offset + 4 < imageData.Length)
            {
                // JPEG markers start with 0xFF, followed by marker type (not 0xFF, not 0x00)
                // Marker length (2 bytes BE) includes length bytes but not the marker bytes.
                int marker = imageData[offset];
                if (marker != 0xFF)
                    break;
                offset++;
                int markerType = imageData[offset++];
                if (markerType == 0x00 || markerType == 0xFF)
                    continue;
                if (markerType == 0xD9) // EOI
                    break;
                if (offset + 2 > imageData.Length)
                    break;
                int segmentLength = (imageData[offset] << 8) | imageData[offset + 1];
                if (segmentLength < 2 || offset + segmentLength > imageData.Length)
                    break;
                // SOF markers: 0xC0-0xC3, 0xC5-0xC7, 0xC9-0xCB, 0xCD-0xCF
                if ((markerType >= 0xC0 && markerType <= 0xC3)
                    || (markerType >= 0xC5 && markerType <= 0xC7)
                    || (markerType >= 0xC9 && markerType <= 0xCB)
                    || (markerType >= 0xCD && markerType <= 0xCF))
                {
                    if (offset + 7 > imageData.Length)
                        break;
                    // After length (2 bytes), precision (1 byte), height (2 bytes BE), width (2 bytes BE)
                    int jpegHeight = (imageData[offset + 3] << 8) | imageData[offset + 4];
                    int jpegWidth = (imageData[offset + 5] << 8) | imageData[offset + 6];
                    return (jpegWidth, jpegHeight);
                }
                offset += segmentLength;
            }
        }

        // Unsupported format
        return null;
    }

    private TypstImageElement? ExtractImage(SlidePart slidePart, P.Picture picture)
    {
        var blipFill = picture.BlipFill;
        if (blipFill == null) return null;

        var blip = blipFill.Blip;
        if (blip == null) return null;

        var embed = blip.Embed?.Value;
        if (string.IsNullOrEmpty(embed)) return null;

        var imagePart = slidePart.GetPartById(embed) as ImagePart;
        if (imagePart == null) return null;

        // Determine file extension
        var extension = imagePart.ContentType switch
        {
            "image/png" => "png",
            "image/jpeg" => "jpg",
            "image/gif" => "gif",
            "image/bmp" => "bmp",
            "image/tiff" => "tiff",
            "image/x-icon" => "ico",
            "image/svg+xml" => "svg",
            _ => "bin"
        };

        _imageCounter++;
        var fileName = $"image_{_imageCounter}.{extension}";
        var fullPath = Path.Combine(_assetsDirectory, fileName);

        // Read image data into memory to detect native dimensions before writing
        byte[] imageData;
        using (var stream = imagePart.GetStream())
        {
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            imageData = ms.ToArray();
        }
        File.WriteAllBytes(fullPath, imageData);

        var dimensions = GetNativeImageDimensions(imageData);

        return new TypstImageElement
        {
            FileName = fileName,
            FullPath = fullPath,
            Width = 0, // Will be set from shape position
            Height = 0,
            PixelWidth = dimensions?.Width,
            PixelHeight = dimensions?.Height
        };
    }

    private TypstImageElement? ExtractImageFromBlipFill(SlidePart slidePart, Drawing.BlipFill blipFill)
    {
        var blip = blipFill.Blip;
        if (blip == null) return null;

        var embed = blip.Embed?.Value;
        if (string.IsNullOrEmpty(embed)) return null;

        var imagePart = slidePart.GetPartById(embed) as ImagePart;
        if (imagePart == null) return null;

        // Determine file extension
        var extension = imagePart.ContentType switch
        {
            "image/png" => "png",
            "image/jpeg" => "jpg",
            "image/gif" => "gif",
            "image/bmp" => "bmp",
            "image/tiff" => "tiff",
            "image/x-icon" => "ico",
            "image/svg+xml" => "svg",
            _ => "bin"
        };

        _imageCounter++;
        var fileName = $"image_{_imageCounter}.{extension}";
        var fullPath = Path.Combine(_assetsDirectory, fileName);

        // Read image data into memory to detect native dimensions before writing
        byte[] imageData;
        using (var stream = imagePart.GetStream())
        {
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            imageData = ms.ToArray();
        }
        File.WriteAllBytes(fullPath, imageData);

        var dimensions = GetNativeImageDimensions(imageData);

        return new TypstImageElement
        {
            FileName = fileName,
            FullPath = fullPath,
            Width = 0,
            Height = 0,
            PixelWidth = dimensions?.Width,
            PixelHeight = dimensions?.Height
        };
    }

    private TypstTableElement ExtractTable(Drawing.Table table, StyleResolver styleResolver)
    {
        var rows = new List<List<TypstTableCell>>();
        var columnWidths = new List<double>();
        var rowHeights = new List<double>();

        // Extract column widths from table grid
        var tableGrid = table.TableGrid;
        if (tableGrid != null)
        {
            foreach (var gridCol in tableGrid.Elements<Drawing.GridColumn>())
            {
                var width = gridCol.Width?.Value ?? 0;
                columnWidths.Add(EmuToPt((long)width));
            }
        }

        var tableRows = table.Elements<Drawing.TableRow>().ToList();
        int totalRows = tableRows.Count;
        int totalCols = columnWidths.Count;

        // Read table style and flags
        var tableProps = table.TableProperties;
        string? styleId = null;
        bool firstRowFlag = false, bandRowFlag = false, firstColFlag = false, lastColFlag = false, lastRowFlag = false;

        if (tableProps != null)
        {
            var styleIdElem = tableProps.ChildElements.FirstOrDefault(e => e.LocalName == "tableStyleId");
            if (styleIdElem != null)
                styleId = styleIdElem.InnerText;

            firstRowFlag = tableProps.FirstRow?.Value == true;
            bandRowFlag = tableProps.BandRow?.Value == true;
            firstColFlag = tableProps.FirstColumn?.Value == true;
            lastColFlag = tableProps.LastColumn?.Value == true;
            lastRowFlag = tableProps.LastRow?.Value == true;
        }

        _tableStyles.TryGetValue(styleId ?? "", out var tableStyle);

        // Derive border color and width from wholeTbl border; fallback to black/1.0
        var borderColor = "#000000";
        var borderWidth = 1.0;
        TableStylePart? wholeTblPart = null;
        if (tableStyle?.Parts.TryGetValue("wholeTbl", out wholeTblPart) == true)
        {
            borderColor = wholeTblPart.BorderTopColor ?? borderColor;
            borderWidth = wholeTblPart.BorderTopWidth ?? borderWidth;
        }

        for (int rowIndex = 0; rowIndex < totalRows; rowIndex++)
        {
            var row = tableRows[rowIndex];
            var rowHeight = row.Height?.Value ?? 0;
            rowHeights.Add(EmuToPt((long)rowHeight));

            var rowCells = new List<TypstTableCell>();
            var cells = row.Elements<Drawing.TableCell>().ToList();
            for (int colIndex = 0; colIndex < cells.Count; colIndex++)
            {
                var cell = cells[colIndex];
                var stylePart = ResolveCellStylePart(rowIndex, colIndex, totalRows, totalCols, tableStyle, firstRowFlag, bandRowFlag, firstColFlag, lastColFlag, lastRowFlag);

                // Merge explicit cell borders with style part
                stylePart = MergeExplicitCellBorders(cell, stylePart, styleResolver);

                stylePart = ApplyTableGridBorders(stylePart, rowIndex, colIndex, totalRows, totalCols);

                var cellFormatting = ExtractCellFormatting(cell, stylePart, styleResolver);
                var paragraphs = ExtractCellParagraphs(cell, cellFormatting, styleResolver);
                var cellText = paragraphs.Count > 0
                    ? string.Join("\n\n", paragraphs.Select(p => p.Content)).Trim()
                    : ExtractCellText(cell);
                var explicitBg = ExtractCellBackground(cell);
                var bgColor = explicitBg.Specified ? explicitBg.Color : (stylePart?.BackgroundCleared == true ? null : stylePart?.BackgroundColor);

                rowCells.Add(new TypstTableCell
                {
                    Content = cellText,
                    Paragraphs = paragraphs,
                    Formatting = cellFormatting,
                    BackgroundColor = bgColor,
                    Insets = ExtractCellInsets(cell),
                    VerticalAlign = ExtractCellVerticalAlign(cell),
                    StylePart = stylePart
                });
            }
            rows.Add(rowCells);
        }

        return new TypstTableElement
        {
            Rows = rows,
            ColumnWidths = columnWidths,
            RowHeights = rowHeights,
            BorderColor = borderColor,
            BorderWidth = borderWidth
        };
    }

    private string ExtractCellText(Drawing.TableCell cell)
    {
        var sb = new StringBuilder();
        var textBody = cell.TextBody;
        if (textBody == null) return "";

        foreach (var paragraph in textBody.Elements<Drawing.Paragraph>())
        {
            if (sb.Length > 0)
                sb.Append("\n\n");

            foreach (var child in paragraph.ChildElements)
            {
                if (child is Drawing.Run run && run.Text?.Text != null)
                {
                    sb.Append(run.Text.Text);
                }
                else if (child.LocalName == "br")
                {
                    sb.Append('\n');
                }
            }
        }

        return sb.ToString().Trim();
    }

    private List<TypstParagraph> ExtractCellParagraphs(Drawing.TableCell cell, TypstTextFormatting cellFormatting, StyleResolver? styleResolver)
    {
        var paragraphs = new List<TypstParagraph>();
        var textBody = cell.TextBody;
        if (textBody == null) return paragraphs;

        foreach (var paragraph in textBody.Elements<Drawing.Paragraph>())
        {
            var align = ExtractTableParagraphAlignment(paragraph) ?? "left";
            var paragraphFormatting = cellFormatting with { Align = align };
            var runs = new List<TypstTextRun>();
            var content = new StringBuilder();

            foreach (var child in paragraph.ChildElements)
            {
                if (child is Drawing.Run run)
                {
                    var runText = run.Text?.Text;
                    if (string.IsNullOrEmpty(runText))
                        continue;

                    var runFormatting = MergeRunWithParagraphDefaults(paragraphFormatting, run, styleResolver, requireExplicitOnOff: true);
                    content.Append(runText);
                    runs.Add(new TypstTextRun
                    {
                        Content = runText,
                        Formatting = runFormatting
                    });
                }
                else if (child.LocalName == "br")
                {
                    content.Append('\n');
                    runs.Add(new TypstTextRun
                    {
                        IsLineBreak = true,
                        Formatting = paragraphFormatting
                    });
                }
            }

            TrimTrailingTabs(content, runs);

            paragraphs.Add(new TypstParagraph
            {
                Content = content.ToString(),
                Runs = runs,
                Formatting = paragraphFormatting
            });
        }

        return paragraphs;
    }

    private static string? ExtractTableParagraphAlignment(Drawing.Paragraph paragraph)
    {
        var pPr = paragraph.Elements<Drawing.ParagraphProperties>().FirstOrDefault();
        if (pPr == null) return null;

        var match = Regex.Match(pPr.OuterXml, "\\balgn=\"(?<value>[^\"]+)\"");
        if (!match.Success) return null;

        return match.Groups["value"].Value switch
        {
            "ctr" => "center",
            "r" => "right",
            "just" => "left",
            _ => "left"
        };
    }

    private static string? BuildCellInset(TypstTableCellInsets? insets)
    {
        if (insets == null)
            return null;

        var parts = new List<string>();
        if (insets.Left.HasValue)
            parts.Add($"left: {FormatPt(insets.Left.Value)}");
        if (insets.Right.HasValue)
            parts.Add($"right: {FormatPt(insets.Right.Value)}");
        if (insets.Top.HasValue)
            parts.Add($"top: {FormatPt(insets.Top.Value)}");
        if (insets.Bottom.HasValue)
            parts.Add($"bottom: {FormatPt(insets.Bottom.Value)}");

        return parts.Count > 0 ? $"({string.Join(", ", parts)})" : null;
    }

    private static TypstTableCellInsets? ExtractCellInsets(Drawing.TableCell cell)
    {
        var tcPr = cell.TableCellProperties;
        if (tcPr == null)
            return null;

        var left = GetEmuAttributeAsPt(tcPr, "marL");
        var right = GetEmuAttributeAsPt(tcPr, "marR");
        var top = GetEmuAttributeAsPt(tcPr, "marT");
        var bottom = GetEmuAttributeAsPt(tcPr, "marB");

        if (!left.HasValue && !right.HasValue && !top.HasValue && !bottom.HasValue)
            return null;

        return new TypstTableCellInsets
        {
            Left = left,
            Right = right,
            Top = top,
            Bottom = bottom
        };
    }

    private static string? ExtractCellVerticalAlign(Drawing.TableCell cell)
    {
        var anchor = GetAttributeValue(cell.TableCellProperties, "anchor");
        return anchor == "ctr" ? "center" : null;
    }

    private static double? GetEmuAttributeAsPt(OpenXmlElement? element, string attributeName)
    {
        var value = GetAttributeValue(element, attributeName);
        return long.TryParse(value, CultureInfo.InvariantCulture, out var emu) ? EmuToPt(emu) : null;
    }

    private TypstTextFormatting ExtractCellFormatting(Drawing.TableCell cell, TableStylePart? stylePart, StyleResolver? styleResolver)
    {
        var formatting = new TypstTextFormatting();

        // Apply inherited table style first; direct run properties override below.
        if (stylePart != null)
        {
            if (stylePart.TextBold.HasValue)
                formatting = formatting with { Bold = stylePart.TextBold.Value };
            if (stylePart.TextItalic.HasValue)
                formatting = formatting with { Italic = stylePart.TextItalic.Value };
            if (!string.IsNullOrEmpty(stylePart.TextColor))
                formatting = formatting with { Color = stylePart.TextColor };
            if (stylePart.TextFontSize.HasValue)
                formatting = formatting with { FontSize = stylePart.TextFontSize.Value };
        }

        var firstRun = cell.TextBody?
            .Elements<Drawing.Paragraph>()
            .SelectMany(p => p.Elements<Drawing.Run>())
            .FirstOrDefault(r => r.RunProperties != null);
        if (firstRun?.RunProperties != null)
        {
            formatting = ApplyDirectRunFormatting(formatting, firstRun.RunProperties, styleResolver);
        }

        return formatting;
    }

    private TypstTextFormatting ApplyDirectRunFormatting(TypstTextFormatting formatting, Drawing.RunProperties runProps, StyleResolver? styleResolver)
    {
        if (runProps.FontSize?.Value != null)
            formatting = formatting with { FontSize = runProps.FontSize.Value / 100.0 };
        var bold = GetExplicitOnOffAttribute(runProps, "b");
        if (bold.HasValue)
            formatting = formatting with { Bold = bold.Value };
        var italic = GetExplicitOnOffAttribute(runProps, "i");
        if (italic.HasValue)
            formatting = formatting with { Italic = italic.Value };

        var color = ExtractRunColor(runProps, styleResolver);
        if (!string.IsNullOrEmpty(color))
            formatting = formatting with { Color = color };

        var latinFont = runProps.Elements<Drawing.LatinFont>().FirstOrDefault();
        if (latinFont?.Typeface != null)
            formatting = formatting with { FontFamily = latinFont.Typeface.Value ?? formatting.FontFamily };

        return formatting;
    }

    private (bool Specified, string? Color) ExtractCellBackground(Drawing.TableCell cell)
    {
        var cellProps = cell.TableCellProperties;
        if (cellProps == null) return (false, null);

        if (cellProps.Elements<Drawing.NoFill>().Any())
            return (true, null);

        var fill = cellProps.Elements<Drawing.SolidFill>().FirstOrDefault();
        if (fill != null)
        {
            return (true, ExtractColor(fill));
        }

        return (false, null);
    }

    private void LoadTableStyles(StyleResolver? styleResolver)
    {
        var part = _document.PresentationPart!.TableStylesPart;
        if (part?.TableStyleList == null) return;

        foreach (var style in part.TableStyleList.ChildElements.OfType<Drawing.TableStyleEntry>())
        {
            var styleId = style.StyleId?.Value;
            if (string.IsNullOrEmpty(styleId)) continue;

            var definition = new TableStyleDefinition { StyleId = styleId };
            ExtractTableStylePart(style.WholeTable, "wholeTbl", definition, styleResolver);
            ExtractTableStylePart(style.Band1Horizontal, "band1H", definition, styleResolver);
            ExtractTableStylePart(style.Band2Horizontal, "band2H", definition, styleResolver);
            ExtractTableStylePart(style.FirstRow, "firstRow", definition, styleResolver);
            ExtractTableStylePart(style.LastRow, "lastRow", definition, styleResolver);
            ExtractTableStylePart(style.FirstColumn, "firstCol", definition, styleResolver);
            ExtractTableStylePart(style.LastColumn, "lastCol", definition, styleResolver);

            _tableStyles[styleId] = definition;
        }
    }

    private void ExtractTableStylePart(Drawing.TablePartStyleType? part, string key, TableStyleDefinition definition, StyleResolver? styleResolver)
    {
        if (part == null) return;

        var tcStyle = part.TableCellStyle;
        if (tcStyle == null) return;

        string? bgColor = null;
        bool backgroundCleared = false;
        var fill = tcStyle.Elements<Drawing.FillProperties>().FirstOrDefault();
        if (fill != null)
        {
            backgroundCleared = fill.Elements<Drawing.NoFill>().Any();
            var solidFill = fill.Elements<Drawing.SolidFill>().FirstOrDefault();
            if (solidFill != null)
            {
                bgColor = ExtractSolidFillColor(solidFill, styleResolver);
            }
        }

        var borders = tcStyle.TableCellBorders;
        var (borderTop, borderTopWidth, borderTopNone) = ExtractBorderInfo(borders?.TopBorder, styleResolver);
        var (borderBottom, borderBottomWidth, borderBottomNone) = ExtractBorderInfo(borders?.BottomBorder, styleResolver);
        var (borderLeft, borderLeftWidth, borderLeftNone) = ExtractBorderInfo(borders?.LeftBorder, styleResolver);
        var (borderRight, borderRightWidth, borderRightNone) = ExtractBorderInfo(borders?.RightBorder, styleResolver);
        var (borderInsideH, borderInsideHWidth, borderInsideHNone) = ExtractBorderInfo(GetTableBorderByLocalName(borders, "insideH"), styleResolver);
        var (borderInsideV, borderInsideVWidth, borderInsideVNone) = ExtractBorderInfo(GetTableBorderByLocalName(borders, "insideV"), styleResolver);

        // Parse text style
        var tcTxStyle = part.TableCellTextStyle;
        bool? textBold = null;
        bool? textItalic = null;
        string? textColor = null;
        double? textFontSize = null;

        if (tcTxStyle != null)
        {
            var defRPr = tcTxStyle.Elements<Drawing.DefaultRunProperties>().FirstOrDefault();
            if (defRPr != null)
            {
                textBold = defRPr.Bold?.Value;
                textItalic = defRPr.Italic?.Value;
                textFontSize = defRPr.FontSize?.Value != null ? defRPr.FontSize.Value / 100.0 : (double?)null;

                var solidFill = defRPr.Elements<Drawing.SolidFill>().FirstOrDefault();
                if (solidFill != null)
                    textColor = ExtractSolidFillColor(solidFill, styleResolver);
            }
            ApplyDirectTableTextStyle(tcTxStyle, styleResolver, ref textBold, ref textItalic, ref textColor, ref textFontSize);
        }

        definition.Parts[key] = new TableStylePart
        {
            BackgroundColor = bgColor,
            BackgroundCleared = backgroundCleared,
            BorderTopColor = borderTop,
            BorderBottomColor = borderBottom,
            BorderLeftColor = borderLeft,
            BorderRightColor = borderRight,
            BorderInsideHColor = borderInsideH,
            BorderInsideVColor = borderInsideV,
            BorderTopWidth = borderTopWidth,
            BorderBottomWidth = borderBottomWidth,
            BorderLeftWidth = borderLeftWidth,
            BorderRightWidth = borderRightWidth,
            BorderInsideHWidth = borderInsideHWidth,
            BorderInsideVWidth = borderInsideVWidth,
            BorderTopNone = borderTopNone,
            BorderBottomNone = borderBottomNone,
            BorderLeftNone = borderLeftNone,
            BorderRightNone = borderRightNone,
            BorderInsideHNone = borderInsideHNone,
            BorderInsideVNone = borderInsideVNone,
            BorderTopState = ToBorderState(borderTop, borderTopWidth, borderTopNone),
            BorderBottomState = ToBorderState(borderBottom, borderBottomWidth, borderBottomNone),
            BorderLeftState = ToBorderState(borderLeft, borderLeftWidth, borderLeftNone),
            BorderRightState = ToBorderState(borderRight, borderRightWidth, borderRightNone),
            BorderInsideHState = ToBorderState(borderInsideH, borderInsideHWidth, borderInsideHNone),
            BorderInsideVState = ToBorderState(borderInsideV, borderInsideVWidth, borderInsideVNone),
            TextBold = textBold,
            TextItalic = textItalic,
            TextColor = textColor,
            TextFontSize = textFontSize
        };
    }

    private string? ExtractBorderColor(Drawing.TableCellBorders? borders, string side, StyleResolver? styleResolver)
    {
        if (borders == null) return null;

        Drawing.Outline? outline = side switch
        {
            "top" => borders.TopBorder?.Outline,
            "bottom" => borders.BottomBorder?.Outline,
            "left" => borders.LeftBorder?.Outline,
            "right" => borders.RightBorder?.Outline,
            _ => null
        };

        if (outline == null) return null;

        var solidFill = outline.Elements<Drawing.SolidFill>().FirstOrDefault();
        if (solidFill != null)
        {
            return ExtractSolidFillColor(solidFill, styleResolver);
        }

        return null;
    }

    private (string? Color, double? Width, bool IsNone) ExtractBorderInfo(OpenXmlElement? border, StyleResolver? styleResolver)
    {
        if (border == null) return (null, null, false);

        var outline = border.GetFirstChild<Drawing.Outline>();
        if (outline == null) return (null, null, false);

        // Check for noFill
        var noFill = outline.Elements<Drawing.NoFill>().FirstOrDefault();
        if (noFill != null) return (null, null, true);

        var width = outline.Width != null ? outline.Width.Value / 12700.0 : (double?)null; // EMU to pt
        var color = ExtractBorderColorFromOutline(outline, styleResolver);

        return (color, width, false);
    }

    private static OpenXmlElement? GetTableBorderByLocalName(Drawing.TableCellBorders? borders, string localName)
        => borders?.ChildElements.FirstOrDefault(e => string.Equals(e.LocalName, localName, StringComparison.OrdinalIgnoreCase));

    private static TableBorderState ToBorderState(string? color, double? width, bool isNone)
        => isNone ? TableBorderState.None : (color != null || width.HasValue ? TableBorderState.Visible : TableBorderState.Inherit);

    private static void ApplyDirectTableTextStyle(
        OpenXmlElement tcTxStyle,
        StyleResolver? styleResolver,
        ref bool? textBold,
        ref bool? textItalic,
        ref string? textColor,
        ref double? textFontSize)
    {
        var xml = tcTxStyle.OuterXml;
        var b = System.Text.RegularExpressions.Regex.Match(xml, @"\sb=""([01]|true|false|on|off)""", RegexOptions.IgnoreCase);
        if (b.Success) textBold = ParseOnOff(b.Groups[1].Value);
        var i = System.Text.RegularExpressions.Regex.Match(xml, @"\si=""([01]|true|false|on|off)""", RegexOptions.IgnoreCase);
        if (i.Success) textItalic = ParseOnOff(i.Groups[1].Value);
        var sz = System.Text.RegularExpressions.Regex.Match(xml, @"\ssz=""(\d+)""");
        if (sz.Success && int.TryParse(sz.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var szValue))
            textFontSize = szValue / 100.0;

        var srgb = System.Text.RegularExpressions.Regex.Match(xml, @"<[^>]*srgbClr[^>]*\sval=""([0-9A-Fa-f]{6})""");
        if (srgb.Success)
        {
            textColor = "#" + srgb.Groups[1].Value.ToUpperInvariant();
            return;
        }

        var scheme = System.Text.RegularExpressions.Regex.Match(xml, @"<[^>]*schemeClr[^>]*\sval=""([^""\s/>]+)""");
        if (scheme.Success)
        {
            var resolved = styleResolver?.ResolveSchemeColor(scheme.Groups[1].Value);
            if (!string.IsNullOrEmpty(resolved))
                textColor = resolved.StartsWith('#') ? resolved : "#" + resolved;
        }
    }

    private static bool ParseOnOff(string value)
        => value == "1"
            || value.Equals("true", StringComparison.OrdinalIgnoreCase)
            || value.Equals("on", StringComparison.OrdinalIgnoreCase);

    private static string? ExtractBorderColorFromOutline(Drawing.Outline outline, StyleResolver? styleResolver)
    {
        var solidFill = outline.Elements<Drawing.SolidFill>().FirstOrDefault();
        if (solidFill != null)
        {
            return ExtractSolidFillColorStatic(solidFill, styleResolver);
        }
        return null;
    }

    private string? ExtractSolidFillColor(Drawing.SolidFill solidFill, StyleResolver? styleResolver)
        => ExtractSolidFillColorStatic(solidFill, styleResolver);

    private static string? ExtractSolidFillColorStatic(Drawing.SolidFill solidFill, StyleResolver? styleResolver)
    {
        var rgb = solidFill.RgbColorModelHex;
        if (rgb?.Val != null)
        {
            var color = ParseHexColor(rgb.Val.Value!);
            return ApplyColorModifiers(color, rgb.ChildElements);
        }

        var schemeColor = solidFill.SchemeColor;
        if (schemeColor != null)
        {
            string? schemeName = null;
            // Use regex as primary method — SDK enum parsing is unreliable
            var match = System.Text.RegularExpressions.Regex.Match(
                schemeColor.OuterXml,
                @"val=""([^""]+)""");
            if (match.Success)
            {
                schemeName = match.Groups[1].Value;
            }

            if (!string.IsNullOrEmpty(schemeName))
            {
                var resolved = styleResolver?.ResolveSchemeColor(schemeName);
                if (!string.IsNullOrEmpty(resolved))
                {
                    var color = ParseHexColor(resolved);
                    return ApplyColorModifiers(color, schemeColor.ChildElements);
                }
            }
        }

        return null;
    }

    private static TableStylePart? ResolveCellStylePart(int row, int col, int rowCount, int colCount, TableStyleDefinition? style, bool firstRowFlag, bool bandRowFlag, bool firstColFlag, bool lastColFlag, bool lastRowFlag)
    {
        if (style == null) return null;
        TableStylePart? result = null;

        if (style.Parts.TryGetValue("wholeTbl", out var wholeTbl))
            result = OverlayTableStylePart(result, wholeTbl);

        if (bandRowFlag)
        {
            var bandKey = (row % 2 == 1) ? "band1H" : "band2H";
            if (style.Parts.TryGetValue(bandKey, out var bandPart))
                result = OverlayTableStylePart(result, bandPart);
        }

        if (firstRowFlag && row == 0 && style.Parts.TryGetValue("firstRow", out var firstRow))
            result = OverlayTableStylePart(result, firstRow);
        if (lastRowFlag && row == rowCount - 1 && style.Parts.TryGetValue("lastRow", out var lastRow))
            result = OverlayTableStylePart(result, lastRow);
        if (firstColFlag && col == 0 && style.Parts.TryGetValue("firstCol", out var firstCol))
            result = OverlayTableStylePart(result, firstCol);
        if (lastColFlag && col == colCount - 1 && style.Parts.TryGetValue("lastCol", out var lastCol))
            result = OverlayTableStylePart(result, lastCol);

        return result;
    }

    private static TableStylePart OverlayTableStylePart(TableStylePart? basePart, TableStylePart overlay)
    {
        if (basePart == null) return overlay;

        return new TableStylePart
        {
            BackgroundColor = overlay.BackgroundCleared ? null : (overlay.BackgroundColor ?? basePart.BackgroundColor),
            BackgroundCleared = overlay.BackgroundCleared || (basePart.BackgroundCleared && overlay.BackgroundColor == null),
            BorderTopColor = Overlay(overlay.BorderTopState, overlay.BorderTopColor, basePart.BorderTopColor),
            BorderBottomColor = Overlay(overlay.BorderBottomState, overlay.BorderBottomColor, basePart.BorderBottomColor),
            BorderLeftColor = Overlay(overlay.BorderLeftState, overlay.BorderLeftColor, basePart.BorderLeftColor),
            BorderRightColor = Overlay(overlay.BorderRightState, overlay.BorderRightColor, basePart.BorderRightColor),
            BorderInsideHColor = Overlay(overlay.BorderInsideHState, overlay.BorderInsideHColor, basePart.BorderInsideHColor),
            BorderInsideVColor = Overlay(overlay.BorderInsideVState, overlay.BorderInsideVColor, basePart.BorderInsideVColor),
            BorderTopWidth = Overlay(overlay.BorderTopState, overlay.BorderTopWidth, basePart.BorderTopWidth),
            BorderBottomWidth = Overlay(overlay.BorderBottomState, overlay.BorderBottomWidth, basePart.BorderBottomWidth),
            BorderLeftWidth = Overlay(overlay.BorderLeftState, overlay.BorderLeftWidth, basePart.BorderLeftWidth),
            BorderRightWidth = Overlay(overlay.BorderRightState, overlay.BorderRightWidth, basePart.BorderRightWidth),
            BorderInsideHWidth = Overlay(overlay.BorderInsideHState, overlay.BorderInsideHWidth, basePart.BorderInsideHWidth),
            BorderInsideVWidth = Overlay(overlay.BorderInsideVState, overlay.BorderInsideVWidth, basePart.BorderInsideVWidth),
            BorderTopState = OverlayState(basePart.BorderTopState, overlay.BorderTopState),
            BorderBottomState = OverlayState(basePart.BorderBottomState, overlay.BorderBottomState),
            BorderLeftState = OverlayState(basePart.BorderLeftState, overlay.BorderLeftState),
            BorderRightState = OverlayState(basePart.BorderRightState, overlay.BorderRightState),
            BorderTopExplicit = basePart.BorderTopExplicit || overlay.BorderTopExplicit,
            BorderBottomExplicit = basePart.BorderBottomExplicit || overlay.BorderBottomExplicit,
            BorderLeftExplicit = basePart.BorderLeftExplicit || overlay.BorderLeftExplicit,
            BorderRightExplicit = basePart.BorderRightExplicit || overlay.BorderRightExplicit,
            BorderInsideHState = OverlayState(basePart.BorderInsideHState, overlay.BorderInsideHState),
            BorderInsideVState = OverlayState(basePart.BorderInsideVState, overlay.BorderInsideVState),
            BorderTopNone = OverlayState(basePart.BorderTopState, overlay.BorderTopState) == TableBorderState.None,
            BorderBottomNone = OverlayState(basePart.BorderBottomState, overlay.BorderBottomState) == TableBorderState.None,
            BorderLeftNone = OverlayState(basePart.BorderLeftState, overlay.BorderLeftState) == TableBorderState.None,
            BorderRightNone = OverlayState(basePart.BorderRightState, overlay.BorderRightState) == TableBorderState.None,
            BorderInsideHNone = OverlayState(basePart.BorderInsideHState, overlay.BorderInsideHState) == TableBorderState.None,
            BorderInsideVNone = OverlayState(basePart.BorderInsideVState, overlay.BorderInsideVState) == TableBorderState.None,
            TextBold = overlay.TextBold ?? basePart.TextBold,
            TextItalic = overlay.TextItalic ?? basePart.TextItalic,
            TextColor = overlay.TextColor ?? basePart.TextColor,
            TextFontSize = overlay.TextFontSize ?? basePart.TextFontSize
        };
    }

    private static T? Overlay<T>(TableBorderState overlayState, T? overlay, T? inherited)
        where T : struct
        => overlayState == TableBorderState.Inherit ? inherited : overlay;

    private static string? Overlay(TableBorderState overlayState, string? overlay, string? inherited)
        => overlayState == TableBorderState.Inherit ? inherited : overlay;

    private static TableBorderState OverlayState(TableBorderState inherited, TableBorderState overlay)
        => overlay == TableBorderState.Inherit ? inherited : overlay;

    private TableStylePart? MergeExplicitCellBorders(Drawing.TableCell cell, TableStylePart? stylePart, StyleResolver? styleResolver)
    {
        var cellProps = cell.TableCellProperties;
        if (cellProps == null) return stylePart;

        // CT_TableCellProperties carries cell borders as direct a:lnL/a:lnR/a:lnT/a:lnB
        // children of a:tcPr (a:tcBdr only exists inside tableStyles.xml cell styles,
        // never in a:tcPr — reading tcBdr here silently dropped every explicit cell border).
        var explicitLeft = ExtractCellBorderLine(GetChildByLocalName(cellProps, "lnL"), styleResolver);
        var explicitRight = ExtractCellBorderLine(GetChildByLocalName(cellProps, "lnR"), styleResolver);
        var explicitTop = ExtractCellBorderLine(GetChildByLocalName(cellProps, "lnT"), styleResolver);
        var explicitBottom = ExtractCellBorderLine(GetChildByLocalName(cellProps, "lnB"), styleResolver);

        // If no explicit borders at all, return style part as-is. Missing ln* sides
        // must remain inherited and must not synthesize per-cell stroke state.
        if (explicitTop.State == TableBorderState.Inherit &&
            explicitBottom.State == TableBorderState.Inherit &&
            explicitLeft.State == TableBorderState.Inherit &&
            explicitRight.State == TableBorderState.Inherit)
            return stylePart;

        return new TableStylePart
        {
            BackgroundColor = stylePart?.BackgroundColor,
            BackgroundCleared = stylePart?.BackgroundCleared ?? false,
            BorderTopColor = explicitTop.State == TableBorderState.None ? null : (explicitTop.Color ?? stylePart?.BorderTopColor),
            BorderBottomColor = explicitBottom.State == TableBorderState.None ? null : (explicitBottom.Color ?? stylePart?.BorderBottomColor),
            BorderLeftColor = explicitLeft.State == TableBorderState.None ? null : (explicitLeft.Color ?? stylePart?.BorderLeftColor),
            BorderRightColor = explicitRight.State == TableBorderState.None ? null : (explicitRight.Color ?? stylePart?.BorderRightColor),
            BorderInsideHColor = stylePart?.BorderInsideHColor,
            BorderInsideVColor = stylePart?.BorderInsideVColor,
            BorderTopWidth = explicitTop.State == TableBorderState.None ? null : (explicitTop.Width ?? stylePart?.BorderTopWidth),
            BorderBottomWidth = explicitBottom.State == TableBorderState.None ? null : (explicitBottom.Width ?? stylePart?.BorderBottomWidth),
            BorderLeftWidth = explicitLeft.State == TableBorderState.None ? null : (explicitLeft.Width ?? stylePart?.BorderLeftWidth),
            BorderRightWidth = explicitRight.State == TableBorderState.None ? null : (explicitRight.Width ?? stylePart?.BorderRightWidth),
            BorderInsideHWidth = stylePart?.BorderInsideHWidth,
            BorderInsideVWidth = stylePart?.BorderInsideVWidth,
            BorderTopState = explicitTop.State == TableBorderState.Inherit ? stylePart?.BorderTopState ?? TableBorderState.Inherit : explicitTop.State,
            BorderBottomState = explicitBottom.State == TableBorderState.Inherit ? stylePart?.BorderBottomState ?? TableBorderState.Inherit : explicitBottom.State,
            BorderLeftState = explicitLeft.State == TableBorderState.Inherit ? stylePart?.BorderLeftState ?? TableBorderState.Inherit : explicitLeft.State,
            BorderRightState = explicitRight.State == TableBorderState.Inherit ? stylePart?.BorderRightState ?? TableBorderState.Inherit : explicitRight.State,
            BorderTopExplicit = explicitTop.State != TableBorderState.Inherit,
            BorderBottomExplicit = explicitBottom.State != TableBorderState.Inherit,
            BorderLeftExplicit = explicitLeft.State != TableBorderState.Inherit,
            BorderRightExplicit = explicitRight.State != TableBorderState.Inherit,
            BorderInsideHState = stylePart?.BorderInsideHState ?? TableBorderState.Inherit,
            BorderInsideVState = stylePart?.BorderInsideVState ?? TableBorderState.Inherit,
            BorderTopNone = (explicitTop.State == TableBorderState.Inherit ? stylePart?.BorderTopState ?? TableBorderState.Inherit : explicitTop.State) == TableBorderState.None,
            BorderBottomNone = (explicitBottom.State == TableBorderState.Inherit ? stylePart?.BorderBottomState ?? TableBorderState.Inherit : explicitBottom.State) == TableBorderState.None,
            BorderLeftNone = (explicitLeft.State == TableBorderState.Inherit ? stylePart?.BorderLeftState ?? TableBorderState.Inherit : explicitLeft.State) == TableBorderState.None,
            BorderRightNone = (explicitRight.State == TableBorderState.Inherit ? stylePart?.BorderRightState ?? TableBorderState.Inherit : explicitRight.State) == TableBorderState.None,
            BorderInsideHNone = stylePart?.BorderInsideHNone ?? false,
            BorderInsideVNone = stylePart?.BorderInsideVNone ?? false,
            TextBold = stylePart?.TextBold,
            TextItalic = stylePart?.TextItalic,
            TextColor = stylePart?.TextColor,
            TextFontSize = stylePart?.TextFontSize
        };
    }

    private static OpenXmlElement? GetChildByLocalName(OpenXmlElement element, string localName)
        => element.ChildElements.FirstOrDefault(e => string.Equals(e.LocalName, localName, StringComparison.OrdinalIgnoreCase));

    /// <summary>
    /// Reads one explicit cell border edge (an <c>a:lnL</c>/<c>a:lnR</c>/<c>a:lnT</c>/<c>a:lnB</c>
    /// line-properties element under <c>a:tcPr</c>): absence → Inherit, <c>a:noFill</c> → None,
    /// anything else → Visible with the declared width (EMU) and solid fill color.
    /// Dash styles are not modeled yet; any present line is emitted solid.
    /// </summary>
    private static (TableBorderState State, string? Color, double? Width) ExtractCellBorderLine(OpenXmlElement? line, StyleResolver? styleResolver)
    {
        if (line == null)
            return (TableBorderState.Inherit, null, null);

        if (line.Elements<Drawing.NoFill>().Any())
            return (TableBorderState.None, null, null);

        double? width = null;
        var widthAttr = GetAttributeValue(line, "w");
        if (long.TryParse(widthAttr, NumberStyles.Integer, CultureInfo.InvariantCulture, out var emu))
            width = EmuToPt(emu);

        var solidFill = line.Elements<Drawing.SolidFill>().FirstOrDefault();
        var color = solidFill != null ? ExtractSolidFillColorStatic(solidFill, styleResolver) : null;

        return (TableBorderState.Visible, color, width);
    }

    private static TableStylePart? ApplyTableGridBorders(TableStylePart? stylePart, int row, int col, int rowCount, int colCount)
    {
        if (stylePart == null) return null;

        if (stylePart.BorderTopState == TableBorderState.Inherit &&
            stylePart.BorderBottomState == TableBorderState.Inherit &&
            stylePart.BorderLeftState == TableBorderState.Inherit &&
            stylePart.BorderRightState == TableBorderState.Inherit &&
            stylePart.BorderInsideHState == TableBorderState.Inherit &&
            stylePart.BorderInsideVState == TableBorderState.Inherit)
            return stylePart;

        var topState = stylePart.BorderTopExplicit ? stylePart.BorderTopState : (row == 0 ? stylePart.BorderTopState : TableBorderState.None);
        var bottomState = stylePart.BorderBottomExplicit ? stylePart.BorderBottomState : (row == rowCount - 1 ? stylePart.BorderBottomState : stylePart.BorderInsideHState);
        var leftState = stylePart.BorderLeftExplicit ? stylePart.BorderLeftState : (col == 0 ? stylePart.BorderLeftState : TableBorderState.None);
        var rightState = stylePart.BorderRightExplicit ? stylePart.BorderRightState : (col == colCount - 1 ? stylePart.BorderRightState : stylePart.BorderInsideVState);

        // This cell participates in per-cell stroke emission (at least one edge state is
        // defined). An edge that is still Inherit here is defined nowhere — not by the
        // cell, not by the table style, not by inside-border defaults. PowerPoint renders
        // such edges with no stroke ("No Style, No Grid" semantics), so force None instead
        // of leaving the edge to Typst's default table grid. This is what makes partial
        // stroke tables (e.g. horizontal-only rules) emit exactly the defined edges.
        topState = topState == TableBorderState.Inherit ? TableBorderState.None : topState;
        bottomState = bottomState == TableBorderState.Inherit ? TableBorderState.None : bottomState;
        leftState = leftState == TableBorderState.Inherit ? TableBorderState.None : leftState;
        rightState = rightState == TableBorderState.Inherit ? TableBorderState.None : rightState;

        return new TableStylePart
        {
            BackgroundColor = stylePart.BackgroundColor,
            BackgroundCleared = stylePart.BackgroundCleared,
            BorderTopColor = stylePart.BorderTopColor,
            BorderBottomColor = stylePart.BorderBottomExplicit ? stylePart.BorderBottomColor : (row == rowCount - 1 ? stylePart.BorderBottomColor : stylePart.BorderInsideHColor),
            BorderLeftColor = stylePart.BorderLeftColor,
            BorderRightColor = stylePart.BorderRightExplicit ? stylePart.BorderRightColor : (col == colCount - 1 ? stylePart.BorderRightColor : stylePart.BorderInsideVColor),
            BorderTopWidth = stylePart.BorderTopWidth,
            BorderBottomWidth = stylePart.BorderBottomExplicit ? stylePart.BorderBottomWidth : (row == rowCount - 1 ? stylePart.BorderBottomWidth : stylePart.BorderInsideHWidth),
            BorderLeftWidth = stylePart.BorderLeftWidth,
            BorderRightWidth = stylePart.BorderRightExplicit ? stylePart.BorderRightWidth : (col == colCount - 1 ? stylePart.BorderRightWidth : stylePart.BorderInsideVWidth),
            BorderTopState = topState,
            BorderBottomState = bottomState,
            BorderLeftState = leftState,
            BorderRightState = rightState,
            BorderTopExplicit = stylePart.BorderTopExplicit,
            BorderBottomExplicit = stylePart.BorderBottomExplicit,
            BorderLeftExplicit = stylePart.BorderLeftExplicit,
            BorderRightExplicit = stylePart.BorderRightExplicit,
            BorderInsideHState = stylePart.BorderInsideHState,
            BorderInsideVState = stylePart.BorderInsideVState,
            BorderTopNone = topState == TableBorderState.None,
            BorderBottomNone = bottomState == TableBorderState.None,
            BorderLeftNone = leftState == TableBorderState.None,
            BorderRightNone = rightState == TableBorderState.None,
            BorderInsideHNone = stylePart.BorderInsideHNone,
            BorderInsideVNone = stylePart.BorderInsideVNone,
            TextBold = stylePart.TextBold,
            TextItalic = stylePart.TextItalic,
            TextColor = stylePart.TextColor,
            TextFontSize = stylePart.TextFontSize
        };
    }

    private static (byte R, byte G, byte B) ParseHexColor(string hex)
    {
        hex = hex.TrimStart('#');
        if (hex.Length == 6)
        {
            return (
                (byte)int.Parse(hex.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
                (byte)int.Parse(hex.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
                (byte)int.Parse(hex.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture)
            );
        }
        return (0, 0, 0);
    }

    private static string FormatHexColor((byte R, byte G, byte B) color, byte alpha = 255)
    {
        if (alpha == 255)
            return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
        return $"#{color.R:X2}{color.G:X2}{color.B:X2}{alpha:X2}";
    }

    private static (byte R, byte G, byte B) ApplyTint((byte R, byte G, byte B) color, int tint)
    {
        // DrawingML tint values specify how much of the source color to keep;
        // the remainder is blended toward white. Blend in linear light so very
        // light tints match PowerPoint's rendered colors more closely.
        double sourceWeight = Math.Clamp(tint / 100000.0, 0.0, 1.0);
        return (
            ApplyTintChannel(color.R, sourceWeight),
            ApplyTintChannel(color.G, sourceWeight),
            ApplyTintChannel(color.B, sourceWeight)
        );
    }

    private static byte ApplyTintChannel(byte channel, double sourceWeight)
    {
        var linear = Math.Pow(channel / 255.0, 2.2);
        var tinted = linear * sourceWeight + (1.0 - sourceWeight);
        return (byte)Math.Round(Math.Pow(tinted, 1.0 / 2.2) * 255.0, MidpointRounding.AwayFromZero);
    }

    private static byte ApplyAlpha(int alpha)
    {
        return (byte)(alpha / 100000.0 * 255);
    }

    private static string? ApplyColorModifiers((byte R, byte G, byte B) color, OpenXmlElementList modifiers)
    {
        byte alpha = 255;
        foreach (var mod in modifiers)
        {
            switch (mod.LocalName)
            {
                case "tint":
                    if (TryGetOoxmlVal(mod, out var tintVal))
                        color = ApplyTint(color, tintVal);
                    break;
                case "alpha":
                    if (TryGetOoxmlVal(mod, out var alphaVal))
                        alpha = ApplyAlpha(alphaVal);
                    break;
            }
        }
        // Fully transparent solid fill is visually identical to a:noFill — report no
        // color so callers treat the fill/stroke/text color as absent.
        return alpha == 0 ? null : FormatHexColor(color, alpha);
    }

    private static bool TryGetOoxmlVal(OpenXmlElement element, out int value)
    {
        var match = Regex.Match(element.OuterXml, @"\bval=""(-?\d+)""");
        return int.TryParse(
            match.Success ? match.Groups[1].Value : null,
            NumberStyles.Integer,
            CultureInfo.InvariantCulture,
            out value);
    }

    private (double X, double Y, double Width, double Height, double Rotation) GetElementPosition(ShapeProperties? shapeProperties)
    {
        if (shapeProperties?.Transform2D == null)
            return (0, 0, 100, 50, 0);

        var transform = shapeProperties.Transform2D;
        var x = transform.Offset?.X?.Value ?? 0;
        var y = transform.Offset?.Y?.Value ?? 0;
        var width = transform.Extents?.Cx?.Value ?? 100;
        var height = transform.Extents?.Cy?.Value ?? 50;
        var rotation = transform.Rotation?.Value ?? 0;

        return (EmuToPt((long)x), EmuToPt((long)y), EmuToPt((long)width), EmuToPt((long)height), rotation / 60000.0);
    }

    private (double X, double Y, double Width, double Height, double Rotation)? GetLayoutPlaceholderPosition(SlidePart slidePart, P.Shape shape)
    {
        var layoutPart = slidePart.SlideLayoutPart;
        if (layoutPart?.SlideLayout?.CommonSlideData?.ShapeTree == null)
            return null;

        // Get placeholder info from the slide shape
        var slidePh = GetPlaceholderInfo(shape);
        if (slidePh == null)
            return null;

        // Find matching placeholder in layout
        foreach (var layoutShape in layoutPart.SlideLayout.CommonSlideData.ShapeTree.ChildElements.OfType<P.Shape>())
        {
            var layoutPh = GetPlaceholderInfo(layoutShape);
            if (layoutPh == null)
                continue;

            var slideType = slidePh.Value.Type;
            var slideIdx = slidePh.Value.Index;
            var layoutType = layoutPh.Value.Type;
            var layoutIdx = layoutPh.Value.Index;

            // Match by type or index
            bool matches = !string.IsNullOrEmpty(slideType) && slideType == layoutType;
            if (!matches && slideIdx.HasValue && layoutIdx.HasValue)
                matches = slideIdx.Value == layoutIdx.Value;

            if (matches)
            {
                var layoutXfrm = layoutShape.ShapeProperties?.Transform2D;
                if (layoutXfrm != null)
                {
                    var x = layoutXfrm.Offset?.X?.Value ?? 0;
                    var y = layoutXfrm.Offset?.Y?.Value ?? 0;
                    var width = layoutXfrm.Extents?.Cx?.Value ?? 100;
                    var height = layoutXfrm.Extents?.Cy?.Value ?? 50;
                    var rotation = layoutXfrm.Rotation?.Value ?? 0;
                    return (EmuToPt((long)x), EmuToPt((long)y), EmuToPt((long)width), EmuToPt((long)height), rotation / 60000.0);
                }
            }
        }

        return null;
    }

    private (string? Type, int? Index)? GetPlaceholderInfo(P.Shape shape)
    {
        var nvSpPr = shape.NonVisualShapeProperties;
        if (nvSpPr == null) return null;

        PlaceholderShape? ph = nvSpPr.Elements<PlaceholderShape>().FirstOrDefault();
        if (ph == null)
        {
            var appProps = nvSpPr.ApplicationNonVisualDrawingProperties;
            ph = appProps?.Elements<PlaceholderShape>().FirstOrDefault();
        }

        if (ph == null) return null;

        string? type = null;
        int? idx = null;

        var outerXml = ph.OuterXml;
        if (!string.IsNullOrEmpty(outerXml))
        {
            var typeMatch = System.Text.RegularExpressions.Regex.Match(outerXml, @"type\s*=\s*""([^""]*)""");
            if (typeMatch.Success)
                type = typeMatch.Groups[1].Value;

            var idxMatch = System.Text.RegularExpressions.Regex.Match(outerXml, @"idx\s*=\s*""([^""]*)""");
            if (idxMatch.Success && int.TryParse(idxMatch.Groups[1].Value, out var parsedIdx))
                idx = parsedIdx;
        }

        return (type, idx);
    }

    private (double X, double Y, double Width, double Height) GetGraphicFramePosition(P.GraphicFrame graphicFrame)
    {
        var transform = graphicFrame.Transform;
        if (transform == null)
            return (0, 0, 100, 50);

        var x = transform.Offset?.X?.Value ?? 0;
        var y = transform.Offset?.Y?.Value ?? 0;
        var width = transform.Extents?.Cx?.Value ?? 100;
        var height = transform.Extents?.Cy?.Value ?? 50;

        return (EmuToPt((long)x), EmuToPt((long)y), EmuToPt((long)width), EmuToPt((long)height));
    }

    private (uint Id, string Name) GetElementIdAndName(OpenXmlElement? nvProperties)
    {
        if (nvProperties == null) return (0, "Unknown");

        var cnvPr = nvProperties.ChildElements.FirstOrDefault(e => e.LocalName == "cNvPr");
        if (cnvPr == null) return (0, "Unknown");

        uint id = 0;
        var idAttr = cnvPr.GetAttribute("id", "");
        if (!string.IsNullOrEmpty(idAttr.Value) && uint.TryParse(idAttr.Value, out var parsedId))
        {
            id = parsedId;
        }

        var nameAttr = cnvPr.GetAttribute("name", "");
        var name = !string.IsNullOrEmpty(nameAttr.Value) ? nameAttr.Value : "Unknown";

        return (id, name);
    }

    private void ExtractFonts()
    {
        var fontPartIndex = 0;

        var embeddedFonts = _document.PresentationPart!.Presentation!
            .Descendants()
            .Where(e => e.LocalName == "embeddedFont")
            .ToList();

        foreach (var embeddedFont in embeddedFonts)
        {
            var family = embeddedFont.ChildElements.FirstOrDefault(e => e.LocalName == "font")?.GetAttribute("typeface", "").Value;
            foreach (var fontReference in embeddedFont.ChildElements.Where(e => e.LocalName is "regular" or "bold" or "italic" or "boldItalic"))
            {
                var relationshipId = fontReference.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships").Value;
                if (string.IsNullOrEmpty(relationshipId)) continue;

                if (_document.PresentationPart.GetPartById(relationshipId) is FontPart fontPart)
                {
                    ExtractFontFile(fontPart, fontPartIndex, family);
                    fontPartIndex++;
                }
            }
        }

        if (fontPartIndex > 0)
            return;

        // Extract fonts from all parts
        var allParts = new List<OpenXmlPart>();
        allParts.Add(_document.PresentationPart!);
        allParts.AddRange(_document.PresentationPart!.SlideParts);
        
        foreach (var part in allParts)
        {
            foreach (var p in part.Parts)
            {
                if (p.OpenXmlPart is FontPart fontPart)
                {
                    ExtractFontFile(fontPart, fontPartIndex, family: null);
                    fontPartIndex++;
                }
            }
        }
    }

    private void ExtractFontFile(FontPart fontPart, int index, string? family)
    {
        using var stream = fontPart.GetStream();
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        var data = ms.ToArray();
        
        // Try to extract actual font from EOT wrapper
        var otfPos = data.AsSpan().IndexOf(System.Text.Encoding.ASCII.GetBytes("OTTO"));
        var ttfPos = data.AsSpan().IndexOf(new byte[] { 0x00, 0x01, 0x00, 0x00 });
        
        int pos;
        string ext;
        if (otfPos >= 0)
        {
            pos = otfPos;
            ext = ".otf";
        }
        else if (ttfPos >= 0)
        {
            pos = ttfPos;
            ext = ".ttf";
        }
        else
        {
            // Can't extract, save as-is
            var fileName = $"font{index}.fntdata";
            File.WriteAllBytes(Path.Combine(_fontsDirectory, fileName), data);
            return;
        }
        
        var fontDataLength = TryReadEotFontDataLength(data, pos);
        var fontData = data[pos..(pos + fontDataLength)];
        var fontName = $"font{index}{ext}";
        File.WriteAllBytes(Path.Combine(_fontsDirectory, fontName), fontData);

        var metrics = OpenTypeFontMetricsReader.TryRead(fontData);
        if (metrics != null && !string.IsNullOrWhiteSpace(family))
        {
            _fontMetrics[family] = metrics;
            if (family.EndsWith(" Bold", StringComparison.OrdinalIgnoreCase))
            {
                _fontMetrics.TryAdd(family[..^5], metrics);
            }
        }
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

    private Dictionary<string, string> ExtractThemeFonts()
    {
        var themeFonts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        
        // Try to get theme from the first slide's master
        var firstSlideId = _document.PresentationPart!.Presentation!.SlideIdList?.ChildElements.OfType<SlideId>().FirstOrDefault();
        if (firstSlideId == null) return themeFonts;
        
        var firstSlidePart = (SlidePart)_document.PresentationPart.GetPartById(firstSlideId.RelationshipId!);
        var layoutPart = firstSlidePart.SlideLayoutPart;
        var masterPart = layoutPart?.SlideMasterPart;
        var themePart = masterPart?.ThemePart;
        
        if (themePart?.Theme?.ThemeElements?.FontScheme == null)
            return themeFonts;
        
        var fontScheme = themePart.Theme.ThemeElements.FontScheme;
        
        // Major font (used for titles, headings)
        var majorLatin = fontScheme.MajorFont?.LatinFont?.Typeface?.Value;
        if (!string.IsNullOrEmpty(majorLatin))
        {
            themeFonts["+mj-lt"] = majorLatin;
            themeFonts["major-latin"] = majorLatin;
        }
        
        // Minor font (used for body text)
        var minorLatin = fontScheme.MinorFont?.LatinFont?.Typeface?.Value;
        if (!string.IsNullOrEmpty(minorLatin))
        {
            themeFonts["+mn-lt"] = minorLatin;
            themeFonts["minor-latin"] = minorLatin;
        }
        
        return themeFonts;
    }

    private static double EmuToPt(long emu)
    {
        return emu / 12700.0;
    }

    /// <summary>OOXML default for <c>lIns</c>/<c>rIns</c> on <c>a:bodyPr</c> (0.1").</summary>
    private const int DefaultHorizontalInsetEmu = 91440;

    /// <summary>OOXML default for <c>tIns</c>/<c>bIns</c> on <c>a:bodyPr</c> (0.05").</summary>
    private const int DefaultVerticalInsetEmu = 45720;

    private static string FormatPt(double pt)
    {
        return pt.ToString("F2", CultureInfo.InvariantCulture) + "pt";
    }

    private static string? GetAttributeValue(OpenXmlElement? element, string attributeName)
    {
        if (element == null)
            return null;

        var match = System.Text.RegularExpressions.Regex.Match(
            element.OuterXml,
            $@"\b{System.Text.RegularExpressions.Regex.Escape(attributeName)}\s*=\s*""([^""]*)""");

        return match.Success ? match.Groups[1].Value : null;
    }

    private string ResolveThemeFont(string? fontRef)
    {
        if (string.IsNullOrEmpty(fontRef))
            return "Arial";
        
        // Check if it's a theme font reference
        if (_themeFonts.TryGetValue(fontRef, out var resolvedFont))
            return resolvedFont;
        
        return fontRef;
    }

    private static string SubstituteUnavailableFont(string fontFamily, HashSet<string> availableFonts)
    {
        if (availableFonts.Contains(fontFamily))
            return fontFamily;

        if (fontFamily.StartsWith("Aptos", StringComparison.OrdinalIgnoreCase))
        {
            if (availableFonts.Contains("Aptos"))
                return "Aptos";

            if (availableFonts.Contains("Carlito"))
                return "Carlito";
        }

        if (fontFamily.Equals("Calibri", StringComparison.OrdinalIgnoreCase))
        {
            if (availableFonts.Contains("Aptos"))
                return "Aptos";

            if (availableFonts.Contains("Carlito"))
                return "Carlito";
        }

        return fontFamily;
    }

    private static string EscapeTypstText(string text)
    {
        if (string.IsNullOrEmpty(text)) return "";
        
        return text
            .Replace("\\", "\\\\")
            .Replace("[", "\\[")
            .Replace("]", "\\]")
            .Replace("*", "\\*")
            .Replace("_", "\\_")
            .Replace("#", "\\#")
            .Replace("`", "\\`")
            .Replace("$", "\\$")
            .Replace("%", "\\%")
            .Replace("&", "\\&")
            .Replace("@", "\\@")
            .Replace("^", "\\^")
            .Replace("~", "\\~")
            .Replace("<", "\\<")
            .Replace(">", "\\>")
            .Replace("/", "\\/")
            .Replace("{", "\\{")
            .Replace("}", "\\}");
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDirectory))
            {
                Directory.Delete(_tempDirectory, true);
            }
        }
        catch
        {
            // Best effort cleanup
        }
    }
}
