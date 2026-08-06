using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using PptxEditor.Core.Generation.Model;
using PptxEditor.Core.Models;

namespace PptxEditor.Core.Generation.Schema;

/// <summary>
/// Parser + validator for the generation JSON vocabulary.
/// Reads with System.Text.Json, builds the <see cref="GenerationDocument"/> model and
/// collects loud, actionable errors (JSON path + suggestion) instead of failing fast, so
/// AI and human authors can fix every problem at once. A pre-scan seeds the id registry
/// with every explicit 'id' before the validation walk, so a synthesized fallback id never
/// shadows an authored id regardless of document order. CSS-isms (flex-wrap, z-index,
/// percentages, wrap, …) are rejected explicitly — they are v1 non-goals.
/// </summary>
public sealed class GenerationDocumentParser
{
    /// <summary>The only vocabulary version supported by this parser.</summary>
    public const string SupportedVersion = "2.0";

    private static readonly Regex HexColorPattern = new(@"^#[0-9a-fA-F]{6}$", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex PercentPattern = new(@"^-?\d+(?:\.\d+)?%$", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex IdPattern = new(@"^[a-z][a-z0-9_-]*$", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly IReadOnlyDictionary<string, LayoutMode> LayoutModes = new Dictionary<string, LayoutMode>(StringComparer.OrdinalIgnoreCase)
    {
        ["row"] = LayoutMode.Row, ["column"] = LayoutMode.Column, ["grid"] = LayoutMode.Grid
    };

    private static readonly IReadOnlyDictionary<string, Justify> Justifies = new Dictionary<string, Justify>(StringComparer.OrdinalIgnoreCase)
    {
        ["start"] = Justify.Start, ["center"] = Justify.Center, ["end"] = Justify.End,
        ["space-between"] = Justify.SpaceBetween, ["space-evenly"] = Justify.SpaceEvenly
    };

    private static readonly IReadOnlyDictionary<string, AlignItems> Alignments = new Dictionary<string, AlignItems>(StringComparer.OrdinalIgnoreCase)
    {
        ["start"] = AlignItems.Start, ["center"] = AlignItems.Center, ["end"] = AlignItems.End, ["stretch"] = AlignItems.Stretch
    };

    private static readonly IReadOnlyDictionary<string, OverflowPolicy> OverflowPolicies = new Dictionary<string, OverflowPolicy>(StringComparer.OrdinalIgnoreCase)
    {
        ["error"] = OverflowPolicy.Error, ["shrink"] = OverflowPolicy.Shrink, ["clip"] = OverflowPolicy.Clip
    };

    private static readonly IReadOnlyDictionary<string, TextAlign> TextAligns = new Dictionary<string, TextAlign>(StringComparer.OrdinalIgnoreCase)
    {
        ["left"] = TextAlign.Left, ["center"] = TextAlign.Center, ["right"] = TextAlign.Right
    };

    private static readonly IReadOnlyDictionary<string, TextAnchor> TextAnchors = new Dictionary<string, TextAnchor>(StringComparer.OrdinalIgnoreCase)
    {
        ["top"] = TextAnchor.Top, ["middle"] = TextAnchor.Middle, ["bottom"] = TextAnchor.Bottom
    };

    private static readonly IReadOnlyDictionary<string, LineOrientation> LineOrientations = new Dictionary<string, LineOrientation>(StringComparer.OrdinalIgnoreCase)
    {
        ["horizontal"] = LineOrientation.Horizontal, ["vertical"] = LineOrientation.Vertical
    };

    private static readonly IReadOnlyDictionary<string, ImageFitMode> ImageFits = new Dictionary<string, ImageFitMode>(StringComparer.OrdinalIgnoreCase)
    {
        ["fill"] = ImageFitMode.Fill, ["crop"] = ImageFitMode.Crop, ["contain"] = ImageFitMode.Contain
    };

    private static readonly IReadOnlyDictionary<string, CardStyle> CardStyles = new Dictionary<string, CardStyle>(StringComparer.OrdinalIgnoreCase)
    {
        ["flat"] = CardStyle.Flat, ["outline"] = CardStyle.Outline, ["shadow"] = CardStyle.Shadow
    };

    private static readonly IReadOnlyDictionary<string, SlideSize> SlideSizes = new Dictionary<string, SlideSize>(StringComparer.OrdinalIgnoreCase)
    {
        ["16:9"] = SlideSize.Widescreen16x9, ["4:3"] = SlideSize.Standard4x3
    };

    private static readonly IReadOnlyDictionary<string, string> ComponentNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        ["card"] = "card", ["kpi"] = "kpi", ["title_block"] = "title_block", ["bullet_list"] = "bullet_list",
        ["divider"] = "divider", ["badge"] = "badge", ["image_card"] = "image_card", ["table_block"] = "table_block"
    };

    // P10 (archetype slide functions): slide-root-only types. The parser
    // wraps them as a bare container holding one archetype-named component marker; the
    // archetype layer (ArchetypeExpander) owns expansion. Archetype names are NOT valid
    // child element types.
    private static readonly IReadOnlyDictionary<string, string> ArchetypeSlideNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        ["cover"] = "cover", ["section"] = "section", ["kpi_row"] = "kpi_row", ["two_col"] = "two_col", ["table_slide"] = "table_slide"
    };

    private static readonly IReadOnlyDictionary<string, string> CssIsms = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        ["flexwrap"] = "CSS 'flex-wrap' is not supported: layout never wraps (v1 non-goal).",
        ["wrap"] = "'wrap' is not supported: layout never wraps (v1 non-goal).",
        ["overflowwrap"] = "CSS 'overflow-wrap' is not supported: layout never wraps (v1 non-goal).",
        ["zindex"] = "CSS 'z-index' is not supported: paint order = document order (v1 non-goal).",
        ["position"] = "CSS 'position' is not supported: use 'at' ({\"x\":…,\"y\":…}) on children of layout-less parents.",
        ["display"] = "CSS 'display' is not supported: declare layout via 'layout' ({\"mode\":\"row\"|\"column\"|\"grid\"}).",
        ["float"] = "CSS floats are not supported: use row/column/grid layout.",
        ["clear"] = "CSS floats are not supported: use row/column/grid layout.",
        ["margin"] = "CSS margins are not supported: use container 'padding' and layout 'gap'.",
        ["flex"] = "CSS 'flex' is not supported: use 'size' ({\"grow\": n}).",
        ["flexdirection"] = "CSS 'flex-direction' is not supported: use layout mode 'row' or 'column'.",
        ["justifycontent"] = "CSS 'justify-content' is not supported: use layout 'justify'.",
        ["alignitems"] = "CSS 'align-items' is not supported: use layout 'align'.",
        ["width"] = "CSS 'width' is not supported: use 'size' ({\"w\": pt}).",
        ["height"] = "CSS 'height' is not supported: use 'size' ({\"h\": pt}).",
        ["top"] = "CSS absolute offsets are not supported: use 'at' ({\"x\":…,\"y\":…}).",
        ["left"] = "CSS absolute offsets are not supported: use 'at' ({\"x\":…,\"y\":…}).",
        ["right"] = "CSS absolute offsets are not supported: use 'at' ({\"x\":…,\"y\":…}).",
        ["bottom"] = "CSS absolute offsets are not supported: use 'at' ({\"x\":…,\"y\":…}).",
        ["transform"] = "CSS 'transform' is not supported in v1.",
        ["style"] = "CSS/HTML input is a v1 non-goal: style via design tokens and element properties.",
        ["class"] = "CSS/HTML input is a v1 non-goal: style via design tokens and element properties.",
        ["classname"] = "CSS/HTML input is a v1 non-goal: style via design tokens and element properties.",
        ["gridtemplatecolumns"] = "CSS grid templates are not supported: use 'layout' ({\"mode\":\"grid\",\"cols\": n}).",
        ["gridtemplaterows"] = "CSS grid templates are not supported: use 'layout' ({\"mode\":\"grid\",\"cols\": n})."
    };

    private static readonly IReadOnlySet<string> RootProps = Set("version", "slideSize", "design", "slides");
    private static readonly IReadOnlySet<string> DesignProps = Set("palette", "fonts", "shape", "metrics");
    private static readonly IReadOnlySet<string> FontsProps = Set("display", "body");
    private static readonly IReadOnlySet<string> ShapeProps = Set("cornerRadius", "cardStyle");
    private static readonly IReadOnlySet<string> MetricsProps = Set("marginPt", "gutterPt", "titleSizePt", "bodySizePt");
    private static readonly IReadOnlySet<string> LayoutProps = Set("mode", "gap", "cols", "rowGap", "columnGap", "justify", "align");
    private static readonly IReadOnlySet<string> SizeProps = Set("w", "h", "grow", "aspect", "alignSelf");
    private static readonly IReadOnlySet<string> AtProps = Set("x", "y");
    private static readonly IReadOnlySet<string> ContainerProps = Set("type", "id", "layout", "padding", "overflow", "children", "fill", "stroke", "radius", "shadow", "size", "at", "notes");
    private static readonly IReadOnlySet<string> TextProps = Set("type", "id", "text", "runs", "font", "fontSize", "color", "bold", "italic", "textAlign", "anchor", "insets", "overflow", "shadow", "size", "at");
    private static readonly IReadOnlySet<string> RunProps = Set("text", "font", "fontSize", "color", "bold", "italic");
    private static readonly IReadOnlySet<string> RectProps = Set("type", "id", "fill", "stroke", "radius", "shadow", "size", "at", "overflow");
    private static readonly IReadOnlySet<string> EllipseProps = Set("type", "id", "fill", "stroke", "shadow", "size", "at", "overflow");
    private static readonly IReadOnlySet<string> LineProps = Set("type", "id", "orientation", "stroke", "size", "at", "overflow");
    private static readonly IReadOnlySet<string> ImageProps = Set("type", "id", "src", "fit", "crop", "alt", "size", "at", "overflow");
    private static readonly IReadOnlySet<string> CropProps = Set("left", "top", "right", "bottom");
    private static readonly IReadOnlySet<string> GroupProps = Set("type", "id", "children", "size", "at", "overflow");
    private static readonly IReadOnlySet<string> ComponentProps = Set("type", "id", "content", "size", "at");
    private static readonly IReadOnlySet<string> ArchetypeSlideProps = Set("type", "id", "content", "notes");
    private static readonly IReadOnlySet<string> GradientProps = Set("angle", "stops");
    private static readonly IReadOnlySet<string> StopProps = Set("color", "offset", "alpha");
    private static readonly IReadOnlySet<string> StrokeProps = Set("color", "width");
    private static readonly IReadOnlySet<string> RadiusProps = Set("tl", "tr", "br", "bl");
    private static readonly IReadOnlySet<string> ShadowProps = Set("color", "dx", "dy", "blur", "alpha");

    private readonly List<GenerationIssue> _errors = [];
    private readonly List<GenerationIssue> _warnings = [];
    private IReadOnlyDictionary<string, string> _palette = new Dictionary<string, string>();

    // Document-wide id registry: every id (user-authored or parser-synthesized) maps to
    // the JSON path that first declared it, so duplicates fail loudly with both paths.
    // Explicit ids are seeded by a pre-scan pass (SeedExplicitIds) before any parsing, so
    // an authored id always wins over a synthesized fallback regardless of document order.
    private readonly Dictionary<string, string> _ids = new(StringComparer.Ordinal);

    // Paths whose id was synthesized as a fallback rather than authored, so a duplicate
    // error can mark the claimant "(auto-generated)" instead of sending users to hunt for
    // an "id" property that was never written.
    private readonly HashSet<string> _synthesizedPaths = new(StringComparer.Ordinal);

    // Per-element-type counter for synthesized fallback ids ("text-0", "card-1", …). The
    // counter is per Validate() call, so re-parsing the same JSON yields identical ids.
    private readonly Dictionary<string, int> _typeCounters = new(StringComparer.Ordinal);

    /// <summary>
    /// Validates a generation JSON document in a single pass. Never throws for contract
    /// violations — inspect <see cref="GenerationValidationResult.Errors"/> instead.
    /// </summary>
    public GenerationValidationResult Validate(string json)
    {
        ArgumentNullException.ThrowIfNull(json);
        _errors.Clear();
        _warnings.Clear();
        _palette = new Dictionary<string, string>();
        _ids.Clear();
        _synthesizedPaths.Clear();
        _typeCounters.Clear();

        JsonDocument document;
        try
        {
            document = JsonDocument.Parse(json);
        }
        catch (JsonException ex)
        {
            _errors.Add(new GenerationIssue("$", $"malformed JSON: {ex.Message}", null, GenerationIssueSeverity.Error));
            return Result(null);
        }

        using (document)
        {
            var root = document.RootElement;
            if (root.ValueKind != JsonValueKind.Object)
            {
                _errors.Add(new GenerationIssue("$", "the document root must be a JSON object ({\"version\":…,\"design\":…,\"slides\":[…]}).", null, GenerationIssueSeverity.Error));
                return Result(null);
            }

            // Pass 1: seed the registry with every explicit id before any element is parsed,
            // so an authored id later in the document is never shadowed by a fallback that
            // an earlier un-ided element synthesized (explicit user ids always win).
            SeedExplicitIds(root);

            // Pass 2: parse, synthesizing fallbacks only for elements without explicit ids.
            var parsed = ParseDocument(root);
            return Result(_errors.Count == 0 ? parsed : null);
        }
    }

    /// <summary>
    /// Validates and returns the document, throwing <see cref="GenerationValidationException"/>
    /// listing every error when the document is invalid.
    /// </summary>
    public GenerationDocument Parse(string json) => Validate(json).ThrowIfInvalid().Document!;

    private GenerationValidationResult Result(GenerationDocument? document) =>
        new(document, [.. _errors], [.. _warnings]);

    private static IReadOnlySet<string> Set(params string[] names) => new HashSet<string>(names, StringComparer.OrdinalIgnoreCase);

    private void Error(string path, string message, string? suggestion = null) =>
        _errors.Add(new GenerationIssue(path, message, suggestion, GenerationIssueSeverity.Error));

    private void Warn(string path, string message) =>
        _warnings.Add(new GenerationIssue(path, message, null, GenerationIssueSeverity.Warning));

    private GenerationDocument ParseDocument(JsonElement root)
    {
        CheckUnknownProps(root, "$", "the document root", RootProps);

        string? version = null;
        if (!TryGet(root, "version", out var versionEl))
        {
            Error("$", $"'version' is required (\"{SupportedVersion}\").");
        }
        else
        {
            version = StringValue(versionEl, "$.version");
            if (version is not null && version != SupportedVersion)
            {
                Error("$.version", $"unsupported version '{version}'; expected \"{SupportedVersion}\".");
            }
        }

        var slideSize = SlideSize.Widescreen16x9;
        if (TryGet(root, "slideSize", out var slideSizeEl))
        {
            slideSize = EnumValue(slideSizeEl, "$.slideSize", SlideSizes, "slide size") ?? slideSize;
        }

        DesignTokens design;
        if (!TryGet(root, "design", out var designEl))
        {
            Error("$", "'design' is required (design tokens).");
            design = new DesignTokens { Palette = new Dictionary<string, string>() };
        }
        else if (designEl.ValueKind != JsonValueKind.Object)
        {
            Error("$.design", "must be an object ({\"palette\":…,\"fonts\":…,\"shape\":…,\"metrics\":…}).");
            design = new DesignTokens { Palette = new Dictionary<string, string>() };
        }
        else
        {
            design = ParseDesign(designEl, "$.design");
        }
        _palette = design.Palette;

        var slides = new List<ContainerElement>();
        if (!TryGet(root, "slides", out var slidesEl))
        {
            Error("$", "'slides' is required (at least one slide).");
        }
        else if (slidesEl.ValueKind != JsonValueKind.Array)
        {
            Error("$.slides", "must be an array of container elements.");
        }
        else
        {
            var index = 0;
            foreach (var slideEl in slidesEl.EnumerateArray())
            {
                var slide = ParseSlide(slideEl, $"$.slides[{index}]");
                if (slide is not null)
                {
                    slides.Add(slide);
                }
                index++;
            }
            if (index == 0)
            {
                Error("$.slides", "at least one slide is required.");
            }
        }

        return new GenerationDocument
        {
            Version = version ?? SupportedVersion,
            Design = design,
            Slides = slides,
            SlideSize = slideSize
        };
    }

    private DesignTokens ParseDesign(JsonElement el, string path)
    {
        CheckUnknownProps(el, path, "design", DesignProps);

        var palette = new Dictionary<string, string>(StringComparer.Ordinal);
        if (!TryGet(el, "palette", out var paletteEl))
        {
            Error(path, "'palette' is required (token name → #RRGGBB).");
        }
        else if (paletteEl.ValueKind != JsonValueKind.Object)
        {
            Error($"{path}.palette", "must be an object mapping token names to #RRGGBB colors.");
        }
        else
        {
            foreach (var color in paletteEl.EnumerateObject())
            {
                var colorPath = $"{path}.palette.{color.Name}";
                if (color.Value.ValueKind != JsonValueKind.String || !HexColorPattern.IsMatch(color.Value.GetString() ?? string.Empty))
                {
                    Error(colorPath, "palette colors must be #RRGGBB hex literals.");
                    continue;
                }
                palette[color.Name] = color.Value.GetString()!;
            }
        }

        var fonts = new FontTokens();
        if (TryGet(el, "fonts", out var fontsEl))
        {
            if (fontsEl.ValueKind != JsonValueKind.Object)
            {
                Error($"{path}.fonts", "must be an object ({\"display\":…,\"body\":…}).");
            }
            else
            {
                CheckUnknownProps(fontsEl, $"{path}.fonts", "fonts", FontsProps);
                fonts = new FontTokens
                {
                    Display = StringProp(fontsEl, "display", $"{path}.fonts"),
                    Body = StringProp(fontsEl, "body", $"{path}.fonts")
                };
            }
        }

        var shape = new ShapeTokens();
        if (TryGet(el, "shape", out var shapeEl))
        {
            if (shapeEl.ValueKind != JsonValueKind.Object)
            {
                Error($"{path}.shape", "must be an object ({\"cornerRadius\":…,\"cardStyle\":…}).");
            }
            else
            {
                CheckUnknownProps(shapeEl, $"{path}.shape", "shape", ShapeProps);
                shape = new ShapeTokens
                {
                    CornerRadius = NumberProp(shapeEl, "cornerRadius", $"{path}.shape", min: 0) ?? 0,
                    CardStyle = EnumProp(shapeEl, "cardStyle", $"{path}.shape", CardStyles, CardStyle.Flat, "card style")
                };
            }
        }

        var metrics = new MetricTokens();
        if (TryGet(el, "metrics", out var metricsEl))
        {
            if (metricsEl.ValueKind != JsonValueKind.Object)
            {
                Error($"{path}.metrics", "must be an object of pt values ({\"marginPt\":…,\"gutterPt\":…,\"titleSizePt\":…,\"bodySizePt\":…}).");
            }
            else
            {
                CheckUnknownProps(metricsEl, $"{path}.metrics", "metrics", MetricsProps);
                metrics = new MetricTokens
                {
                    MarginPt = NumberProp(metricsEl, "marginPt", $"{path}.metrics", min: 0) ?? 43,
                    GutterPt = NumberProp(metricsEl, "gutterPt", $"{path}.metrics", min: 0) ?? 18,
                    TitleSizePt = NumberProp(metricsEl, "titleSizePt", $"{path}.metrics", minExclusive: 0) ?? 30,
                    BodySizePt = NumberProp(metricsEl, "bodySizePt", $"{path}.metrics", minExclusive: 0) ?? 14
                };
            }
        }

        return new DesignTokens { Palette = palette, Fonts = fonts, Shape = shape, Metrics = metrics };
    }

    private ContainerElement? ParseSlide(JsonElement el, string path)
    {
        if (el.ValueKind != JsonValueKind.Object)
        {
            Error(path, "each slide must be a JSON object (a container element).");
            return null;
        }
        if (!TryGet(el, "type", out var typeEl) || typeEl.ValueKind != JsonValueKind.String)
        {
            Error(path, "each slide's root requires \"type\": \"container\" (the root is the slide canvas) or an archetype slide type (cover, section, kpi_row, two_col, table_slide).");
            return null;
        }
        var slideType = typeEl.GetString()!;
        if (ArchetypeSlideNames.TryGetValue(slideType, out var archetype))
        {
            return ParseArchetypeSlide(el, path, archetype);
        }
        if (!string.Equals(slideType, "container", StringComparison.OrdinalIgnoreCase))
        {
            Error(path, $"each slide's root must be a 'container' element or an archetype slide type ({string.Join(", ", ArchetypeSlideNames.Values)}) (got '{slideType}').");
            return null;
        }
        return ParseContainer(el, path, isRoot: true, parentHasLayout: false);
    }

    private ContainerElement ParseArchetypeSlide(JsonElement el, string path, string name)
    {
        CheckUnknownProps(el, path, $"a '{name}' archetype slide", ArchetypeSlideProps);
        var id = IdProp(el, path, name);

        JsonElement? content = null;
        if (TryGet(el, "content", out var contentEl))
        {
            if (contentEl.ValueKind != JsonValueKind.Object)
            {
                Error($"{path}.content", "must be an object (archetype payload).");
            }
            else
            {
                content = contentEl.Clone();
            }
        }

        // Marker shape the archetype layer (P10) recognizes: a bare root container whose
        // only child is the archetype-named component node. ArchetypeExpander composes it.
        // Both the root and the marker carry the slide id; expansion propagates it to the
        // composed slide and derives role-based child ids from it.
        return new ContainerElement
        {
            Id = id,
            Children = [new ComponentElement { Id = id, Name = name, Content = content }],
            Notes = NotesProp(el, path, isRoot: true)
        };
    }

    private GenElement? ParseElement(JsonElement el, string path, bool parentHasLayout)
    {
        if (el.ValueKind != JsonValueKind.Object)
        {
            Error(path, "each element must be a JSON object with a 'type'.");
            return null;
        }
        if (!TryGet(el, "type", out var typeEl) || typeEl.ValueKind != JsonValueKind.String)
        {
            Error(path, "'type' is required (container|text|rect|ellipse|line|connector|image|group, or a component: card, kpi, title_block, bullet_list, divider, badge, image_card, table_block).");
            return null;
        }

        var type = typeEl.GetString()!;
        switch (type.ToLowerInvariant())
        {
            case "container":
                return ParseContainer(el, path, isRoot: false, parentHasLayout: parentHasLayout);
            case "text":
                return ParseText(el, path, parentHasLayout);
            case "rect":
                return ParseRect(el, path, parentHasLayout);
            case "ellipse":
                return ParseEllipse(el, path, parentHasLayout);
            case "line":
                return ParseLine(el, path, parentHasLayout, isConnector: false);
            case "connector":
                return ParseLine(el, path, parentHasLayout, isConnector: true);
            case "image":
                return ParseImage(el, path, parentHasLayout);
            case "group":
                return ParseGroup(el, path, parentHasLayout);
            default:
                if (ComponentNames.TryGetValue(type, out var canonical))
                {
                    return ParseComponent(el, path, canonical, parentHasLayout);
                }
                if (ArchetypeSlideNames.ContainsKey(type))
                {
                    Error(path, $"'{type}' is an archetype slide type: it is only valid as a slide root, not as a child element.");
                    return null;
                }
                Error(
                    path,
                    $"unknown element type '{type}'. Expected container|text|rect|ellipse|line|connector|image|group, or a component: {string.Join(", ", ComponentNames.Values)}.",
                    Suggest(type, ContainerishTypeNames()));
                return null;
        }
    }

    private static IEnumerable<string> ContainerishTypeNames() =>
        new[] { "container", "text", "rect", "ellipse", "line", "connector", "image", "group" }.Concat(ComponentNames.Values);

    private ContainerElement ParseContainer(JsonElement el, string path, bool isRoot, bool parentHasLayout)
    {
        CheckUnknownProps(el, path, "a container", ContainerProps);
        var id = IdProp(el, path, isRoot ? "slide" : "container");
        var (size, at) = ParseSizeAndAt(el, path, isRoot, parentHasLayout);

        LayoutSpec? layout = null;
        var hasLayoutProp = TryGet(el, "layout", out var layoutEl);
        if (hasLayoutProp)
        {
            if (layoutEl.ValueKind != JsonValueKind.Object)
            {
                Error($"{path}.layout", "must be an object ({\"mode\":\"row\"|\"column\"|\"grid\"}).");
            }
            else
            {
                layout = ParseLayout(layoutEl, $"{path}.layout");
            }
        }

        var children = ParseChildren(el, path, hasLayoutProp);

        return new ContainerElement
        {
            Id = id,
            Size = size,
            At = at,
            Layout = layout,
            Padding = InsetsProp(el, "padding", path),
            Overflow = EnumProp(el, "overflow", path, OverflowPolicies, OverflowPolicy.Error, "overflow policy"),
            Children = children ?? [],
            Fill = FillProp(el, path),
            Stroke = StrokeProp(el, path),
            Radius = RadiusProp(el, path),
            Shadow = ShadowProp(el, path),
            Notes = NotesProp(el, path, isRoot)
        };
    }

    private TextElement ParseText(JsonElement el, string path, bool parentHasLayout)
    {
        CheckUnknownProps(el, path, "a text element", TextProps);
        var id = IdProp(el, path, "text");
        var (size, at) = ParseSizeAndAt(el, path, isRoot: false, parentHasLayout: parentHasLayout);

        var hasText = TryGet(el, "text", out var textEl);
        var hasRuns = TryGet(el, "runs", out var runsEl);
        string? value = null;
        List<TextRun>? runs = null;
        if (hasText == hasRuns)
        {
            Error(path, "a text element needs exactly one of 'text' (string) or 'runs' (array).");
        }
        if (hasText)
        {
            if (textEl.ValueKind != JsonValueKind.String)
            {
                Error($"{path}.text", "must be a string.");
            }
            else
            {
                value = textEl.GetString()!;
            }
        }
        if (hasRuns)
        {
            runs = ParseRuns(runsEl, $"{path}.runs");
        }

        return new TextElement
        {
            Id = id,
            Size = size,
            At = at,
            Value = value,
            Runs = runs,
            Font = StringProp(el, "font", path),
            FontSize = NumberProp(el, "fontSize", path, minExclusive: 0),
            Color = ColorProp(el, "color", path),
            Bold = BoolProp(el, "bold", path),
            Italic = BoolProp(el, "italic", path),
            TextAlign = EnumProp(el, "textAlign", path, TextAligns, Model.TextAlign.Left, "text alignment"),
            Anchor = EnumProp(el, "anchor", path, TextAnchors, TextAnchor.Top, "text anchor"),
            Insets = InsetsProp(el, "insets", path),
            Overflow = EnumProp(el, "overflow", path, OverflowPolicies, OverflowPolicy.Shrink, "overflow policy"),
            Shadow = ShadowProp(el, path)
        };
    }

    private List<TextRun>? ParseRuns(JsonElement el, string path)
    {
        if (el.ValueKind != JsonValueKind.Array)
        {
            Error(path, "must be an array of run objects ({\"text\":…,\"bold\":…,\"color\":…}).");
            return null;
        }
        var runs = new List<TextRun>();
        var index = 0;
        foreach (var runEl in el.EnumerateArray())
        {
            var runPath = $"{path}[{index}]";
            if (runEl.ValueKind != JsonValueKind.Object)
            {
                Error(runPath, "each run must be an object ({\"text\":…}).");
            }
            else
            {
                CheckUnknownProps(runEl, runPath, "a text run", RunProps);
                var text = StringProp(runEl, "text", runPath, required: true);
                runs.Add(new TextRun
                {
                    Text = text ?? string.Empty,
                    Font = StringProp(runEl, "font", runPath),
                    FontSize = NumberProp(runEl, "fontSize", runPath, minExclusive: 0),
                    Color = ColorProp(runEl, "color", runPath),
                    Bold = BoolProp(runEl, "bold", runPath),
                    Italic = BoolProp(runEl, "italic", runPath)
                });
            }
            index++;
        }
        if (runs.Count == 0)
        {
            Error(path, "'runs' must contain at least one run.");
        }
        return runs;
    }

    private RectElement ParseRect(JsonElement el, string path, bool parentHasLayout)
    {
        CheckUnknownProps(el, path, "a rect element", RectProps);
        RejectOverflow(el, path);
        var id = IdProp(el, path, "rect");
        var (size, at) = ParseSizeAndAt(el, path, isRoot: false, parentHasLayout: parentHasLayout);
        return new RectElement
        {
            Id = id,
            Size = size,
            At = at,
            Fill = FillProp(el, path),
            Stroke = StrokeProp(el, path),
            Radius = RadiusProp(el, path),
            Shadow = ShadowProp(el, path)
        };
    }

    private EllipseElement ParseEllipse(JsonElement el, string path, bool parentHasLayout)
    {
        CheckUnknownProps(el, path, "an ellipse element", EllipseProps);
        RejectOverflow(el, path);
        var id = IdProp(el, path, "ellipse");
        var (size, at) = ParseSizeAndAt(el, path, isRoot: false, parentHasLayout: parentHasLayout);
        return new EllipseElement
        {
            Id = id,
            Size = size,
            At = at,
            Fill = FillProp(el, path),
            Stroke = StrokeProp(el, path),
            Shadow = ShadowProp(el, path)
        };
    }

    private LineElement ParseLine(JsonElement el, string path, bool parentHasLayout, bool isConnector)
    {
        CheckUnknownProps(el, path, isConnector ? "a connector" : "a line element", LineProps);
        RejectOverflow(el, path);
        var id = IdProp(el, path, isConnector ? "connector" : "line");
        var (size, at) = ParseSizeAndAt(el, path, isRoot: false, parentHasLayout: parentHasLayout);
        return new LineElement
        {
            Id = id,
            Size = size,
            At = at,
            IsConnector = isConnector,
            Orientation = EnumProp(el, "orientation", path, LineOrientations, LineOrientation.Horizontal, "line orientation"),
            Stroke = StrokeProp(el, path)
        };
    }

    private ImageElement ParseImage(JsonElement el, string path, bool parentHasLayout)
    {
        CheckUnknownProps(el, path, "an image element", ImageProps);
        RejectOverflow(el, path);
        var id = IdProp(el, path, "image");
        var (size, at) = ParseSizeAndAt(el, path, isRoot: false, parentHasLayout: parentHasLayout);

        var src = StringProp(el, "src", path, required: true, allowEmpty: false);

        SourceRect? crop = null;
        if (TryGet(el, "crop", out var cropEl))
        {
            if (cropEl.ValueKind != JsonValueKind.Object)
            {
                Error($"{path}.crop", "must be an object ({\"left\":…,\"top\":…,\"right\":…,\"bottom\":…}) in 1/1000ths of a percent.");
            }
            else
            {
                CheckUnknownProps(cropEl, $"{path}.crop", "a crop rectangle", CropProps);
                crop = new SourceRect(
                    IntProp(cropEl, "left", $"{path}.crop", required: true, min: 0, max: 100000) ?? 0,
                    IntProp(cropEl, "top", $"{path}.crop", required: true, min: 0, max: 100000) ?? 0,
                    IntProp(cropEl, "right", $"{path}.crop", required: true, min: 0, max: 100000) ?? 0,
                    IntProp(cropEl, "bottom", $"{path}.crop", required: true, min: 0, max: 100000) ?? 0);
            }
        }

        return new ImageElement
        {
            Id = id,
            Size = size,
            At = at,
            Source = src ?? string.Empty,
            Fit = EnumProp(el, "fit", path, ImageFits, ImageFitMode.Fill, "image fit mode"),
            Crop = crop,
            Alt = StringProp(el, "alt", path)
        };
    }

    private GroupElement ParseGroup(JsonElement el, string path, bool parentHasLayout)
    {
        CheckUnknownProps(el, path, "a group", GroupProps);
        RejectOverflow(el, path);
        var id = IdProp(el, path, "group");
        var (size, at) = ParseSizeAndAt(el, path, isRoot: false, parentHasLayout: parentHasLayout);
        var children = ParseChildren(el, path, parentHasLayout: false);
        return new GroupElement { Id = id, Size = size, At = at, Children = children ?? [] };
    }

    private ComponentElement ParseComponent(JsonElement el, string path, string name, bool parentHasLayout)
    {
        CheckUnknownProps(el, path, $"a '{name}' component", ComponentProps);
        var id = IdProp(el, path, name);
        var (size, at) = ParseSizeAndAt(el, path, isRoot: false, parentHasLayout: parentHasLayout);

        JsonElement? content = null;
        if (TryGet(el, "content", out var contentEl))
        {
            if (contentEl.ValueKind != JsonValueKind.Object)
            {
                Error($"{path}.content", "must be an object (component payload).");
            }
            else
            {
                content = contentEl.Clone();
            }
        }

        return new ComponentElement { Id = id, Size = size, At = at, Name = name, Content = content };
    }

    private void RejectOverflow(JsonElement el, string path)
    {
        if (TryGet(el, "overflow", out _))
        {
            Error($"{path}.overflow", "'overflow' is only valid on containers and text elements.");
        }
    }

    /// <summary>
    /// Reads slide-level speaker notes ('notes', a string). Notes are per-slide metadata:
    /// only the slide root carries them — a nested container with 'notes' is a loud error
    /// rather than silently-dropped data.
    /// </summary>
    private string? NotesProp(JsonElement el, string path, bool isRoot)
    {
        if (!TryGet(el, "notes", out var notesEl))
        {
            return null;
        }
        var notesPath = $"{path}.notes";
        if (!isRoot)
        {
            Error(notesPath, "'notes' is only valid on the slide root (speaker notes are per-slide metadata).");
            return null;
        }
        return StringValue(notesEl, notesPath);
    }

    /// <summary>
    /// Pass-1 pre-scan: registers every EXPLICIT 'id' in the document (with the JSON path
    /// that first declares it, in document order) before any element is parsed. Mirrors the
    /// exact walk of pass 2 — slides, then container/group children — so a synthesized
    /// fallback can never claim a name an author wrote anywhere in the document, no matter
    /// where that id appears relative to the un-ided element.
    /// </summary>
    private void SeedExplicitIds(JsonElement root)
    {
        if (!TryGet(root, "slides", out var slidesEl) || slidesEl.ValueKind != JsonValueKind.Array)
        {
            return;
        }
        var index = 0;
        foreach (var slideEl in slidesEl.EnumerateArray())
        {
            SeedExplicitIdsInElement(slideEl, $"$.slides[{index}]", isSlide: true);
            index++;
        }
    }

    private void SeedExplicitIdsInElement(JsonElement el, string path, bool isSlide)
    {
        if (el.ValueKind != JsonValueKind.Object || !TryGet(el, "type", out var typeEl) || typeEl.ValueKind != JsonValueKind.String)
        {
            return;
        }
        var type = typeEl.GetString()!;

        if (isSlide)
        {
            // Slide roots are 'container' elements or archetype slide types (ParseSlide).
            // An archetype's content is an opaque payload — no element ids live inside it.
            if (!ArchetypeSlideNames.ContainsKey(type) && !string.Equals(type, "container", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }
            SeedExplicitId(el, path);
            if (!string.Equals(type, "container", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }
        }
        else
        {
            // Every element type carries an optional id (IdProp). Only containers and
            // groups nest further elements; component payloads are opaque to the registry.
            SeedExplicitId(el, path);
            if (!string.Equals(type, "container", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(type, "group", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }
        }

        if (TryGet(el, "children", out var childrenEl) && childrenEl.ValueKind == JsonValueKind.Array)
        {
            var childIndex = 0;
            foreach (var childEl in childrenEl.EnumerateArray())
            {
                SeedExplicitIdsInElement(childEl, $"{path}.children[{childIndex}]", isSlide: false);
                childIndex++;
            }
        }
    }

    private void SeedExplicitId(JsonElement el, string path)
    {
        if (!TryGet(el, "id", out var idEl) || idEl.ValueKind != JsonValueKind.String)
        {
            return; // absent or non-string: non-strings are reported by IdProp with a loud path error
        }
        var value = idEl.GetString()!;
        if (!IdPattern.IsMatch(value))
        {
            return; // malformed: reported by IdProp and never reaches the duplicate check
        }
        _ids.TryAdd(value, path); // first declaration in document order wins
    }

    /// <summary>
    /// Reads the optional stable element identifier ('id'). Validates the string shape
    /// (lowercase start, then letters/digits/'-'/'_') and per-document uniqueness, both
    /// with loud JSON-path errors. When 'id' is absent, synthesizes a deterministic
    /// fallback '<c>{type}-{n}</c>' (a per-type document-order counter) so every element
    /// stays addressable after expansion; explicit user ids always win over synthesized
    /// ones — the pass-1 pre-scan seeded every authored id first, so synthesis skips ids
    /// already taken anywhere in the document.
    /// </summary>
    private string? IdProp(JsonElement el, string path, string type)
    {
        if (!TryGet(el, "id", out var idEl))
        {
            // Synthesize "{type}-{n}", skipping every explicit id (pre-seeded in pass 1) as
            // well as any fallback already handed out during this walk.
            var n = _typeCounters.GetValueOrDefault(type, 0);
            while (true)
            {
                var candidate = $"{type}-{n}";
                n++;
                _typeCounters[type] = n;
                if (!_ids.ContainsKey(candidate))
                {
                    _ids[candidate] = path;
                    _synthesizedPaths.Add(path);
                    return candidate;
                }
            }
        }

        var valuePath = $"{path}.id";
        if (idEl.ValueKind != JsonValueKind.String)
        {
            Error(valuePath, "must be a string (a stable element identifier).");
            return null;
        }
        var value = idEl.GetString()!;
        if (!IdPattern.IsMatch(value))
        {
            Error(valuePath,
                $"'{value}' is not a valid id: lowercase start, then letters, digits, '-' or '_' (e.g. \"hero-title\").");
            return null;
        }
        if (_ids.TryGetValue(value, out var firstPath))
        {
            // Pass 1 seeded this element's own id at this very path — only a collision with
            // a different path is a duplicate. Mark a synthesized claimant so users are not
            // sent hunting for an "id" property that was never written.
            if (!string.Equals(firstPath, path, StringComparison.Ordinal))
            {
                Error(valuePath,
                    $"duplicate id '{value}': first declared at {firstPath}" +
                    $"{( _synthesizedPaths.Contains(firstPath) ? " (auto-generated)" : string.Empty)}. Ids must be unique per document.");
                return null;
            }
        }
        _ids[value] = path;
        return value;
    }

    private List<GenElement>? ParseChildren(JsonElement el, string path, bool parentHasLayout)
    {
        if (!TryGet(el, "children", out var childrenEl))
        {
            Error(path, "'children' is required (an array of elements; use an empty array for a blank container).");
            return null;
        }
        if (childrenEl.ValueKind != JsonValueKind.Array)
        {
            Error($"{path}.children", "must be an array of elements.");
            return null;
        }
        var children = new List<GenElement>();
        var index = 0;
        foreach (var childEl in childrenEl.EnumerateArray())
        {
            var child = ParseElement(childEl, $"{path}.children[{index}]", parentHasLayout);
            if (child is not null)
            {
                children.Add(child);
            }
            index++;
        }
        return children;
    }

    private LayoutSpec? ParseLayout(JsonElement el, string path)
    {
        CheckUnknownProps(el, path, "a layout", LayoutProps);

        LayoutMode? mode = null;
        var hasCols = TryGet(el, "cols", out var colsEl);
        if (!TryGet(el, "mode", out var modeEl))
        {
            Error(path, "'mode' is required (row|column|grid).");
        }
        else
        {
            mode = EnumValue(modeEl, $"{path}.mode", LayoutModes, "layout mode");
        }

        int? cols = null;
        if (hasCols)
        {
            if (mode is not null and not LayoutMode.Grid)
            {
                Error($"{path}.cols", "'cols' is only valid for grid layouts.");
            }
            cols = IntValue(colsEl, $"{path}.cols", min: 1);
        }
        else if (mode == LayoutMode.Grid)
        {
            Error(path, "grid layouts require 'cols' (≥ 1).");
        }

        var rowGap = NumberProp(el, "rowGap", path, min: 0);
        if (rowGap is not null && mode is not null and not LayoutMode.Grid)
        {
            Error($"{path}.rowGap", "'rowGap' is only valid for grid layouts.");
        }
        var columnGap = NumberProp(el, "columnGap", path, min: 0);
        if (columnGap is not null && mode is not null and not LayoutMode.Grid)
        {
            Error($"{path}.columnGap", "'columnGap' is only valid for grid layouts.");
        }

        if (mode is null)
        {
            return null;
        }

        return new LayoutSpec
        {
            Mode = mode.Value,
            Gap = NumberProp(el, "gap", path, min: 0) ?? 0,
            Columns = cols,
            RowGap = rowGap,
            ColumnGap = columnGap,
            Justify = EnumProp(el, "justify", path, Justifies, Justify.Start, "justify value"),
            Align = EnumProp(el, "align", path, Alignments, AlignItems.Stretch, "align value")
        };
    }

    private (SizeSpec? Size, PointSpec? At) ParseSizeAndAt(JsonElement el, string path, bool isRoot, bool parentHasLayout)
    {
        SizeSpec? size = null;
        if (TryGet(el, "size", out var sizeEl))
        {
            if (isRoot)
            {
                Error($"{path}.size", "the root slide container takes its size from the slide; 'size' is only valid on children.");
            }
            else if (sizeEl.ValueKind != JsonValueKind.Object)
            {
                Error($"{path}.size", "must be an object ({\"w\":…,\"h\":…,\"grow\":…,\"aspect\":\"16:9\",\"alignSelf\":…}).");
            }
            else
            {
                size = ParseSize(sizeEl, $"{path}.size");
            }
        }

        PointSpec? at = null;
        if (TryGet(el, "at", out var atEl))
        {
            if (isRoot)
            {
                Error($"{path}.at", "the root slide container fills the slide; 'at' is only valid on children of layout-less parents.");
            }
            else if (atEl.ValueKind != JsonValueKind.Object)
            {
                Error($"{path}.at", "must be an object ({\"x\":…,\"y\":…}).");
            }
            else
            {
                at = ParseAt(atEl, $"{path}.at");
            }
        }

        if (!isRoot)
        {
            if (parentHasLayout)
            {
                if (at is not null)
                {
                    Error($"{path}.at", "'at' is not allowed inside a layout container: position via layout mode, grow, justify and align.");
                }
            }
            else
            {
                if (at is null)
                {
                    Error(path, "elements in a layout-less parent require 'at' ({\"x\":…,\"y\":…}).");
                }
                if (size?.Grow is not null)
                {
                    Error($"{path}.size.grow", "'grow' shares space on a layout axis: it requires a layout container parent.");
                }
                var widthResolved = size is { Width: not null } || size is { Aspect: not null, Height: not null };
                var heightResolved = size is { Height: not null } || size is { Aspect: not null, Width: not null };
                if (!widthResolved || !heightResolved)
                {
                    Error(path, "elements in a layout-less parent require a fixed 'size': \"w\" and \"h\", or one dimension plus \"aspect\".");
                }
            }
        }

        return (size, at);
    }

    private SizeSpec ParseSize(JsonElement el, string path)
    {
        CheckUnknownProps(el, path, "a size constraint", SizeProps);

        AspectRatio? aspect = null;
        if (TryGet(el, "aspect", out var aspectEl))
        {
            aspect = ParseAspect(aspectEl, $"{path}.aspect");
        }

        return new SizeSpec
        {
            Width = NumberProp(el, "w", path, minExclusive: 0),
            Height = NumberProp(el, "h", path, minExclusive: 0),
            Grow = NumberProp(el, "grow", path, min: 0),
            Aspect = aspect,
            AlignSelf = EnumProp(el, "alignSelf", path, Alignments, (AlignItems?)null, "alignSelf value")
        };
    }

    private AspectRatio? ParseAspect(JsonElement el, string path)
    {
        if (el.ValueKind != JsonValueKind.String)
        {
            Error(path, "must be a \"W:H\" string (e.g. \"16:9\").");
            return null;
        }
        var raw = el.GetString()!;
        var parts = raw.Split(':');
        if (parts.Length != 2
            || !double.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out var w)
            || !double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out var h)
            || w <= 0
            || h <= 0)
        {
            Error(path, $"'{raw}' is not a valid aspect ratio: expected \"W:H\" with positive numbers (e.g. \"16:9\").");
            return null;
        }
        return AspectRatio.Of(w, h);
    }

    private PointSpec? ParseAt(JsonElement el, string path)
    {
        CheckUnknownProps(el, path, "an absolute position", AtProps);
        var x = NumberProp(el, "x", path, required: true, min: 0);
        var y = NumberProp(el, "y", path, required: true, min: 0);
        if (x is null || y is null)
        {
            return null;
        }
        return new PointSpec(x.Value, y.Value);
    }

    private EdgeInsets? InsetsProp(JsonElement el, string name, string path)
    {
        if (!TryGet(el, name, out var v))
        {
            return null;
        }
        var valuePath = $"{path}.{name}";
        if (v.ValueKind == JsonValueKind.Number || v.ValueKind == JsonValueKind.String)
        {
            var uniform = NumberValue(v, valuePath, min: 0);
            return uniform is null ? null : EdgeInsets.All(uniform.Value);
        }
        if (v.ValueKind != JsonValueKind.Array)
        {
            Error(valuePath, "must be a number (all edges) or an array [v, h] / [top, right, bottom, left].");
            return null;
        }

        var items = new List<double?>();
        foreach (var item in v.EnumerateArray())
        {
            items.Add(NumberValue(item, valuePath, min: 0));
        }
        if (items.Count != 2 && items.Count != 4)
        {
            Error(valuePath, $"must have 2 ([v, h]) or 4 ([top, right, bottom, left]) values (got {items.Count}).");
            return null;
        }
        if (items.Any(i => i is null))
        {
            return null;
        }
        return items.Count == 2
            ? EdgeInsets.Symmetric(items[0]!.Value, items[1]!.Value)
            : new EdgeInsets(items[0]!.Value, items[1]!.Value, items[2]!.Value, items[3]!.Value);
    }

    private FillSpec? FillProp(JsonElement el, string path)
    {
        if (!TryGet(el, "fill", out var v))
        {
            return null;
        }
        var valuePath = $"{path}.fill";
        if (v.ValueKind == JsonValueKind.String)
        {
            var color = ColorValue(v, valuePath);
            return color is null ? null : new SolidFill(color);
        }
        if (v.ValueKind != JsonValueKind.Object)
        {
            Error(valuePath, "must be a color string (palette token or #RRGGBB) or a linear gradient ({\"angle\":…,\"stops\":[…]}).");
            return null;
        }

        CheckUnknownProps(v, valuePath, "a linear gradient fill", GradientProps);
        var angle = NumberProp(v, "angle", valuePath, required: true) ?? 0;

        var stops = new List<GradientStop>();
        if (!TryGet(v, "stops", out var stopsEl))
        {
            Error(valuePath, "'stops' is required on a gradient fill (at least 2 stops).");
            return null;
        }
        if (stopsEl.ValueKind != JsonValueKind.Array)
        {
            Error($"{valuePath}.stops", "must be an array of ({\"color\":…,\"offset\":0..1,\"alpha\":0..1?}).");
            return null;
        }
        var index = 0;
        foreach (var stopEl in stopsEl.EnumerateArray())
        {
            var stopPath = $"{valuePath}.stops[{index}]";
            if (stopEl.ValueKind != JsonValueKind.Object)
            {
                Error(stopPath, "each gradient stop must be an object ({\"color\":…,\"offset\":…}).");
            }
            else
            {
                CheckUnknownProps(stopEl, stopPath, "a gradient stop", StopProps);
                var color = ColorProp(stopEl, "color", stopPath, required: true);
                var offset = NumberProp(stopEl, "offset", stopPath, required: true, min: 0, max: 1);
                if (color is not null && offset is not null)
                {
                    stops.Add(new GradientStop
                    {
                        Color = color,
                        Offset = offset.Value,
                        Alpha = NumberProp(stopEl, "alpha", stopPath, min: 0, max: 1)
                    });
                }
            }
            index++;
        }
        if (stops.Count < 2)
        {
            Error($"{valuePath}.stops", $"a gradient needs at least 2 stops (got {stops.Count}).");
            return null;
        }
        return new LinearGradientFill { Angle = angle, Stops = stops };
    }

    private StrokeSpec? StrokeProp(JsonElement el, string path)
    {
        if (!TryGet(el, "stroke", out var v))
        {
            return null;
        }
        var valuePath = $"{path}.stroke";
        if (v.ValueKind == JsonValueKind.String)
        {
            var color = ColorValue(v, valuePath);
            return color is null ? null : new StrokeSpec { Color = color };
        }
        if (v.ValueKind != JsonValueKind.Object)
        {
            Error(valuePath, "must be a color string or an object ({\"color\":…,\"width\":…}).");
            return null;
        }

        CheckUnknownProps(v, valuePath, "a stroke", StrokeProps);
        var strokeColor = ColorProp(v, "color", valuePath, required: true);
        if (strokeColor is null)
        {
            return null;
        }
        return new StrokeSpec
        {
            Color = strokeColor,
            WidthPt = NumberProp(v, "width", valuePath, minExclusive: 0) ?? 1
        };
    }

    private CornerRadii? RadiusProp(JsonElement el, string path)
    {
        if (!TryGet(el, "radius", out var v))
        {
            return null;
        }
        var valuePath = $"{path}.radius";
        if (v.ValueKind == JsonValueKind.Number || v.ValueKind == JsonValueKind.String)
        {
            var uniform = NumberValue(v, valuePath, min: 0);
            return uniform is null ? null : CornerRadii.All(uniform.Value);
        }
        if (v.ValueKind != JsonValueKind.Object)
        {
            Error(valuePath, "must be a number (all corners) or an object ({\"tl\":…,\"tr\":…,\"br\":…,\"bl\":…}).");
            return null;
        }

        CheckUnknownProps(v, valuePath, "a corner radius", RadiusProps);
        return new CornerRadii(
            NumberProp(v, "tl", valuePath, min: 0) ?? 0,
            NumberProp(v, "tr", valuePath, min: 0) ?? 0,
            NumberProp(v, "br", valuePath, min: 0) ?? 0,
            NumberProp(v, "bl", valuePath, min: 0) ?? 0);
    }

    private ShadowSpec? ShadowProp(JsonElement el, string path)
    {
        if (!TryGet(el, "shadow", out var v))
        {
            return null;
        }
        var valuePath = $"{path}.shadow";
        if (v.ValueKind != JsonValueKind.Object)
        {
            Error(valuePath, "must be an object ({\"color\":…,\"dx\":…,\"dy\":…,\"blur\":…,\"alpha\":…}).");
            return null;
        }

        CheckUnknownProps(v, valuePath, "a shadow", ShadowProps);
        var color = ColorProp(v, "color", valuePath, required: true);
        if (color is null)
        {
            return null;
        }
        return new ShadowSpec
        {
            Color = color,
            Dx = NumberProp(v, "dx", valuePath) ?? 0,
            Dy = NumberProp(v, "dy", valuePath) ?? 0,
            Blur = NumberProp(v, "blur", valuePath, min: 0) ?? 0,
            Alpha = NumberProp(v, "alpha", valuePath, min: 0, max: 1)
        };
    }

    private void CheckUnknownProps(JsonElement obj, string path, string subject, IReadOnlySet<string> allowed)
    {
        foreach (var property in obj.EnumerateObject())
        {
            if (allowed.Contains(property.Name))
            {
                continue;
            }
            var propertyPath = $"{path}.{property.Name}";
            if (CssIsms.TryGetValue(NormalizeCssKey(property.Name), out var cssMessage))
            {
                Error(propertyPath, cssMessage);
                continue;
            }
            Error(propertyPath, $"unknown property '{property.Name}' on {subject}.", Suggest(property.Name, allowed));
        }
    }

    private static string NormalizeCssKey(string name)
    {
        Span<char> buffer = stackalloc char[name.Length];
        var length = 0;
        foreach (var c in name)
        {
            if (c is not ('-' or '_' or ' '))
            {
                buffer[length++] = char.ToLowerInvariant(c);
            }
        }
        return new string(buffer[..length]);
    }

    private static string? Suggest(string name, IEnumerable<string> candidates)
    {
        string? best = null;
        var bestDistance = int.MaxValue;
        foreach (var candidate in candidates)
        {
            var distance = Levenshtein(name, candidate);
            if (distance < bestDistance)
            {
                bestDistance = distance;
                best = candidate;
            }
        }
        return bestDistance <= 2 ? $"Did you mean '{best}'?" : null;
    }

    private static int Levenshtein(string a, string b)
    {
        a = a.ToLowerInvariant();
        b = b.ToLowerInvariant();
        var previous = new int[b.Length + 1];
        var current = new int[b.Length + 1];
        for (var j = 0; j <= b.Length; j++)
        {
            previous[j] = j;
        }
        for (var i = 1; i <= a.Length; i++)
        {
            current[0] = i;
            for (var j = 1; j <= b.Length; j++)
            {
                var cost = a[i - 1] == b[j - 1] ? 0 : 1;
                current[j] = Math.Min(Math.Min(current[j - 1] + 1, previous[j] + 1), previous[j - 1] + cost);
            }
            (previous, current) = (current, previous);
        }
        return previous[b.Length];
    }

    private static bool TryGet(JsonElement obj, string name, out JsonElement value)
    {
        foreach (var property in obj.EnumerateObject())
        {
            if (string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase))
            {
                value = property.Value;
                return true;
            }
        }
        value = default;
        return false;
    }

    private string? StringValue(JsonElement v, string valuePath)
    {
        if (v.ValueKind != JsonValueKind.String)
        {
            Error(valuePath, "must be a string.");
            return null;
        }
        return v.GetString()!;
    }

    private string? StringProp(JsonElement obj, string name, string path, bool required = false, bool allowEmpty = true)
    {
        if (!TryGet(obj, name, out var v))
        {
            if (required)
            {
                Error(path, $"'{name}' is required.");
            }
            return null;
        }
        var value = StringValue(v, $"{path}.{name}");
        if (value is not null && !allowEmpty && value.Length == 0)
        {
            Error($"{path}.{name}", "must not be empty.");
            return null;
        }
        return value;
    }

    private double? NumberProp(
        JsonElement obj,
        string name,
        string path,
        bool required = false,
        double min = double.NegativeInfinity,
        double max = double.PositiveInfinity,
        double? minExclusive = null)
    {
        if (!TryGet(obj, name, out var v))
        {
            if (required)
            {
                Error(path, $"'{name}' is required.");
            }
            return null;
        }
        return NumberValue(v, $"{path}.{name}", min, max, minExclusive);
    }

    private double? NumberValue(
        JsonElement v,
        string valuePath,
        double min = double.NegativeInfinity,
        double max = double.PositiveInfinity,
        double? minExclusive = null)
    {
        if (v.ValueKind == JsonValueKind.String && PercentPattern.IsMatch(v.GetString()!))
        {
            Error(valuePath, $"'{v.GetString()}' is a percentage: percentages are not supported (v1 non-goal); use pt numbers.");
            return null;
        }
        if (v.ValueKind != JsonValueKind.Number || !v.TryGetDouble(out var number))
        {
            Error(valuePath, "must be a JSON number (pt).");
            return null;
        }
        if (minExclusive is not null && number <= minExclusive.Value)
        {
            Error(valuePath, $"must be > {FormatNumber(minExclusive.Value)} (got {FormatNumber(number)}).");
            return null;
        }
        if (number < min)
        {
            Error(valuePath, $"must be ≥ {FormatNumber(min)} (got {FormatNumber(number)}).");
            return null;
        }
        if (number > max)
        {
            Error(valuePath, $"must be ≤ {FormatNumber(max)} (got {FormatNumber(number)}).");
            return null;
        }
        return number;
    }

    private int? IntProp(JsonElement obj, string name, string path, bool required = false, int min = int.MinValue, int max = int.MaxValue)
    {
        if (!TryGet(obj, name, out var v))
        {
            if (required)
            {
                Error(path, $"'{name}' is required.");
            }
            return null;
        }
        return IntValue(v, $"{path}.{name}", min, max);
    }

    private int? IntValue(JsonElement v, string valuePath, int min = int.MinValue, int max = int.MaxValue)
    {
        if (v.ValueKind == JsonValueKind.String && PercentPattern.IsMatch(v.GetString()!))
        {
            Error(valuePath, $"'{v.GetString()}' is a percentage: percentages are not supported (v1 non-goal); use pt numbers.");
            return null;
        }
        if (v.ValueKind != JsonValueKind.Number || !v.TryGetInt32(out var number))
        {
            Error(valuePath, "must be an integer.");
            return null;
        }
        if (number < min || number > max)
        {
            Error(valuePath, $"must be between {min} and {max} (got {number}).");
            return null;
        }
        return number;
    }

    private bool BoolProp(JsonElement obj, string name, string path, bool defaultValue = false)
    {
        if (!TryGet(obj, name, out var v))
        {
            return defaultValue;
        }
        if (v.ValueKind == JsonValueKind.True)
        {
            return true;
        }
        if (v.ValueKind == JsonValueKind.False)
        {
            return false;
        }
        Error($"{path}.{name}", "must be a boolean.");
        return defaultValue;
    }

    private T? EnumValue<T>(JsonElement v, string valuePath, IReadOnlyDictionary<string, T> map, string what) where T : struct
    {
        if (v.ValueKind != JsonValueKind.String)
        {
            Error(valuePath, $"must be a string ({what}: {string.Join(", ", map.Keys)}).");
            return null;
        }
        var raw = v.GetString()!;
        if (map.TryGetValue(raw, out var value))
        {
            return value;
        }
        Error(valuePath, $"'{raw}' is not a valid {what}. Valid values: {string.Join(", ", map.Keys)}.", Suggest(raw, map.Keys));
        return null;
    }

    private T EnumProp<T>(JsonElement obj, string name, string path, IReadOnlyDictionary<string, T> map, T defaultValue, string what) where T : struct
    {
        if (!TryGet(obj, name, out var v))
        {
            return defaultValue;
        }
        return EnumValue(v, $"{path}.{name}", map, what) ?? defaultValue;
    }

    private T? EnumProp<T>(JsonElement obj, string name, string path, IReadOnlyDictionary<string, T> map, T? defaultValue, string what) where T : struct
    {
        if (!TryGet(obj, name, out var v))
        {
            return defaultValue;
        }
        return EnumValue(v, $"{path}.{name}", map, what);
    }

    private string? ColorValue(JsonElement v, string valuePath)
    {
        if (v.ValueKind != JsonValueKind.String)
        {
            Error(valuePath, "must be a color string: a palette token name or #RRGGBB.");
            return null;
        }
        var raw = v.GetString()!;
        if (_palette.ContainsKey(raw))
        {
            return raw;
        }
        if (HexColorPattern.IsMatch(raw))
        {
            Warn(valuePath, $"raw hex color '{raw}' is off-palette: prefer a design palette token.");
            return raw;
        }
        if (raw.StartsWith("#", StringComparison.Ordinal))
        {
            Error(valuePath, $"'{raw}' is not a valid hex color: expected #RRGGBB.");
            return null;
        }
        Error(valuePath, $"unknown color '{raw}': not a palette token and not #RRGGBB hex.", Suggest(raw, _palette.Keys));
        return null;
    }

    private string? ColorProp(JsonElement obj, string name, string path, bool required = false)
    {
        if (!TryGet(obj, name, out var v))
        {
            if (required)
            {
                Error(path, $"'{name}' is required.");
            }
            return null;
        }
        return ColorValue(v, $"{path}.{name}");
    }

    private static string FormatNumber(double value) => value.ToString("0.###", CultureInfo.InvariantCulture);
}
