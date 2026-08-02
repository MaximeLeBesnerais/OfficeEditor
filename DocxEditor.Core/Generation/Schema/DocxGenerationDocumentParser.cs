using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocxEditor.Core.Generation.Contracts;
using DocxEditor.Core.Generation.Design;
using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Schema;

/// <summary>
/// Single-pass parser + validator for the DOCX generation JSON vocabulary (§3 of the
/// generation README). Reads with System.Text.Json, builds the <see cref="DocxGenerationDocument"/>
/// model and collects loud, actionable errors (JSON path + suggestion) instead of failing
/// fast, so AI and human authors can fix every problem at once. Unknown properties are
/// rejected at every vocabulary level; typo hints cover common misspellings.
///
/// The parser is stateless: each call runs an isolated validation session, so a shared
/// instance is safe to reuse and results are deterministic.
/// </summary>
public sealed class DocxGenerationDocumentParser : IDocxGenerationParser
{
    /// <summary>The only vocabulary version supported by this parser.</summary>
    public const string SupportedVersion = DocxGenerationDocument.SupportedVersion;

    string IDocxGenerationParser.SupportedVersion => SupportedVersion;

    private static readonly Regex HexColorPattern = new(@"^#[0-9a-fA-F]{6}$", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly IReadOnlySet<string> FontSlots = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "display", "body" };

    private static readonly IReadOnlyDictionary<string, PageOrientation> PageOrientations = new Dictionary<string, PageOrientation>(StringComparer.OrdinalIgnoreCase)
    {
        ["portrait"] = PageOrientation.Portrait, ["landscape"] = PageOrientation.Landscape
    };

    private static readonly IReadOnlyDictionary<string, PageSizeName> PageSizeNames = new Dictionary<string, PageSizeName>(StringComparer.OrdinalIgnoreCase)
    {
        ["a3"] = PageSizeName.A3, ["a4"] = PageSizeName.A4, ["a5"] = PageSizeName.A5,
        ["b4"] = PageSizeName.B4, ["b5"] = PageSizeName.B5,
        ["letter"] = PageSizeName.Letter, ["legal"] = PageSizeName.Legal,
        ["executive"] = PageSizeName.Executive, ["statement"] = PageSizeName.Statement, ["tabloid"] = PageSizeName.Tabloid
    };

    private static readonly IReadOnlyDictionary<string, SectionBreakType> SectionBreakTypes = new Dictionary<string, SectionBreakType>(StringComparer.OrdinalIgnoreCase)
    {
        ["nextPage"] = SectionBreakType.NextPage, ["continuous"] = SectionBreakType.Continuous,
        ["oddPage"] = SectionBreakType.OddPage, ["evenPage"] = SectionBreakType.EvenPage
    };

    private static readonly IReadOnlyDictionary<string, TextAlignment> TextAlignments = new Dictionary<string, TextAlignment>(StringComparer.OrdinalIgnoreCase)
    {
        ["left"] = TextAlignment.Left, ["center"] = TextAlignment.Center,
        ["right"] = TextAlignment.Right, ["justify"] = TextAlignment.Justify
    };

    private static readonly IReadOnlyDictionary<string, ListKind> ListKinds = new Dictionary<string, ListKind>(StringComparer.OrdinalIgnoreCase)
    {
        ["bullet"] = ListKind.Bullet, ["ordered"] = ListKind.Ordered
    };

    private static readonly IReadOnlyDictionary<string, ImageFitMode> ImageFits = new Dictionary<string, ImageFitMode>(StringComparer.OrdinalIgnoreCase)
    {
        ["fill"] = ImageFitMode.Fill, ["contain"] = ImageFitMode.Contain,
        ["crop"] = ImageFitMode.Crop, ["stretch"] = ImageFitMode.Stretch
    };

    private static readonly IReadOnlyDictionary<string, WrapMode> WrapModes = new Dictionary<string, WrapMode>(StringComparer.OrdinalIgnoreCase)
    {
        ["none"] = WrapMode.None, ["square"] = WrapMode.Square, ["tight"] = WrapMode.Tight,
        ["through"] = WrapMode.Through, ["topAndBottom"] = WrapMode.TopAndBottom,
        ["behindText"] = WrapMode.BehindText, ["inFrontOfText"] = WrapMode.InFrontOfText
    };

    private static readonly IReadOnlyDictionary<string, AnchorReference> AnchorReferences = new Dictionary<string, AnchorReference>(StringComparer.OrdinalIgnoreCase)
    {
        ["page"] = AnchorReference.Page, ["margin"] = AnchorReference.Margin, ["column"] = AnchorReference.Column,
        ["paragraph"] = AnchorReference.Paragraph, ["character"] = AnchorReference.Character
    };

    private static readonly IReadOnlyDictionary<string, CalloutTone> CalloutTones = new Dictionary<string, CalloutTone>(StringComparer.OrdinalIgnoreCase)
    {
        ["note"] = CalloutTone.Note, ["tip"] = CalloutTone.Tip,
        ["warning"] = CalloutTone.Warning, ["error"] = CalloutTone.Error
    };

    private static readonly IReadOnlyDictionary<string, LineOrientation> LineOrientations = new Dictionary<string, LineOrientation>(StringComparer.OrdinalIgnoreCase)
    {
        ["horizontal"] = LineOrientation.Horizontal, ["vertical"] = LineOrientation.Vertical
    };

    private static readonly IReadOnlyDictionary<string, TextRole> TextRoles = new Dictionary<string, TextRole>(StringComparer.OrdinalIgnoreCase)
    {
        ["title"] = TextRole.Title, ["subtitle"] = TextRole.Subtitle, ["eyebrow"] = TextRole.Eyebrow,
        ["heading1"] = TextRole.Heading1, ["heading2"] = TextRole.Heading2, ["heading3"] = TextRole.Heading3,
        ["heading4"] = TextRole.Heading4, ["heading5"] = TextRole.Heading5, ["heading6"] = TextRole.Heading6,
        ["body"] = TextRole.Body, ["muted"] = TextRole.Muted, ["label"] = TextRole.Label,
        ["metric"] = TextRole.Metric, ["metricLabel"] = TextRole.MetricLabel,
        ["tableHeader"] = TextRole.TableHeader, ["tableBody"] = TextRole.TableBody,
        ["callout"] = TextRole.Callout, ["footer"] = TextRole.Footer
    };

    private static readonly IReadOnlyDictionary<string, Density> Densities = new Dictionary<string, Density>(StringComparer.OrdinalIgnoreCase)
    {
        ["compact"] = Density.Compact, ["comfortable"] = Density.Comfortable, ["spacious"] = Density.Spacious
    };

    // Known docx-specific confusions, matched case-insensitively after stripping -_/ and spaces.
    private static readonly IReadOnlyDictionary<string, string> DocxIsms = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        ["pagesize"] = "page size is declared per section: use 'pageSetup' ({\"size\":…,\"margins\":…}).",
        ["margin"] = "use 'margins' (an object with top/right/bottom/left).",
        ["column"] = "use 'columns' (an object with count/spacing/separator).",
        ["break"] = "use 'breakType' (nextPage|continuous|oddPage|evenPage).",
        ["templatepath"] = "use 'template'.",
        ["listtype"] = "use 'kind' (bullet|ordered) on a list.",
        ["alignement"] = "use 'alignment'.",
        ["aligment"] = "use 'alignment'."
    };

    private static readonly IReadOnlySet<string> RootProps = Set("version", "metadata", "design", "template", "sections");
    private static readonly IReadOnlySet<string> MetadataProps = Set("title", "author", "subject", "keywords", "description", "language");
    private static readonly IReadOnlySet<string> DesignProps = Set("theme", "palette", "fonts", "typography", "spacing", "shapes", "page", "layout");
    private static readonly IReadOnlySet<string> FontsProps = Set("display", "body");
    private static readonly IReadOnlySet<string> TypographyTokenProps = Set("font", "size", "color", "bold", "italic", "underline", "allCaps");
    private static readonly IReadOnlySet<string> ShapesProps = Set("cornerRadius", "defaultFill", "defaultStroke", "defaultStrokeWidth");
    private static readonly IReadOnlySet<string> PageProps = Set("size", "orientation", "margins", "defaultFont", "defaultTextColor");
    private static readonly IReadOnlySet<string> LayoutProps = Set("density", "minBodySizePt", "maxTableWidthPt");
    private static readonly IReadOnlySet<string> MarginsProps = Set("top", "right", "bottom", "left");
    private static readonly IReadOnlySet<string> SectionProps = Set("pageSetup", "header", "footer", "blocks", "positioned");
    private static readonly IReadOnlySet<string> PageSetupProps = Set("size", "orientation", "margins", "columns", "breakType");
    private static readonly IReadOnlySet<string> ColumnsProps = Set("count", "spacing", "separator");
    private static readonly IReadOnlySet<string> PageSizeCustomProps = Set("width", "height");
    private static readonly IReadOnlySet<string> ParagraphProps = Set("type", "text", "runs", "style", "token", "role", "alignment", "spacing");
    private static readonly IReadOnlySet<string> HeadingProps = Set("type", "level", "text", "runs", "style", "token", "role", "alignment");
    private static readonly IReadOnlySet<string> ListProps = Set("type", "kind", "start", "items", "style");
    private static readonly IReadOnlySet<string> ListItemProps = Set("text", "runs", "token", "role", "alignment", "spacing");
    private static readonly IReadOnlySet<string> TableProps = Set("type", "rows", "widths", "style", "alignment");
    private static readonly IReadOnlySet<string> RowProps = Set("header", "cells");
    private static readonly IReadOnlySet<string> CellProps = Set("text", "runs", "token", "role", "alignment", "fill");
    private static readonly IReadOnlySet<string> FlowImageProps = Set("type", "src", "fit", "crop", "alt", "width", "height", "style");
    private static readonly IReadOnlySet<string> CalloutProps = Set("type", "tone", "text", "runs", "style", "token", "role");
    private static readonly IReadOnlySet<string> PageBreakProps = Set("type");
    private static readonly IReadOnlySet<string> GroupProps = Set("type", "blocks", "style");
    private static readonly IReadOnlySet<string> RunProps = Set("text", "style", "font", "size", "color", "bold", "italic", "underline", "allCaps");
    private static readonly IReadOnlySet<string> SpacingProps = Set("before", "after", "line");
    private static readonly IReadOnlySet<string> CropProps = Set("left", "top", "right", "bottom");
    private static readonly IReadOnlySet<string> StrokeProps = Set("color", "width");
    private static readonly IReadOnlySet<string> WrapDistancesProps = Set("top", "left", "bottom", "right");
    private static readonly IReadOnlySet<string> TextBoxProps = Set("type", "x", "y", "width", "height", "rotation", "zOrder", "anchor", "wrap", "wrapDistances", "alt", "text", "runs", "token", "alignment", "spacing", "fill", "stroke", "cornerRadius");
    private static readonly IReadOnlySet<string> PositionedImageProps = Set("type", "x", "y", "width", "height", "rotation", "zOrder", "anchor", "wrap", "wrapDistances", "alt", "src", "fit", "crop");
    private static readonly IReadOnlySet<string> RectProps = Set("type", "x", "y", "width", "height", "rotation", "zOrder", "anchor", "wrap", "wrapDistances", "alt", "fill", "stroke", "cornerRadius");
    private static readonly IReadOnlySet<string> LineProps = Set("type", "x", "y", "width", "height", "rotation", "zOrder", "anchor", "wrap", "wrapDistances", "alt", "orientation", "stroke");
    private static readonly IReadOnlySet<string> PositionedCalloutProps = Set("type", "x", "y", "width", "height", "rotation", "zOrder", "anchor", "wrap", "wrapDistances", "alt", "tone", "text", "runs", "token", "alignment", "spacing", "fill", "stroke", "cornerRadius");

    /// <summary>
    /// Validates a generation JSON document in a single pass. Never throws for contract
    /// violations — inspect <see cref="DocxGenerationValidationResult.Errors"/> instead.
    /// </summary>
    public DocxGenerationValidationResult Validate(string json) => new ParseSession().Validate(json);

    /// <summary>
    /// Validates and returns the document, throwing <see cref="DocxGenerationValidationException"/>
    /// listing every error when the document is invalid.
    /// </summary>
    public DocxGenerationDocument Parse(string json) => Validate(json).ThrowIfInvalid().Document!;

    private static IReadOnlySet<string> Set(params string[] names) => new HashSet<string>(names, StringComparer.OrdinalIgnoreCase);

    private static string[] FlowBlockTypeNames() => ["paragraph", "heading", "list", "table", "image", "callout", "pageBreak", "group"];

    private static string[] PositionedTypeNames() => ["textBox", "image", "rect", "line", "callout"];

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

    private static string FormatNumber(double value) => value.ToString("0.###", CultureInfo.InvariantCulture);

    /// <summary>
    /// Per-call validation state (errors, warnings, resolved design tokens). Kept out of the
    /// parser so a shared parser instance is stateless, deterministic and reuse-safe.
    /// </summary>
    private sealed class ParseSession
    {
        private readonly List<DocxGenerationIssue> _errors = [];
        private readonly List<DocxGenerationIssue> _warnings = [];
        private DesignTokens? _design;
        private string? _themeName;
        private IReadOnlyDictionary<string, string> _palette = DesignThemeCatalog.Editorial.Palette;

        public DocxGenerationValidationResult Validate(string json)
        {
            ArgumentNullException.ThrowIfNull(json);

            JsonDocument document;
            try
            {
                document = JsonDocument.Parse(json);
            }
            catch (JsonException ex)
            {
                _errors.Add(new DocxGenerationIssue("$", $"malformed JSON: {ex.Message}", null, DocxGenerationIssueSeverity.Error));
                return Result(null);
            }

            using (document)
            {
                var root = document.RootElement;
                if (root.ValueKind != JsonValueKind.Object)
                {
                    _errors.Add(new DocxGenerationIssue("$", "the document root must be a JSON object ({\"version\":\"1.0\",\"design\":…,\"sections\":[…]}).", null, DocxGenerationIssueSeverity.Error));
                    return Result(null);
                }

                var parsed = ParseDocument(root);
                if (_errors.Count == 0)
                {
                    var semantic = DocxGenerationModelValidator.Validate(parsed);
                    AddDistinct(_errors, semantic.Errors);
                    AddDistinct(_warnings, semantic.Warnings);
                }
                return Result(_errors.Count == 0 ? parsed : null);
            }
        }

        private DocxGenerationValidationResult Result(DocxGenerationDocument? document) =>
            new(document, [.. _errors], [.. _warnings]);

        private void Error(string path, string message, string? suggestion = null) =>
            _errors.Add(new DocxGenerationIssue(path, message, suggestion, DocxGenerationIssueSeverity.Error));

        private void Warn(string path, string message) =>
            _warnings.Add(new DocxGenerationIssue(path, message, null, DocxGenerationIssueSeverity.Warning));

        private static void AddDistinct(
            List<DocxGenerationIssue> target,
            IReadOnlyList<DocxGenerationIssue> additions)
        {
            foreach (var issue in additions)
            {
                if (!target.Contains(issue))
                {
                    target.Add(issue);
                }
            }
        }

        private DocxGenerationDocument ParseDocument(JsonElement root)
        {
            CheckUnknownProps(root, "$", "the document root", RootProps);

            string? version = null;
            if (!TryGet(root, "version", out var versionEl))
            {
                Error("$", "'version' is required (\"1.0\").");
            }
            else
            {
                version = StringValue(versionEl, "$.version");
                if (version is not null && version != SupportedVersion)
                {
                    Error("$.version", $"unsupported version '{version}'; expected \"{SupportedVersion}\".");
                }
            }

            DocxMetadata? metadata = null;
            if (TryGet(root, "metadata", out var metadataEl))
            {
                metadata = ParseMetadata(metadataEl, "$.metadata");
            }

            if (TryGet(root, "design", out var designEl))
            {
                // ParseDesign sets _themeName and the effective (theme + document) palette.
                _design = ParseDesign(designEl, "$.design");
            }

            string? template = null;
            if (TryGet(root, "template", out var templateEl))
            {
                template = StringValue(templateEl, "$.template");
                if (template is not null && template.Length == 0)
                {
                    Error("$.template", "must not be empty (omit it to generate from a blank document).");
                }
            }

            var sections = new List<Section>();
            if (!TryGet(root, "sections", out var sectionsEl))
            {
                Error("$", "'sections' is required (at least one section).");
            }
            else if (sectionsEl.ValueKind != JsonValueKind.Array)
            {
                Error("$.sections", "must be an array of section objects.");
            }
            else
            {
                var index = 0;
                foreach (var sectionEl in sectionsEl.EnumerateArray())
                {
                    var section = ParseSection(sectionEl, $"$.sections[{index}]", isFirst: index == 0);
                    if (section is not null)
                    {
                        sections.Add(section);
                    }
                    index++;
                }
                if (index == 0)
                {
                    Error("$.sections", "at least one section is required.");
                }
            }

            return new DocxGenerationDocument
            {
                Version = version ?? SupportedVersion,
                Metadata = metadata,
                Design = _design,
                TemplatePath = template,
                Sections = sections
            };
        }

        private DocxMetadata? ParseMetadata(JsonElement el, string path)
        {
            if (el.ValueKind != JsonValueKind.Object)
            {
                Error(path, "must be an object ({\"title\":…,\"author\":…,\"subject\":…,\"keywords\":…,\"description\":…,\"language\":…}).");
                return null;
            }
            CheckUnknownProps(el, path, "metadata", MetadataProps);
            return new DocxMetadata
            {
                Title = StringProp(el, "title", path),
                Author = StringProp(el, "author", path),
                Subject = StringProp(el, "subject", path),
                Keywords = StringProp(el, "keywords", path),
                Description = StringProp(el, "description", path),
                Language = StringProp(el, "language", path)
            };
        }

        private DesignTokens? ParseDesign(JsonElement el, string path)
        {
            if (el.ValueKind != JsonValueKind.Object)
            {
                Error(path, "must be an object ({\"theme\":…,\"palette\":…,\"fonts\":…,\"typography\":…,\"spacing\":…,\"shapes\":…,\"page\":…,\"layout\":…}).");
                return null;
            }
            CheckUnknownProps(el, path, "design", DesignProps);

            var theme = StringProp(el, "theme", path);
            if (theme is not null && !DesignThemeCatalog.Themes.ContainsKey(theme))
            {
                Error($"{path}.theme", $"unknown theme '{theme}'. Known themes: {string.Join(", ", DesignThemeCatalog.Themes.Keys)}.", Suggest(theme, DesignThemeCatalog.Themes.Keys));
            }
            _themeName = theme;

            var palette = new Dictionary<string, string>(StringComparer.Ordinal);
            if (TryGet(el, "palette", out var paletteEl))
            {
                if (paletteEl.ValueKind != JsonValueKind.Object)
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
            }
            _palette = DesignThemeCatalog.Resolve(_themeName).EffectivePalette(palette);

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

            var typography = new Dictionary<string, TypographyToken>(StringComparer.Ordinal);
            if (TryGet(el, "typography", out var typographyEl))
            {
                if (typographyEl.ValueKind != JsonValueKind.Object)
                {
                    Error($"{path}.typography", "must be an object mapping token names to typography tokens ({\"font\":…,\"size\":…}).");
                }
                else
                {
                    foreach (var token in typographyEl.EnumerateObject())
                    {
                        var tokenPath = $"{path}.typography.{token.Name}";
                        var parsed = ParseTypographyToken(token.Value, tokenPath, fonts);
                        if (parsed is not null)
                        {
                            typography[token.Name] = parsed;
                        }
                    }
                }
            }

            var spacing = new Dictionary<string, double>(StringComparer.Ordinal);
            if (TryGet(el, "spacing", out var spacingEl))
            {
                if (spacingEl.ValueKind != JsonValueKind.Object)
                {
                    Error($"{path}.spacing", "must be an object mapping token names to pt values.");
                }
                else
                {
                    foreach (var token in spacingEl.EnumerateObject())
                    {
                        var value = NumberValue(token.Value, $"{path}.spacing.{token.Name}", min: 0);
                        if (value is not null)
                        {
                            spacing[token.Name] = value.Value;
                        }
                    }
                }
            }

            var shapes = new ShapeDefaults();
            if (TryGet(el, "shapes", out var shapesEl))
            {
                if (shapesEl.ValueKind != JsonValueKind.Object)
                {
                    Error($"{path}.shapes", "must be an object ({\"cornerRadius\":…,\"defaultFill\":…,\"defaultStroke\":…,\"defaultStrokeWidth\":…}).");
                }
                else
                {
                    CheckUnknownProps(shapesEl, $"{path}.shapes", "shapes", ShapesProps);
                    shapes = new ShapeDefaults
                    {
                        CornerRadiusPt = NumberProp(shapesEl, "cornerRadius", $"{path}.shapes", min: 0) ?? 0,
                        DefaultFill = ColorProp(shapesEl, "defaultFill", $"{path}.shapes"),
                        DefaultStrokeColor = ColorProp(shapesEl, "defaultStroke", $"{path}.shapes"),
                        DefaultStrokeWidthPt = NumberProp(shapesEl, "defaultStrokeWidth", $"{path}.shapes", minExclusive: 0) ?? 1
                    };
                }
            }

            var page = new PageDefaults();
            if (TryGet(el, "page", out var pageEl))
            {
                if (pageEl.ValueKind != JsonValueKind.Object)
                {
                    Error($"{path}.page", "must be an object ({\"size\":…,\"orientation\":…,\"margins\":…,\"defaultFont\":…,\"defaultTextColor\":…}).");
                }
                else
                {
                    CheckUnknownProps(pageEl, $"{path}.page", "page defaults", PageProps);
                    var defaultFont = StringProp(pageEl, "defaultFont", $"{path}.page");
                    if (defaultFont is not null)
                    {
                        CheckFontReference(defaultFont, $"{path}.page.defaultFont", fonts);
                    }
                    page = new PageDefaults
                    {
                        PageSize = EnumProp(pageEl, "size", $"{path}.page", PageSizeNames, (PageSizeName?)null, "page size"),
                        Orientation = EnumProp(pageEl, "orientation", $"{path}.page", PageOrientations, (PageOrientation?)null, "orientation") ?? PageOrientation.Portrait,
                        Margins = TryGet(pageEl, "margins", out var marginsEl) ? ParseMargins(marginsEl, $"{path}.page.margins") : null,
                        DefaultFontFamily = defaultFont,
                        DefaultTextColor = ColorProp(pageEl, "defaultTextColor", $"{path}.page")
                    };
                }
            }

            var layout = new LayoutDefaults();
            if (TryGet(el, "layout", out var layoutEl))
            {
                if (layoutEl.ValueKind != JsonValueKind.Object)
                {
                    Error($"{path}.layout", "must be an object ({\"density\":…,\"minBodySizePt\":…,\"maxTableWidthPt\":…}).");
                }
                else
                {
                    CheckUnknownProps(layoutEl, $"{path}.layout", "layout", LayoutProps);
                    layout = new LayoutDefaults
                    {
                        Density = EnumProp(layoutEl, "density", $"{path}.layout", Densities, (Density?)null, "density"),
                        MinBodySizePt = NumberProp(layoutEl, "minBodySizePt", $"{path}.layout", min: 0),
                        MaxTableWidthPt = NumberProp(layoutEl, "maxTableWidthPt", $"{path}.layout", minExclusive: 0)
                    };
                }
            }

            return new DesignTokens
            {
                Theme = theme,
                Palette = palette,
                Fonts = fonts,
                Typography = typography,
                Spacing = spacing,
                Shapes = shapes,
                Page = page,
                Layout = layout
            };
        }

        private TypographyToken? ParseTypographyToken(JsonElement el, string path, FontTokens fonts)
        {
            if (el.ValueKind != JsonValueKind.Object)
            {
                Error(path, "must be an object ({\"font\":…,\"size\":…,\"color\":…,\"bold\":…,\"italic\":…,\"underline\":…,\"allCaps\":…}).");
                return null;
            }
            CheckUnknownProps(el, path, "a typography token", TypographyTokenProps);
            var font = StringProp(el, "font", path);
            if (font is not null)
            {
                CheckFontReference(font, $"{path}.font", fonts);
            }
            return new TypographyToken
            {
                FontFamily = font,
                SizePt = NumberProp(el, "size", path, minExclusive: 0),
                Color = ColorProp(el, "color", path),
                Bold = BoolProp(el, "bold", path),
                Italic = BoolProp(el, "italic", path),
                Underline = BoolProp(el, "underline", path),
                AllCaps = BoolProp(el, "allCaps", path)
            };
        }

        private void CheckFontReference(string value, string path, FontTokens fonts)
        {
            var theme = DesignThemeCatalog.Resolve(_themeName);
            if (string.Equals(value, "display", StringComparison.OrdinalIgnoreCase) &&
                string.IsNullOrWhiteSpace(fonts.Display) && string.IsNullOrWhiteSpace(theme.DisplayFontFamily))
            {
                Warn(path, "font slot 'display' is not defined in 'design.fonts' nor the active theme; emitters fall back to the document default.");
            }
            else if (string.Equals(value, "body", StringComparison.OrdinalIgnoreCase) &&
                     string.IsNullOrWhiteSpace(fonts.Body) && string.IsNullOrWhiteSpace(theme.BodyFontFamily))
            {
                Warn(path, "font slot 'body' is not defined in 'design.fonts' nor the active theme; emitters fall back to the document default.");
            }
        }

        private Section? ParseSection(JsonElement el, string path, bool isFirst)
        {
            if (el.ValueKind != JsonValueKind.Object)
            {
                Error(path, "each section must be a JSON object ({\"pageSetup\":…,\"blocks\":[…]}).");
                return null;
            }
            CheckUnknownProps(el, path, "a section", SectionProps);

            PageSetup? pageSetup = null;
            if (TryGet(el, "pageSetup", out var pageSetupEl))
            {
                pageSetup = ParsePageSetup(pageSetupEl, $"{path}.pageSetup");
            }

            if (isFirst && pageSetup?.BreakType is not null)
            {
                Warn($"{path}.pageSetup.breakType", "the first section has no preceding section; 'breakType' is ignored.");
            }

            var header = ParseFlowList(el, "header", $"{path}.header");
            var footer = ParseFlowList(el, "footer", $"{path}.footer");

            var blocks = new List<FlowBlock>();
            if (!TryGet(el, "blocks", out var blocksEl))
            {
                Error(path, "'blocks' is required (the body flow content).");
            }
            else if (blocksEl.ValueKind != JsonValueKind.Array)
            {
                Error($"{path}.blocks", "must be an array of flow blocks.");
            }
            else
            {
                var index = 0;
                foreach (var blockEl in blocksEl.EnumerateArray())
                {
                    var block = ParseFlowBlock(blockEl, $"{path}.blocks[{index}]");
                    if (block is not null)
                    {
                        blocks.Add(block);
                    }
                    index++;
                }
                if (index == 0)
                {
                    Error($"{path}.blocks", "a section must contain at least one flow block.");
                }
            }

            var positioned = new List<PositionedElement>();
            if (TryGet(el, "positioned", out var positionedEl))
            {
                if (positionedEl.ValueKind != JsonValueKind.Array)
                {
                    Error($"{path}.positioned", "must be an array of positioned elements.");
                }
                else
                {
                    var index = 0;
                    foreach (var posEl in positionedEl.EnumerateArray())
                    {
                        var element = ParsePositionedElement(posEl, $"{path}.positioned[{index}]");
                        if (element is not null)
                        {
                            positioned.Add(element);
                        }
                        index++;
                    }
                }
            }

            var section = new Section
            {
                PageSetup = pageSetup,
                Header = header,
                Footer = footer,
                Blocks = blocks,
                Positioned = positioned
            };

            CheckSectionGeometry(section, path, pageSetup is null);
            return section;
        }

        private IReadOnlyList<FlowBlock> ParseFlowList(JsonElement el, string name, string path)
        {
            if (!TryGet(el, name, out var listEl))
            {
                return [];
            }
            if (listEl.ValueKind != JsonValueKind.Array)
            {
                Error(path, "must be an array of flow blocks.");
                return [];
            }
            var result = new List<FlowBlock>();
            var index = 0;
            foreach (var itemEl in listEl.EnumerateArray())
            {
                var block = ParseFlowBlock(itemEl, $"{path}[{index}]");
                if (block is not null)
                {
                    result.Add(block);
                }
                index++;
            }
            return result;
        }

        private PageSetup? ParsePageSetup(JsonElement el, string path)
        {
            if (el.ValueKind != JsonValueKind.Object)
            {
                Error(path, "must be an object ({\"size\":…,\"orientation\":…,\"margins\":…,\"columns\":…,\"breakType\":…}).");
                return null;
            }
            CheckUnknownProps(el, path, "a page setup", PageSetupProps);

            var pageSize = ParsePageSize(el, path);
            var orientation = EnumProp(el, "orientation", path, PageOrientations, (PageOrientation?)null, "orientation");
            var margins = TryGet(el, "margins", out var marginsEl) ? ParseMargins(marginsEl, $"{path}.margins") : null;
            var columns = TryGet(el, "columns", out var columnsEl) ? ParseColumns(columnsEl, $"{path}.columns") : null;
            var breakType = EnumProp(el, "breakType", path, SectionBreakTypes, (SectionBreakType?)null, "section break type");

            return new PageSetup
            {
                PageSize = pageSize,
                Orientation = orientation,
                Margins = margins,
                Columns = columns,
                BreakType = breakType
            };
        }

        private PageSize? ParsePageSize(JsonElement el, string path)
        {
            if (!TryGet(el, "size", out var sizeEl))
            {
                return null;
            }
            var sizePath = $"{path}.size";
            if (sizeEl.ValueKind == JsonValueKind.String)
            {
                var name = EnumValue(sizeEl, sizePath, PageSizeNames, "page size");
                return name is null ? null : PageSize.Named(name.Value);
            }
            if (sizeEl.ValueKind != JsonValueKind.Object)
            {
                Error(sizePath, "must be a named size string (\"a4\", \"letter\", …) or an object ({\"width\":…,\"height\":…}).");
                return null;
            }
            CheckUnknownProps(sizeEl, sizePath, "a custom page size", PageSizeCustomProps);
            var width = NumberProp(sizeEl, "width", sizePath, minExclusive: 0);
            var height = NumberProp(sizeEl, "height", sizePath, minExclusive: 0);
            if (width is null || height is null)
            {
                return null;
            }
            return new PageSize { WidthPt = width, HeightPt = height };
        }

        private Margins? ParseMargins(JsonElement el, string path)
        {
            if (el.ValueKind != JsonValueKind.Object)
            {
                Error(path, "must be an object ({\"top\":…,\"right\":…,\"bottom\":…,\"left\":…}) in points.");
                return null;
            }
            CheckUnknownProps(el, path, "margins", MarginsProps);
            return new Margins
            {
                TopPt = NumberProp(el, "top", path, min: 0) ?? 0,
                RightPt = NumberProp(el, "right", path, min: 0) ?? 0,
                BottomPt = NumberProp(el, "bottom", path, min: 0) ?? 0,
                LeftPt = NumberProp(el, "left", path, min: 0) ?? 0
            };
        }

        private PageColumns? ParseColumns(JsonElement el, string path)
        {
            if (el.ValueKind != JsonValueKind.Object)
            {
                Error(path, "must be an object ({\"count\":…,\"spacing\":…,\"separator\":…}).");
                return null;
            }
            CheckUnknownProps(el, path, "columns", ColumnsProps);
            var count = IntProp(el, "count", path, min: 1) ?? 1;
            var spacing = NumberProp(el, "spacing", path, min: 0) ?? 0;
            var separator = BoolProp(el, "separator", path);
            if (count == 1 && (spacing > 0 || separator))
            {
                Warn(path, "'spacing'/'separator' have no effect with a single column (count: 1).");
            }
            return new PageColumns { Count = count, SpacingPt = spacing, SeparatorLine = separator };
        }

        private void CheckSectionGeometry(Section section, string sectionPath, bool pageSetupAbsent)
        {
            var pageSetup = section.PageSetup;
            var orientation = pageSetup?.Orientation ?? _design?.Page.Orientation ?? PageOrientation.Portrait;
            var margins = pageSetup?.Margins ?? _design?.Page.Margins ?? Margins.Defaults;
            PageSize pageSize = pageSetup?.PageSize ?? (_design?.Page.PageSize is { } dps ? PageSize.Named(dps) : PageSize.Default);

            var (pageWidth, pageHeight) = pageSize.EffectiveSize(orientation);
            var geometryPath = pageSetupAbsent ? sectionPath : $"{sectionPath}.pageSetup";

            if (margins.TopPt + margins.BottomPt >= pageHeight)
            {
                Error(geometryPath, $"margins exceed the page height: top+bottom = {FormatNumber(margins.TopPt + margins.BottomPt)}pt ≥ {FormatNumber(pageHeight)}pt.");
            }
            if (margins.LeftPt + margins.RightPt >= pageWidth)
            {
                Error(geometryPath, $"margins exceed the page width: left+right = {FormatNumber(margins.LeftPt + margins.RightPt)}pt ≥ {FormatNumber(pageWidth)}pt.");
            }

            if (pageSetup?.Columns is { } columns)
            {
                var textWidth = pageWidth - margins.LeftPt - margins.RightPt;
                if (textWidth > 0)
                {
                    var gutters = columns.SpacingPt * (columns.Count - 1);
                    if (gutters >= textWidth)
                    {
                        Error(geometryPath, $"columns do not fit the text area: {columns.Count} columns × {FormatNumber(columns.SpacingPt)}pt gutters need {FormatNumber(gutters)}pt ≥ {FormatNumber(textWidth)}pt.");
                    }
                }
            }

            for (var i = 0; i < section.Positioned.Count; i++)
            {
                var position = section.Positioned[i].Position;
                var posPath = $"{sectionPath}.positioned[{i}]";
                var rightEdge = position.X + (position.WidthPt ?? 0);
                if (rightEdge > pageWidth)
                {
                    Warn(posPath, $"extends past the right page edge ({FormatNumber(rightEdge)}pt > {FormatNumber(pageWidth)}pt).");
                }
                var bottomEdge = position.Y + (position.HeightPt ?? 0);
                if (bottomEdge > pageHeight)
                {
                    Warn(posPath, $"extends past the bottom page edge ({FormatNumber(bottomEdge)}pt > {FormatNumber(pageHeight)}pt).");
                }
            }
        }

        private FlowBlock? ParseFlowBlock(JsonElement el, string path)
        {
            if (el.ValueKind != JsonValueKind.Object)
            {
                Error(path, "each flow block must be a JSON object with a 'type'.");
                return null;
            }
            if (!TryGet(el, "type", out var typeEl) || typeEl.ValueKind != JsonValueKind.String)
            {
                Error(path, "'type' is required (paragraph|heading|list|table|image|callout|pageBreak|group).");
                return null;
            }
            var type = typeEl.GetString()!;
            switch (type.ToLowerInvariant())
            {
                case "paragraph":
                    return ParseParagraph(el, path);
                case "heading":
                    return ParseHeading(el, path);
                case "list":
                    return ParseList(el, path);
                case "table":
                    return ParseTable(el, path);
                case "image":
                    return ParseFlowImage(el, path);
                case "callout":
                    return ParseFlowCallout(el, path);
                case "pagebreak":
                    return ParsePageBreak(el, path);
                case "group":
                    return ParseFlowGroup(el, path);
                default:
                    Error(path, $"unknown flow block type '{type}'. Expected paragraph|heading|list|table|image|callout|pageBreak|group.", Suggest(type, FlowBlockTypeNames()));
                    return null;
            }
        }

        private ParagraphBlock? ParseParagraph(JsonElement el, string path)
        {
            CheckUnknownProps(el, path, "a paragraph", ParagraphProps);
            var content = ParseTextModel(el, path);
            if (content is null)
            {
                return null;
            }
            return new ParagraphBlock { Style = StringProp(el, "style", path), Content = content };
        }

        private HeadingBlock? ParseHeading(JsonElement el, string path)
        {
            CheckUnknownProps(el, path, "a heading", HeadingProps);
            var level = IntProp(el, "level", path, min: 1, max: 6) ?? 1;
            var content = ParseTextModel(el, path);
            if (content is null)
            {
                return null;
            }
            return new HeadingBlock { Level = level, Style = StringProp(el, "style", path), Content = content };
        }

        private ListBlock? ParseList(JsonElement el, string path)
        {
            CheckUnknownProps(el, path, "a list", ListProps);
            var kind = EnumProp(el, "kind", path, ListKinds, ListKind.Bullet, "list kind");
            var start = IntProp(el, "start", path, min: 1);
            if (start is not null && kind == ListKind.Bullet)
            {
                Warn($"{path}.start", "'start' is ignored for bullet lists; use kind 'ordered' for numbered lists.");
            }

            var items = new List<TextModel>();
            if (!TryGet(el, "items", out var itemsEl))
            {
                Error(path, "'items' is required (an array of strings or text objects).");
            }
            else if (itemsEl.ValueKind != JsonValueKind.Array)
            {
                Error($"{path}.items", "must be an array of strings or text objects.");
            }
            else
            {
                var index = 0;
                foreach (var itemEl in itemsEl.EnumerateArray())
                {
                    var item = ParseListItem(itemEl, $"{path}.items[{index}]");
                    if (item is not null)
                    {
                        items.Add(item);
                    }
                    index++;
                }
                if (index == 0)
                {
                    Error($"{path}.items", "a list must contain at least one item.");
                }
            }

            return new ListBlock { Kind = kind, StartIndex = start, Style = StringProp(el, "style", path), Items = items };
        }

        private TextModel? ParseListItem(JsonElement el, string path)
        {
            if (el.ValueKind == JsonValueKind.String)
            {
                return new TextModel { Text = el.GetString()! };
            }
            if (el.ValueKind != JsonValueKind.Object)
            {
                Error(path, "each list item must be a string or a text object ({\"text\":…}).");
                return null;
            }
            CheckUnknownProps(el, path, "a list item", ListItemProps);
            return ParseTextModel(el, path);
        }

        private TableBlock? ParseTable(JsonElement el, string path)
        {
            CheckUnknownProps(el, path, "a table", TableProps);

            var rows = new List<TableRow>();
            int? columnCount = null;
            if (!TryGet(el, "rows", out var rowsEl))
            {
                Error(path, "'rows' is required (an array of row objects).");
            }
            else if (rowsEl.ValueKind != JsonValueKind.Array)
            {
                Error($"{path}.rows", "must be an array of row objects ({\"header\":…,\"cells\":[…]}).");
            }
            else
            {
                var index = 0;
                foreach (var rowEl in rowsEl.EnumerateArray())
                {
                    var rowPath = $"{path}.rows[{index}]";
                    var row = ParseTableRow(rowEl, rowPath);
                    if (row is not null)
                    {
                        if (columnCount is null)
                        {
                            columnCount = row.Cells.Count;
                        }
                        else if (row.Cells.Count != columnCount.Value)
                        {
                            Error(rowPath, $"row has {row.Cells.Count} cells but the table has {columnCount.Value} columns; all rows must have the same number of cells.");
                        }
                        rows.Add(row);
                    }
                    index++;
                }
                if (index == 0)
                {
                    Error($"{path}.rows", "a table must contain at least one row.");
                }
            }

            var widths = new List<double>();
            if (TryGet(el, "widths", out var widthsEl))
            {
                if (widthsEl.ValueKind != JsonValueKind.Array)
                {
                    Error($"{path}.widths", "must be an array of pt widths (one per column).");
                }
                else
                {
                    var index = 0;
                    foreach (var wEl in widthsEl.EnumerateArray())
                    {
                        var width = NumberValue(wEl, $"{path}.widths[{index}]", minExclusive: 0);
                        if (width is not null)
                        {
                            widths.Add(width.Value);
                        }
                        index++;
                    }
                }
            }
            if (widths.Count > 0 && columnCount is not null && widths.Count != columnCount.Value)
            {
                Error($"{path}.widths", $"got {widths.Count} column widths but the table has {columnCount.Value} columns.");
            }

            return new TableBlock
            {
                Rows = rows,
                ColumnWidthsPt = widths.Count > 0 ? widths : null,
                Alignment = EnumProp(el, "alignment", path, TextAlignments, (TextAlignment?)null, "alignment"),
                Style = StringProp(el, "style", path)
            };
        }

        private TableRow? ParseTableRow(JsonElement el, string path)
        {
            if (el.ValueKind != JsonValueKind.Object)
            {
                Error(path, "each row must be an object ({\"header\":…,\"cells\":[…]}).");
                return null;
            }
            CheckUnknownProps(el, path, "a table row", RowProps);
            var cells = new List<TableCell>();
            if (!TryGet(el, "cells", out var cellsEl))
            {
                Error(path, "'cells' is required (an array of cell objects).");
            }
            else if (cellsEl.ValueKind != JsonValueKind.Array)
            {
                Error($"{path}.cells", "must be an array of cell objects.");
            }
            else
            {
                var index = 0;
                foreach (var cellEl in cellsEl.EnumerateArray())
                {
                    var cell = ParseTableCell(cellEl, $"{path}.cells[{index}]");
                    if (cell is not null)
                    {
                        cells.Add(cell);
                    }
                    index++;
                }
                if (index == 0)
                {
                    Error($"{path}.cells", "a row must contain at least one cell.");
                }
            }
            return new TableRow { Cells = cells, IsHeader = BoolProp(el, "header", path) };
        }

        private TableCell? ParseTableCell(JsonElement el, string path)
        {
            if (el.ValueKind != JsonValueKind.Object)
            {
                Error(path, "each cell must be an object ({\"text\":…,\"fill\":…}).");
                return null;
            }
            CheckUnknownProps(el, path, "a table cell", CellProps);
            var content = ParseTextModel(el, path, allowEmptyContent: true);
            return new TableCell
            {
                Content = content,
                Fill = ColorProp(el, "fill", path),
                Alignment = EnumProp(el, "alignment", path, TextAlignments, (TextAlignment?)null, "alignment")
            };
        }

        private ImageElement? ParseFlowImage(JsonElement el, string path)
        {
            CheckUnknownProps(el, path, "an inline image", FlowImageProps);
            var src = StringProp(el, "src", path, required: true, allowEmpty: false);
            return new ImageElement
            {
                Source = src ?? string.Empty,
                Fit = EnumProp(el, "fit", path, ImageFits, ImageFitMode.Fill, "image fit mode"),
                Crop = ParseCrop(el, path),
                Alt = StringProp(el, "alt", path),
                WidthPt = NumberProp(el, "width", path, minExclusive: 0),
                HeightPt = NumberProp(el, "height", path, minExclusive: 0),
                Style = StringProp(el, "style", path)
            };
        }

        private CalloutBlock? ParseFlowCallout(JsonElement el, string path)
        {
            CheckUnknownProps(el, path, "a callout", CalloutProps);
            var content = ParseTextModel(el, path);
            if (content is null)
            {
                return null;
            }
            return new CalloutBlock
            {
                Tone = EnumProp(el, "tone", path, CalloutTones, CalloutTone.Note, "callout tone"),
                Content = content,
                Style = StringProp(el, "style", path)
            };
        }

        private PageBreakBlock? ParsePageBreak(JsonElement el, string path)
        {
            CheckUnknownProps(el, path, "a page break", PageBreakProps);
            return new PageBreakBlock();
        }

        private FlowContainerBlock? ParseFlowGroup(JsonElement el, string path)
        {
            CheckUnknownProps(el, path, "a group", GroupProps);
            var blocks = new List<FlowBlock>();
            if (!TryGet(el, "blocks", out var blocksEl))
            {
                Error(path, "'blocks' is required (an array of flow blocks).");
            }
            else if (blocksEl.ValueKind != JsonValueKind.Array)
            {
                Error($"{path}.blocks", "must be an array of flow blocks.");
            }
            else
            {
                var index = 0;
                foreach (var blockEl in blocksEl.EnumerateArray())
                {
                    var block = ParseFlowBlock(blockEl, $"{path}.blocks[{index}]");
                    if (block is not null)
                    {
                        blocks.Add(block);
                    }
                    index++;
                }
                if (index == 0)
                {
                    Error($"{path}.blocks", "a group must contain at least one flow block.");
                }
            }
            return new FlowContainerBlock { Style = StringProp(el, "style", path), Blocks = blocks };
        }

        private TextModel? ParseTextModel(JsonElement el, string path, bool allowEmptyContent = false)
        {
            var hasText = TryGet(el, "text", out var textEl);
            var hasRuns = TryGet(el, "runs", out var runsEl);
            if (hasText == hasRuns)
            {
                if (allowEmptyContent && !hasText)
                {
                    return null;
                }
                Error(path, "text content needs exactly one of 'text' (string) or 'runs' (array).");
                return null;
            }

            string? text = null;
            IReadOnlyList<Run>? runs = null;
            if (hasText)
            {
                if (textEl.ValueKind != JsonValueKind.String)
                {
                    Error($"{path}.text", "must be a string.");
                    return null;
                }
                text = textEl.GetString()!;
            }
            else if (hasRuns)
            {
                runs = ParseRuns(runsEl, $"{path}.runs");
                if (runs is null)
                {
                    return null;
                }
            }

            var token = StringProp(el, "token", path);
            if (token is not null && (_design is null || !_design.Typography.ContainsKey(token)))
            {
                Error($"{path}.token", $"unknown typography token '{token}'.", Suggest(token, TypographyTokenNames()));
            }

            var role = EnumProp(el, "role", path, TextRoles, (TextRole?)null, "text role");
            var alignment = EnumProp(el, "alignment", path, TextAlignments, (TextAlignment?)null, "alignment");
            var spacing = TryGet(el, "spacing", out var spacingEl) ? ParseParagraphSpacing(spacingEl, $"{path}.spacing") : null;

            return new TextModel { Text = text, Runs = runs, Token = token, Role = role, Alignment = alignment, Spacing = spacing };
        }

        private IReadOnlyList<Run>? ParseRuns(JsonElement el, string path)
        {
            if (el.ValueKind != JsonValueKind.Array)
            {
                Error(path, "must be an array of run objects ({\"text\":…}).");
                return null;
            }
            var runs = new List<Run>();
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
                    CheckUnknownProps(runEl, runPath, "a run", RunProps);
                    var text = StringProp(runEl, "text", runPath, required: true, allowEmpty: false);
                    var font = StringProp(runEl, "font", runPath);
                    if (font is not null)
                    {
                        CheckFontReference(font, $"{runPath}.font", _design?.Fonts ?? new FontTokens());
                    }
                    runs.Add(new Run
                    {
                        Text = text ?? string.Empty,
                        Style = StringProp(runEl, "style", runPath),
                        FontFamily = font,
                        FontSizePt = NumberProp(runEl, "size", runPath, minExclusive: 0),
                        Color = ColorProp(runEl, "color", runPath),
                        Bold = BoolProp(runEl, "bold", runPath),
                        Italic = BoolProp(runEl, "italic", runPath),
                        Underline = BoolProp(runEl, "underline", runPath),
                        AllCaps = BoolProp(runEl, "allCaps", runPath)
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

        private ParagraphSpacing? ParseParagraphSpacing(JsonElement el, string path)
        {
            if (el.ValueKind != JsonValueKind.Object)
            {
                Error(path, "must be an object ({\"before\":…,\"after\":…,\"line\":…}).");
                return null;
            }
            CheckUnknownProps(el, path, "paragraph spacing", SpacingProps);
            return new ParagraphSpacing
            {
                BeforePt = SpacingValue(el, "before", path),
                AfterPt = SpacingValue(el, "after", path),
                LineMultiple = NumberProp(el, "line", path, minExclusive: 0)
            };
        }

        private double? SpacingValue(JsonElement obj, string name, string path)
        {
            if (!TryGet(obj, name, out var v))
            {
                return null;
            }
            var valuePath = $"{path}.{name}";
            if (v.ValueKind == JsonValueKind.Number)
            {
                return NumberValue(v, valuePath, min: 0);
            }
            if (v.ValueKind != JsonValueKind.String)
            {
                Error(valuePath, "must be a pt number or a 'design.spacing' token name.");
                return null;
            }
            var token = v.GetString()!;
            if (_design is not null && _design.Spacing.TryGetValue(token, out var value))
            {
                return value;
            }
            Error(valuePath, $"unknown spacing token '{token}'.", Suggest(token, SpacingTokenNames()));
            return null;
        }

        private ImageCrop? ParseCrop(JsonElement el, string path)
        {
            if (!TryGet(el, "crop", out var cropEl))
            {
                return null;
            }
            var cropPath = $"{path}.crop";
            if (cropEl.ValueKind != JsonValueKind.Object)
            {
                Error(cropPath, "must be an object ({\"left\":…,\"top\":…,\"right\":…,\"bottom\":…}) with 0..1 fractions of each edge.");
                return null;
            }
            CheckUnknownProps(cropEl, cropPath, "a crop rectangle", CropProps);
            var left = NumberProp(cropEl, "left", cropPath, min: 0, max: 1) ?? 0;
            var top = NumberProp(cropEl, "top", cropPath, min: 0, max: 1) ?? 0;
            var right = NumberProp(cropEl, "right", cropPath, min: 0, max: 1) ?? 0;
            var bottom = NumberProp(cropEl, "bottom", cropPath, min: 0, max: 1) ?? 0;
            if (left + right >= 1)
            {
                Error(cropPath, "left + right crop ≥ 1 removes the whole image horizontally.");
            }
            if (top + bottom >= 1)
            {
                Error(cropPath, "top + bottom crop ≥ 1 removes the whole image vertically.");
            }
            return new ImageCrop { Left = left, Top = top, Right = right, Bottom = bottom };
        }

        private PositionedElement? ParsePositionedElement(JsonElement el, string path)
        {
            if (el.ValueKind != JsonValueKind.Object)
            {
                Error(path, "each positioned element must be a JSON object with a 'type'.");
                return null;
            }
            if (!TryGet(el, "type", out var typeEl) || typeEl.ValueKind != JsonValueKind.String)
            {
                Error(path, "'type' is required (textBox|image|rect|line|callout).");
                return null;
            }
            var type = typeEl.GetString()!;
            switch (type.ToLowerInvariant())
            {
                case "textbox":
                    return ParsePositionedTextBox(el, path);
                case "image":
                    return ParsePositionedImage(el, path);
                case "rect":
                    return ParsePositionedRectangle(el, path);
                case "line":
                    return ParsePositionedLine(el, path);
                case "callout":
                    return ParsePositionedCallout(el, path);
                default:
                    Error(path, $"unknown positioned element type '{type}'. Expected textBox|image|rect|line|callout.", Suggest(type, PositionedTypeNames()));
                    return null;
            }
        }

        private PositionSpec? ParsePositionSpec(JsonElement el, string path, bool widthRequired, bool heightRequired, bool heightForbidden)
        {
            var x = NumberProp(el, "x", path, required: true, min: 0);
            var y = NumberProp(el, "y", path, required: true, min: 0);
            var width = NumberProp(el, "width", path, required: widthRequired, minExclusive: 0);
            var height = NumberProp(el, "height", path, required: heightRequired, minExclusive: 0);
            if (heightForbidden && TryGet(el, "height", out _))
            {
                Error($"{path}.height", "a line has no height; use 'width' as its length and 'orientation' for its axis.");
            }

            var rotation = NumberProp(el, "rotation", path) ?? 0;
            var zOrder = IntProp(el, "zOrder", path) ?? 0;
            var anchor = EnumProp(el, "anchor", path, AnchorReferences, AnchorReference.Margin, "anchor reference");
            var wrap = EnumProp(el, "wrap", path, WrapModes, WrapMode.Square, "wrap mode");
            var wrapDistances = ParseWrapDistances(el, path);
            if (wrapDistances is not null && wrap is WrapMode.None or WrapMode.BehindText or WrapMode.InFrontOfText)
            {
                Warn($"{path}.wrapDistances", "'wrapDistances' only applies to square/tight/through wraps (or topAndBottom).");
            }

            if (x is null || y is null)
            {
                return null;
            }
            return new PositionSpec
            {
                X = x.Value,
                Y = y.Value,
                WidthPt = width,
                HeightPt = height,
                Rotation = rotation,
                ZOrder = zOrder,
                Anchor = anchor,
                Wrap = wrap,
                WrapDistances = wrapDistances,
                Alt = StringProp(el, "alt", path)
            };
        }

        private WrapDistances? ParseWrapDistances(JsonElement el, string path)
        {
            if (!TryGet(el, "wrapDistances", out var wEl))
            {
                return null;
            }
            var wPath = $"{path}.wrapDistances";
            if (wEl.ValueKind != JsonValueKind.Object)
            {
                Error(wPath, "must be an object ({\"top\":…,\"left\":…,\"bottom\":…,\"right\":…}) in points.");
                return null;
            }
            CheckUnknownProps(wEl, wPath, "wrap distances", WrapDistancesProps);
            return new WrapDistances
            {
                TopPt = NumberProp(wEl, "top", wPath, min: 0) ?? 0,
                LeftPt = NumberProp(wEl, "left", wPath, min: 0) ?? 0,
                BottomPt = NumberProp(wEl, "bottom", wPath, min: 0) ?? 0,
                RightPt = NumberProp(wEl, "right", wPath, min: 0) ?? 0
            };
        }

        private StrokeSpec? ParseStroke(JsonElement el, string path)
        {
            if (!TryGet(el, "stroke", out var strokeEl))
            {
                return null;
            }
            var strokePath = $"{path}.stroke";
            if (strokeEl.ValueKind == JsonValueKind.String)
            {
                var color = ColorValue(strokeEl, strokePath);
                return color is null ? null : new StrokeSpec { Color = color };
            }
            if (strokeEl.ValueKind != JsonValueKind.Object)
            {
                Error(strokePath, "must be a color string or an object ({\"color\":…,\"width\":…}).");
                return null;
            }
            CheckUnknownProps(strokeEl, strokePath, "a stroke", StrokeProps);
            var strokeColor = ColorProp(strokeEl, "color", strokePath, required: true);
            if (strokeColor is null)
            {
                return null;
            }
            return new StrokeSpec { Color = strokeColor, WidthPt = NumberProp(strokeEl, "width", strokePath, minExclusive: 0) ?? 1 };
        }

        private PositionedTextBox? ParsePositionedTextBox(JsonElement el, string path)
        {
            CheckUnknownProps(el, path, "a text box", TextBoxProps);
            var position = ParsePositionSpec(el, path, widthRequired: true, heightRequired: true, heightForbidden: false);
            var content = ParseTextModel(el, path);
            if (position is null || content is null)
            {
                return null;
            }
            return new PositionedTextBox
            {
                Position = position,
                Content = content,
                Fill = ColorProp(el, "fill", path),
                Stroke = ParseStroke(el, path),
                CornerRadiusPt = NumberProp(el, "cornerRadius", path, min: 0)
            };
        }

        private PositionedImage? ParsePositionedImage(JsonElement el, string path)
        {
            CheckUnknownProps(el, path, "a floating image", PositionedImageProps);
            var position = ParsePositionSpec(el, path, widthRequired: true, heightRequired: true, heightForbidden: false);
            var src = StringProp(el, "src", path, required: true, allowEmpty: false);
            if (position is null)
            {
                return null;
            }
            return new PositionedImage
            {
                Position = position,
                Source = src ?? string.Empty,
                Fit = EnumProp(el, "fit", path, ImageFits, ImageFitMode.Fill, "image fit mode"),
                Crop = ParseCrop(el, path)
            };
        }

        private PositionedRectangle? ParsePositionedRectangle(JsonElement el, string path)
        {
            CheckUnknownProps(el, path, "a rectangle", RectProps);
            var position = ParsePositionSpec(el, path, widthRequired: true, heightRequired: true, heightForbidden: false);
            if (position is null)
            {
                return null;
            }
            return new PositionedRectangle
            {
                Position = position,
                Fill = ColorProp(el, "fill", path),
                Stroke = ParseStroke(el, path),
                CornerRadiusPt = NumberProp(el, "cornerRadius", path, min: 0)
            };
        }

        private PositionedLine? ParsePositionedLine(JsonElement el, string path)
        {
            CheckUnknownProps(el, path, "a line", LineProps);
            var position = ParsePositionSpec(el, path, widthRequired: true, heightRequired: false, heightForbidden: true);
            if (position is null)
            {
                return null;
            }
            return new PositionedLine
            {
                Position = position,
                Orientation = EnumProp(el, "orientation", path, LineOrientations, LineOrientation.Horizontal, "line orientation"),
                Stroke = ParseStroke(el, path)
            };
        }

        private PositionedCallout? ParsePositionedCallout(JsonElement el, string path)
        {
            CheckUnknownProps(el, path, "a callout", PositionedCalloutProps);
            var position = ParsePositionSpec(el, path, widthRequired: true, heightRequired: true, heightForbidden: false);
            var content = ParseTextModel(el, path);
            if (position is null || content is null)
            {
                return null;
            }
            return new PositionedCallout
            {
                Position = position,
                Tone = EnumProp(el, "tone", path, CalloutTones, CalloutTone.Note, "callout tone"),
                Content = content,
                Fill = ColorProp(el, "fill", path),
                Stroke = ParseStroke(el, path),
                CornerRadiusPt = NumberProp(el, "cornerRadius", path, min: 0)
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
                if (DocxIsms.TryGetValue(NormalizeKey(property.Name), out var hint))
                {
                    Error(propertyPath, hint);
                    continue;
                }
                Error(propertyPath, $"unknown property '{property.Name}' on {subject}.", Suggest(property.Name, allowed));
            }
        }

        private static string NormalizeKey(string name)
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

        private IEnumerable<string> TypographyTokenNames() => _design?.Typography.Keys ?? Enumerable.Empty<string>();

        private IEnumerable<string> SpacingTokenNames() => _design?.Spacing.Keys ?? Enumerable.Empty<string>();

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
            if (v.ValueKind != JsonValueKind.Number || !v.TryGetDouble(out var number))
            {
                Error(valuePath, "must be a JSON number (pt).");
                return null;
            }
            if (!double.IsFinite(number))
            {
                Error(valuePath, "must be a finite number.");
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
    }
}
