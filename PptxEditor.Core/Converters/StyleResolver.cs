using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace PptxEditor.Core.Converters;

/// <summary>
/// Resolves OOXML style cascade: slide -> layout -> master -> theme
/// </summary>
public sealed class StyleResolver
{
    private readonly PresentationDocument _document;
    private readonly SlidePart _slidePart;
    private readonly SlideLayoutPart? _layoutPart;
    private readonly SlideMasterPart? _masterPart;
    private readonly ThemePart? _themePart;
    private readonly Dictionary<string, string> _schemeColors;
    private readonly TextStyleCache _textStyles;

    public StyleResolver(PresentationDocument document, SlidePart slidePart)
    {
        _document = document;
        _slidePart = slidePart;
        _layoutPart = GetSlideLayoutPart();
        _masterPart = _layoutPart?.SlideMasterPart;
        _themePart = _masterPart?.ThemePart;
        _schemeColors = LoadSchemeColors();
        _textStyles = LoadTextStyles();
    }

    /// <summary>
    /// Resolves background color using cascade: slide -> layout -> master
    /// </summary>
    public string? ResolveBackgroundColor()
    {
        // Check slide first
        var slideBg = GetSlideBackground();
        if (!string.IsNullOrEmpty(slideBg))
            return slideBg;

        // Check layout
        var layoutBg = GetLayoutBackground();
        if (!string.IsNullOrEmpty(layoutBg))
            return layoutBg;

        // Check master
        var masterBg = GetMasterBackground();
        if (!string.IsNullOrEmpty(masterBg))
            return masterBg;

        return null;
    }

    /// <summary>
    /// Gets default text formatting for a placeholder type
    /// </summary>
    public DefaultTextStyle GetDefaultTextStyle(PlaceholderValues? placeholderType)
    {
        // Shapes without placeholders should NOT inherit body style defaults
        if (placeholderType == null)
            return new DefaultTextStyle();

        string key;
        if (placeholderType == PlaceholderValues.Title || placeholderType == PlaceholderValues.CenteredTitle)
            key = "Title";
        else if (placeholderType == PlaceholderValues.Body)
            key = "Body";
        else if (placeholderType == PlaceholderValues.SubTitle)
            key = "Body";
        else if (placeholderType == PlaceholderValues.Object)
            key = "Other";
        else if (placeholderType == PlaceholderValues.Footer
            || placeholderType == PlaceholderValues.SlideNumber
            || placeholderType == PlaceholderValues.DateAndTime
            || placeholderType == PlaceholderValues.Header)
            // Footer/date/slide-number/header placeholders take their defaults from the
            // master's "other" text style when the master carries no matching
            // placeholder of its own (PowerPoint's fallback for aux placeholders).
            key = "Other";
        else
            return new DefaultTextStyle();

        return _textStyles.GetStyle(key);
    }

    /// <summary>
    /// Theme color scheme map (scheme name -> #RRGGBB) loaded from the theme of the
    /// slide's master, including sysClr LastColor fallbacks.
    /// Keys (when present in the theme): dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink.
    /// Exposes raw scheme slots — clrMap remapping is NOT applied here.
    /// </summary>
    public IReadOnlyDictionary<string, string> SchemeColors => _schemeColors;

    /// <summary>
    /// Gets the master text style for an exact style key ("Title", "Body", "Other") and
    /// zero-based level, or null when the master defines no such entry.
    /// Unlike <see cref="GetDefaultTextStyle"/>, no level-0 or Body fallback is applied.
    /// </summary>
    public DefaultTextStyle? GetTextStyle(string key, int level)
        => _textStyles.TryGetStyle(key, level);

    /// <summary>
    /// Resolves a scheme color to RGB
    /// </summary>
    public string? ResolveSchemeColor(string schemeColorName)
    {
        schemeColorName = schemeColorName switch
        {
            "bg1" => "lt1",
            "tx1" => "dk1",
            "bg2" => "lt2",
            "tx2" => "dk2",
            _ => schemeColorName
        };

        if (_schemeColors.TryGetValue(schemeColorName, out var rgb))
            return rgb;
        return null;
    }

    #region Background Resolution

    private string? GetSlideBackground()
    {
        var bg = _slidePart.Slide?.CommonSlideData?.Background;
        if (bg?.BackgroundProperties != null)
        {
            return ExtractBackgroundColor(bg.BackgroundProperties);
        }
        if (bg?.BackgroundStyleReference != null)
        {
            return ExtractBackgroundReferenceColor(bg.BackgroundStyleReference);
        }
        return null;
    }

    private string? GetLayoutBackground()
    {
        var bg = _layoutPart?.SlideLayout?.CommonSlideData?.Background;
        if (bg?.BackgroundProperties != null)
        {
            return ExtractBackgroundColor(bg.BackgroundProperties);
        }
        if (bg?.BackgroundStyleReference != null)
        {
            return ExtractBackgroundReferenceColor(bg.BackgroundStyleReference);
        }
        return null;
    }

    private string? GetMasterBackground()
    {
        var bg = _masterPart?.SlideMaster?.CommonSlideData?.Background;
        if (bg?.BackgroundProperties != null)
        {
            return ExtractBackgroundColor(bg.BackgroundProperties);
        }
        if (bg?.BackgroundStyleReference != null)
        {
            return ExtractBackgroundReferenceColor(bg.BackgroundStyleReference);
        }
        return null;
    }

    /// <summary>
    /// Resolves the color of a &lt;p:bgRef&gt; background reference. The <c>idx</c> attribute
    /// points into the theme's <c>bgFillStyleLst</c> (1001-based); honoring the themed fill
    /// style is out of scope — the referenced color is resolved directly, which is exact for
    /// the common <c>idx="1001"</c> case (a solid phClr fill style) and a close approximation
    /// otherwise.
    /// </summary>
    private string? ExtractBackgroundReferenceColor(BackgroundStyleReference bgRef)
    {
        var rgb = bgRef.Elements<Drawing.RgbColorModelHex>().FirstOrDefault();
        if (rgb?.Val != null)
            return $"#{rgb.Val.Value}";

        var schemeClr = bgRef.Elements<Drawing.SchemeColor>().FirstOrDefault();
        if (schemeClr == null)
            return null;

        // Raw XML attribute read per AGENTS.pptx.md rule 1 — SDK enum parsing is unreliable.
        var schemeName = GetAttributeValue(schemeClr, "val");
        return string.IsNullOrEmpty(schemeName) ? null : ResolveSchemeColor(schemeName);
    }

    private string? ExtractBackgroundColor(BackgroundProperties bgProps)
    {
        var solidFill = bgProps.Elements<Drawing.SolidFill>().FirstOrDefault();
        if (solidFill == null)
            return null;

        // Check for RGB color
        var rgb = solidFill.RgbColorModelHex;
        if (rgb?.Val != null)
            return $"#{rgb.Val.Value}";

        // Check for scheme color
        var schemeClr = solidFill.SchemeColor;
        if (schemeClr != null)
        {
            // Try to get the value from the enum property
            if (schemeClr.Val != null && schemeClr.Val.HasValue)
            {
                var valStr = schemeClr.Val.Value.ToString();
                // Check if it's not empty/default
                if (!string.IsNullOrEmpty(valStr) && valStr != "None")
                {
                    var resolved = ResolveSchemeColor(valStr);
                    if (!string.IsNullOrEmpty(resolved))
                        return resolved;
                }
            }
            
            // Fallback: parse from OuterXml since SDK enum parsing is unreliable
            var outerXml = schemeClr.OuterXml;
            if (!string.IsNullOrEmpty(outerXml))
            {
                var match = System.Text.RegularExpressions.Regex.Match(outerXml, @"val\s*=\s*""([^""]*)""");
                if (match.Success)
                {
                    var resolved = ResolveSchemeColor(match.Groups[1].Value);
                    if (!string.IsNullOrEmpty(resolved))
                        return resolved;
                }
            }
        }

        return null;
    }

    #endregion

    #region Text Style Resolution

    private TextStyleCache LoadTextStyles()
    {
        var cache = new TextStyleCache();

        if (_masterPart?.SlideMaster?.TextStyles == null)
            return cache;

        var txStyles = _masterPart.SlideMaster.TextStyles;

        // Title style - txStyles.TitleStyle is an OpenXmlElement
        LoadLevelStyles(cache, "Title", txStyles.TitleStyle);

        // Body style
        LoadLevelStyles(cache, "Body", txStyles.BodyStyle);

        // Other style
        LoadLevelStyles(cache, "Other", txStyles.OtherStyle);

        return cache;
    }

    private void LoadLevelStyles(TextStyleCache cache, string styleType, OpenXmlElement? styleList)
    {
        if (styleList == null)
            return;

        // Iterate through level paragraph properties
        foreach (var child in styleList.ChildElements)
        {
            var localName = child.LocalName;
            if (!localName.StartsWith("lvl") || !localName.EndsWith("pPr"))
                continue;

            var level = ParseLevel(localName);
            if (level == null)
                continue;

            // Find default run properties within paragraph properties
            var defRPr = child.Elements<Drawing.DefaultRunProperties>().FirstOrDefault();
            if (defRPr == null)
                continue;

            var style = new DefaultTextStyle
            {
                FontSize = defRPr.FontSize?.Value != null ? defRPr.FontSize.Value / 100.0 : null,
                Bold = defRPr.Bold?.Value,
                Italic = defRPr.Italic?.Value,
                Underline = defRPr.Underline?.Value != null && defRPr.Underline.Value != Drawing.TextUnderlineValues.None,
                Color = ExtractColorFromRunProperties(defRPr),
                Caps = ExtractCapAttribute(defRPr)
            };

            var latinFont = defRPr.Elements<Drawing.LatinFont>().FirstOrDefault();
            if (latinFont?.Typeface != null)
                style.FontFamily = latinFont.Typeface.Value;

            cache.AddStyle(styleType, level.Value, style);
        }
    }

    private static int? ParseLevel(string localName)
    {
        // Format: lvl1pPr, lvl2pPr, etc.
        if (localName.Length < 6 || !localName.StartsWith("lvl") || !localName.EndsWith("pPr"))
            return null;

        var levelStr = localName.Substring(3, localName.Length - 6);
        if (int.TryParse(levelStr, out var level))
            return level - 1; // Convert 1-based to 0-based

        return null;
    }

    private string? ExtractColorFromRunProperties(Drawing.DefaultRunProperties defRPr)
    {
        var solidFill = defRPr.Elements<Drawing.SolidFill>().FirstOrDefault();
        if (solidFill == null)
            return null;

        // Fast path: explicit sRGB color.
        var rgb = solidFill.RgbColorModelHex;
        if (rgb?.Val != null)
            return $"#{rgb.Val.Value}";

        // Scheme color (the common case in master txStyles, e.g. <a:schemeClr val="tx1"/>).
        var schemeClr = solidFill.SchemeColor;
        if (schemeClr == null)
            return null;

        // Raw XML attribute read per AGENTS.pptx.md rule 1 — the SDK enum ToString()
        // yields names like "Text1" which ResolveSchemeColor aliases do not cover.
        var schemeName = GetAttributeValue(schemeClr, "val");
        if (string.IsNullOrEmpty(schemeName))
            return null;

        var resolved = ResolveSchemeColor(schemeName);
        if (resolved == null)
            return null;

        return ApplyLumTransforms(schemeClr, resolved);
    }

    /// <summary>
    /// Applies lumMod/lumOff luminance transforms (HSL space, per ECMA-376) to a resolved
    /// scheme color. Other color transforms (tint, shade, satMod, ...) are not applied.
    /// </summary>
    private static string ApplyLumTransforms(Drawing.SchemeColor schemeClr, string hexColor)
    {
        var lumMod = ReadTransformValue(schemeClr, "lumMod");
        var lumOff = ReadTransformValue(schemeClr, "lumOff");
        if (lumMod == null && lumOff == null)
            return hexColor;

        var mod = (lumMod ?? 100000) / 100000.0;
        var off = (lumOff ?? 0) / 100000.0;

        var r = Convert.ToInt32(hexColor.Substring(1, 2), 16) / 255.0;
        var g = Convert.ToInt32(hexColor.Substring(3, 2), 16) / 255.0;
        var b = Convert.ToInt32(hexColor.Substring(5, 2), 16) / 255.0;

        RgbToHsl(r, g, b, out var h, out var s, out var l);
        l = Math.Clamp(l * mod + off, 0.0, 1.0);
        var (rr, gg, bb) = HslToRgb(h, s, l);

        return string.Create(System.Globalization.CultureInfo.InvariantCulture,
            $"#{(int)Math.Round(rr * 255):X2}{(int)Math.Round(gg * 255):X2}{(int)Math.Round(bb * 255):X2}");
    }

    private static int? ReadTransformValue(Drawing.SchemeColor schemeClr, string localName)
    {
        var element = schemeClr.ChildElements.FirstOrDefault(e => e.LocalName == localName);
        if (element == null)
            return null;

        var valAttr = GetAttributeValue(element, "val");
        return int.TryParse(valAttr, out var val) ? val : null;
    }

    private static void RgbToHsl(double r, double g, double b, out double h, out double s, out double l)
    {
        var max = Math.Max(r, Math.Max(g, b));
        var min = Math.Min(r, Math.Min(g, b));
        l = (max + min) / 2.0;
        h = 0.0;
        s = 0.0;

        if (max == min)
            return;

        var d = max - min;
        s = l > 0.5 ? d / (2.0 - max - min) : d / (max + min);

        if (max == r)
            h = (g - b) / d + (g < b ? 6.0 : 0.0);
        else if (max == g)
            h = (b - r) / d + 2.0;
        else
            h = (r - g) / d + 4.0;
        h /= 6.0;
    }

    private static (double R, double G, double B) HslToRgb(double h, double s, double l)
    {
        if (s == 0.0)
            return (l, l, l);

        var q = l < 0.5 ? l * (1.0 + s) : l + s - l * s;
        var p = 2.0 * l - q;
        return (HueToRgb(p, q, h + 1.0 / 3.0), HueToRgb(p, q, h), HueToRgb(p, q, h - 1.0 / 3.0));
    }

    private static double HueToRgb(double p, double q, double t)
    {
        if (t < 0.0) t += 1.0;
        if (t > 1.0) t -= 1.0;
        if (t < 1.0 / 6.0) return p + (q - p) * 6.0 * t;
        if (t < 1.0 / 2.0) return q;
        if (t < 2.0 / 3.0) return p + (q - p) * (2.0 / 3.0 - t) * 6.0;
        return p;
    }

    public DefaultTextStyle GetLayoutPlaceholderLstStyle(int? idx, PlaceholderValues? type, int level)
    {
        var shape = FindLayoutPlaceholder(idx, type);
        if (shape != null)
            return ExtractLstStyleDefRPr(shape.TextBody, level);
        return new DefaultTextStyle();
    }

    public DefaultTextStyle GetMasterPlaceholderLstStyle(int? idx, PlaceholderValues? type, int level)
    {
        var shape = FindMasterPlaceholder(idx, type);
        if (shape != null)
            return ExtractLstStyleDefRPr(shape.TextBody, level);
        return new DefaultTextStyle();
    }

    #region Bullet Resolution

    public (string? BulletChar, string? AutoNumberType, bool HasBullet, bool HasBulletNone) GetLayoutPlaceholderBulletInfo(int? idx, PlaceholderValues? type, int level)
    {
        var shape = FindLayoutPlaceholder(idx, type);
        if (shape != null)
            return ExtractBulletInfoFromTextBodyLstStyle(shape.TextBody, level);
        return (null, null, false, false);
    }

    public (string? BulletChar, string? AutoNumberType, bool HasBullet, bool HasBulletNone) GetMasterPlaceholderBulletInfo(int? idx, PlaceholderValues? type, int level)
    {
        var shape = FindMasterPlaceholder(idx, type);
        if (shape != null)
            return ExtractBulletInfoFromTextBodyLstStyle(shape.TextBody, level);
        return (null, null, false, false);
    }

    public (string? BulletChar, string? AutoNumberType, bool HasBullet, bool HasBulletNone) GetMasterTxStyleBulletInfo(PlaceholderValues? placeholderType, int level)
    {
        // Shapes without placeholders should NOT inherit body style bullets
        if (placeholderType == null)
            return (null, null, false, false);

        string key;
        if (placeholderType == PlaceholderValues.Title || placeholderType == PlaceholderValues.CenteredTitle)
            key = "Title";
        else if (placeholderType == PlaceholderValues.Body)
            key = "Body";
        else if (placeholderType == PlaceholderValues.SubTitle)
            key = "Body";
        else if (placeholderType == PlaceholderValues.Object)
            key = "Other";
        else
            return (null, null, false, false);

        OpenXmlElement? styleList = key switch
        {
            "Title" => _masterPart?.SlideMaster?.TextStyles?.TitleStyle,
            "Body" => _masterPart?.SlideMaster?.TextStyles?.BodyStyle,
            "Other" => _masterPart?.SlideMaster?.TextStyles?.OtherStyle,
            _ => null
        };

        if (styleList == null)
            return (null, null, false, false);

        var levelName = $"lvl{level + 1}pPr";
        var lvlPpr = styleList.ChildElements.FirstOrDefault(e => e.LocalName == levelName);
        if (lvlPpr == null)
            return (null, null, false, false);

        return ExtractBulletInfo(lvlPpr);
    }

    private static (string? BulletChar, string? AutoNumberType, bool HasBullet, bool HasBulletNone) ExtractBulletInfoFromTextBodyLstStyle(OpenXmlElement? textBody, int level)
    {
        if (textBody == null) return (null, null, false, false);

        var lstStyle = textBody.ChildElements.FirstOrDefault(e => e.LocalName == "lstStyle");
        if (lstStyle == null) return (null, null, false, false);

        var levelName = $"lvl{level + 1}pPr";
        var lvlPpr = lstStyle.ChildElements.FirstOrDefault(e => e.LocalName == levelName);
        if (lvlPpr == null) return (null, null, false, false);

        return ExtractBulletInfo(lvlPpr);
    }

    /// <summary>
    /// Reads bullet color info from a paragraph-properties-like element (a:pPr or lvlNpPr):
    /// a:buClr (explicit srgbClr/schemeClr, scheme resolved through the theme) or
    /// a:buClrTx (bullet follows the text color). Absent buClr/buClrTx returns (null, false).
    /// </summary>
    internal static (string? Color, bool FollowsText) ExtractBulletColorInfo(OpenXmlElement? element, StyleResolver? resolver)
    {
        if (element == null) return (null, false);

        var buClr = element.ChildElements.FirstOrDefault(e => e.LocalName == "buClr");
        if (buClr != null)
        {
            var rgb = buClr.Elements<Drawing.RgbColorModelHex>().FirstOrDefault();
            if (rgb?.Val?.Value != null)
                return ($"#{rgb.Val.Value}", false);

            var schemeClr = buClr.Elements<Drawing.SchemeColor>().FirstOrDefault();
            if (schemeClr != null)
            {
                // Raw XML attribute read per AGENTS.pptx.md rule 1 — SDK enum parsing is unreliable.
                var schemeName = GetAttributeValue(schemeClr, "val");
                if (!string.IsNullOrEmpty(schemeName))
                {
                    var resolved = resolver?.ResolveSchemeColor(schemeName);
                    if (!string.IsNullOrEmpty(resolved))
                        return (resolved, false);
                }
            }
            return (null, false);
        }

        if (element.ChildElements.Any(e => e.LocalName == "buClrTx"))
            return (null, true);

        return (null, false);
    }

    public (string? Color, bool FollowsText) GetLayoutPlaceholderBulletColor(int? idx, PlaceholderValues? type, int level)
    {
        var shape = FindLayoutPlaceholder(idx, type);
        return shape != null ? ExtractBulletColorFromTextBodyLstStyle(shape.TextBody, level) : (null, false);
    }

    public (string? Color, bool FollowsText) GetMasterPlaceholderBulletColor(int? idx, PlaceholderValues? type, int level)
    {
        var shape = FindMasterPlaceholder(idx, type);
        return shape != null ? ExtractBulletColorFromTextBodyLstStyle(shape.TextBody, level) : (null, false);
    }

    public (string? Color, bool FollowsText) GetMasterTxStyleBulletColor(PlaceholderValues? placeholderType, int level)
    {
        // Shapes without placeholders should NOT inherit body style bullet colors
        if (placeholderType == null)
            return (null, false);

        string key;
        if (placeholderType == PlaceholderValues.Title || placeholderType == PlaceholderValues.CenteredTitle)
            key = "Title";
        else if (placeholderType == PlaceholderValues.Body)
            key = "Body";
        else if (placeholderType == PlaceholderValues.SubTitle)
            key = "Body";
        else if (placeholderType == PlaceholderValues.Object)
            key = "Other";
        else
            return (null, false);

        OpenXmlElement? styleList = key switch
        {
            "Title" => _masterPart?.SlideMaster?.TextStyles?.TitleStyle,
            "Body" => _masterPart?.SlideMaster?.TextStyles?.BodyStyle,
            "Other" => _masterPart?.SlideMaster?.TextStyles?.OtherStyle,
            _ => null
        };

        if (styleList == null)
            return (null, false);

        var levelName = $"lvl{level + 1}pPr";
        var lvlPpr = styleList.ChildElements.FirstOrDefault(e => e.LocalName == levelName);
        if (lvlPpr == null)
            return (null, false);

        return ExtractBulletColorInfo(lvlPpr, this);
    }

    /// <summary>
    /// Paragraph alignment (raw <c>algn</c> value) inherited from the master txStyles
    /// (titleStyle/bodyStyle/otherStyle) for placeholder paragraphs — e.g. decks whose
    /// titles are right-aligned via <c>titleStyle algn="r"</c>. Null when not defined.
    /// </summary>
    public string? GetMasterTxStyleAlignment(PlaceholderValues? placeholderType, int level)
    {
        if (placeholderType == null)
            return null;

        OpenXmlElement? styleList;
        if (placeholderType == PlaceholderValues.Title || placeholderType == PlaceholderValues.CenteredTitle)
            styleList = _masterPart?.SlideMaster?.TextStyles?.TitleStyle;
        else if (placeholderType == PlaceholderValues.Body || placeholderType == PlaceholderValues.SubTitle)
            styleList = _masterPart?.SlideMaster?.TextStyles?.BodyStyle;
        else if (placeholderType == PlaceholderValues.Object)
            styleList = _masterPart?.SlideMaster?.TextStyles?.OtherStyle;
        else
            return null;

        if (styleList == null)
            return null;

        var lvlPpr = styleList.ChildElements.FirstOrDefault(e => e.LocalName == $"lvl{level + 1}pPr");
        return ReadAlgnAttribute(lvlPpr);
    }

    /// <summary>
    /// Paragraph alignment inherited from the layout placeholder's own lstStyle.
    /// Layout wins over master in the placeholder inheritance chain.
    /// </summary>
    public string? GetLayoutPlaceholderAlignment(int? idx, PlaceholderValues? type, int level)
    {
        var shape = FindLayoutPlaceholder(idx, type);
        var lstStyle = shape?.TextBody?.ChildElements.FirstOrDefault(e => e.LocalName == "lstStyle");
        var lvlPpr = lstStyle?.ChildElements.FirstOrDefault(e => e.LocalName == $"lvl{level + 1}pPr");
        return ReadAlgnAttribute(lvlPpr);
    }

    private static string? ReadAlgnAttribute(OpenXmlElement? lvlPpr)
    {
        // The lvlNpPr elements derive from TextParagraphPropertiesType; read the typed
        // Alignment property — GetAttribute throws KeyNotFoundException on elements
        // whose schema doesn't declare algn. Use InnerText: in SDK 3.x
        // TextAlignmentTypeValues is a struct whose ToString() does not yield the
        // lexical XML value ("r", "ctr", ...).
        return (lvlPpr as Drawing.TextParagraphPropertiesType)?.Alignment?.InnerText;
    }

    private (string? Color, bool FollowsText) ExtractBulletColorFromTextBodyLstStyle(OpenXmlElement? textBody, int level)
    {
        if (textBody == null) return (null, false);

        var lstStyle = textBody.ChildElements.FirstOrDefault(e => e.LocalName == "lstStyle");
        if (lstStyle == null) return (null, false);

        var levelName = $"lvl{level + 1}pPr";
        var lvlPpr = lstStyle.ChildElements.FirstOrDefault(e => e.LocalName == levelName);
        if (lvlPpr == null) return (null, false);

        return ExtractBulletColorInfo(lvlPpr, this);
    }

    private static (string? BulletChar, string? AutoNumberType, bool HasBullet, bool HasBulletNone) ExtractBulletInfo(OpenXmlElement? element)
    {
        if (element == null) return (null, null, false, false);

        var buChar = element.ChildElements.FirstOrDefault(e => e.LocalName == "buChar");
        if (buChar != null)
        {
            var charAttr = GetAttributeValue(buChar, "char");
            if (!string.IsNullOrEmpty(charAttr))
                return (charAttr, null, true, false);
        }

        var buAutoNum = element.ChildElements.FirstOrDefault(e => e.LocalName == "buAutoNum");
        if (buAutoNum != null)
        {
            var typeAttr = GetAttributeValue(buAutoNum, "type");
            if (!string.IsNullOrEmpty(typeAttr))
                return (null, typeAttr, true, false);
        }

        var buNone = element.ChildElements.FirstOrDefault(e => e.LocalName == "buNone");
        if (buNone != null)
            return (null, null, false, true);

        return (null, null, false, false);
    }

    #endregion

    #region List Indentation Resolution

    public (double? MarginLeft, double? Indent) GetLayoutPlaceholderListIndents(int? idx, PlaceholderValues? type, int level)
    {
        var shape = FindLayoutPlaceholder(idx, type);
        return shape != null ? ExtractListIndentsFromTextBodyLstStyle(shape.TextBody, level) : (null, null);
    }

    public (double? MarginLeft, double? Indent) GetMasterPlaceholderListIndents(int? idx, PlaceholderValues? type, int level)
    {
        var shape = FindMasterPlaceholder(idx, type);
        return shape != null ? ExtractListIndentsFromTextBodyLstStyle(shape.TextBody, level) : (null, null);
    }

    public (double? MarginLeft, double? Indent) GetMasterTxStyleListIndents(PlaceholderValues? placeholderType, int level)
    {
        // Shapes without placeholders should NOT inherit body style indents
        if (placeholderType == null)
            return (null, null);

        string key;
        if (placeholderType == PlaceholderValues.Title || placeholderType == PlaceholderValues.CenteredTitle)
            key = "Title";
        else if (placeholderType == PlaceholderValues.Body)
            key = "Body";
        else if (placeholderType == PlaceholderValues.SubTitle)
            key = "Body";
        else if (placeholderType == PlaceholderValues.Object)
            key = "Other";
        else
            return (null, null);

        OpenXmlElement? styleList = key switch
        {
            "Title" => _masterPart?.SlideMaster?.TextStyles?.TitleStyle,
            "Body" => _masterPart?.SlideMaster?.TextStyles?.BodyStyle,
            "Other" => _masterPart?.SlideMaster?.TextStyles?.OtherStyle,
            _ => null
        };

        if (styleList == null)
            return (null, null);

        var levelName = $"lvl{level + 1}pPr";
        var lvlPpr = styleList.ChildElements.FirstOrDefault(e => e.LocalName == levelName);
        if (lvlPpr == null)
            return (null, null);

        return ExtractListIndentsFromElement(lvlPpr);
    }

    private static (double? MarginLeft, double? Indent) ExtractListIndentsFromTextBodyLstStyle(OpenXmlElement? textBody, int level)
    {
        if (textBody == null) return (null, null);

        var lstStyle = textBody.ChildElements.FirstOrDefault(e => e.LocalName == "lstStyle");
        if (lstStyle == null) return (null, null);

        var levelName = $"lvl{level + 1}pPr";
        var lvlPpr = lstStyle.ChildElements.FirstOrDefault(e => e.LocalName == levelName);
        if (lvlPpr == null) return (null, null);

        return ExtractListIndentsFromElement(lvlPpr);
    }

    private static (double? MarginLeft, double? Indent) ExtractListIndentsFromElement(OpenXmlElement? element)
    {
        if (element == null) return (null, null);

        double? marginLeft = null;
        double? indent = null;

        // Raw XML attribute reads per AGENTS.pptx.md rule 1 — SDK attribute access is unreliable.
        var marLAttr = GetAttributeValue(element, "marL");
        if (!string.IsNullOrEmpty(marLAttr) && int.TryParse(marLAttr, out var marL))
            marginLeft = marL / 12700.0;

        var indentAttr = GetAttributeValue(element, "indent");
        if (!string.IsNullOrEmpty(indentAttr) && int.TryParse(indentAttr, out var ind))
            indent = ind / 12700.0;

        return (marginLeft, indent);
    }

    #endregion

    #region Line Spacing Resolution

    public double? GetLayoutPlaceholderLineSpacing(int? idx, PlaceholderValues? type, int level)
    {
        var shape = FindLayoutPlaceholder(idx, type);
        if (shape != null)
            return ExtractLineSpacingFromTextBodyLstStyle(shape.TextBody, level);
        return null;
    }

    public double? GetMasterPlaceholderLineSpacing(int? idx, PlaceholderValues? type, int level)
    {
        var shape = FindMasterPlaceholder(idx, type);
        if (shape != null)
            return ExtractLineSpacingFromTextBodyLstStyle(shape.TextBody, level);
        return null;
    }

    public double? GetMasterTxStyleLineSpacing(PlaceholderValues? placeholderType, int level)
    {
        // Shapes without placeholders should NOT inherit body style line spacing
        if (placeholderType == null)
            return null;

        string key;
        if (placeholderType == PlaceholderValues.Title || placeholderType == PlaceholderValues.CenteredTitle)
            key = "Title";
        else if (placeholderType == PlaceholderValues.Body)
            key = "Body";
        else if (placeholderType == PlaceholderValues.SubTitle)
            key = "Body";
        else if (placeholderType == PlaceholderValues.Object)
            key = "Other";
        else
            return null;

        OpenXmlElement? styleList = key switch
        {
            "Title" => _masterPart?.SlideMaster?.TextStyles?.TitleStyle,
            "Body" => _masterPart?.SlideMaster?.TextStyles?.BodyStyle,
            "Other" => _masterPart?.SlideMaster?.TextStyles?.OtherStyle,
            _ => null
        };

        if (styleList == null)
            return null;

        var levelName = $"lvl{level + 1}pPr";
        var lvlPpr = styleList.ChildElements.FirstOrDefault(e => e.LocalName == levelName);
        if (lvlPpr == null)
            return null;

        return ExtractLineSpacingFromElement(lvlPpr);
    }

    #endregion

    #region Paragraph Spacing Resolution

    public (double? SpaceBefore, double? SpaceAfter) GetLayoutPlaceholderSpacing(int? idx, PlaceholderValues? type, int level)
    {
        var shape = FindLayoutPlaceholder(idx, type);
        if (shape != null)
            return ExtractSpacingFromTextBodyLstStyle(shape.TextBody, level);
        return (null, null);
    }

    public (double? SpaceBefore, double? SpaceAfter) GetMasterPlaceholderSpacing(int? idx, PlaceholderValues? type, int level)
    {
        var shape = FindMasterPlaceholder(idx, type);
        if (shape != null)
            return ExtractSpacingFromTextBodyLstStyle(shape.TextBody, level);
        return (null, null);
    }

    public (double? SpaceBefore, double? SpaceAfter) GetMasterTxStyleSpacing(PlaceholderValues? placeholderType, int level)
    {
        // Shapes without placeholders should NOT inherit body style spacing
        if (placeholderType == null)
            return (null, null);

        string key;
        if (placeholderType == PlaceholderValues.Title || placeholderType == PlaceholderValues.CenteredTitle)
            key = "Title";
        else if (placeholderType == PlaceholderValues.Body)
            key = "Body";
        else if (placeholderType == PlaceholderValues.SubTitle)
            key = "Body";
        else if (placeholderType == PlaceholderValues.Object)
            key = "Other";
        else
            return (null, null);

        OpenXmlElement? styleList = key switch
        {
            "Title" => _masterPart?.SlideMaster?.TextStyles?.TitleStyle,
            "Body" => _masterPart?.SlideMaster?.TextStyles?.BodyStyle,
            "Other" => _masterPart?.SlideMaster?.TextStyles?.OtherStyle,
            _ => null
        };

        if (styleList == null)
            return (null, null);

        var levelName = $"lvl{level + 1}pPr";
        var lvlPpr = styleList.ChildElements.FirstOrDefault(e => e.LocalName == levelName);
        if (lvlPpr == null)
            return (null, null);

        return ExtractSpacingFromElement(lvlPpr);
    }

    private static (double? SpaceBefore, double? SpaceAfter) ExtractSpacingFromTextBodyLstStyle(OpenXmlElement? textBody, int level)
    {
        if (textBody == null) return (null, null);

        var lstStyle = textBody.ChildElements.FirstOrDefault(e => e.LocalName == "lstStyle");
        if (lstStyle == null) return (null, null);

        var levelName = $"lvl{level + 1}pPr";
        var lvlPpr = lstStyle.ChildElements.FirstOrDefault(e => e.LocalName == levelName);
        if (lvlPpr == null) return (null, null);

        return ExtractSpacingFromElement(lvlPpr);
    }

    private static (double? SpaceBefore, double? SpaceAfter) ExtractSpacingFromElement(OpenXmlElement? element)
    {
        if (element == null) return (null, null);

        double? spcBef = null;
        double? spcAft = null;

        var spcBefEl = element.ChildElements.FirstOrDefault(e => e.LocalName == "spcBef");
        if (spcBefEl != null)
        {
            var spcPts = spcBefEl.ChildElements.FirstOrDefault(e => e.LocalName == "spcPts");
            if (spcPts != null)
            {
                var valAttr = GetAttributeValue(spcPts, "val");
                if (int.TryParse(valAttr, out var ptsHundredths))
                    spcBef = ptsHundredths / 100.0;
            }
            var spcPct = spcBefEl.ChildElements.FirstOrDefault(e => e.LocalName == "spcPct");
            if (spcPct != null)
            {
                var valAttr = GetAttributeValue(spcPct, "val");
                if (int.TryParse(valAttr, out var pct))
                    spcBef = pct / 100000.0;
            }
        }

        var spcAftEl = element.ChildElements.FirstOrDefault(e => e.LocalName == "spcAft");
        if (spcAftEl != null)
        {
            var spcPts = spcAftEl.ChildElements.FirstOrDefault(e => e.LocalName == "spcPts");
            if (spcPts != null)
            {
                var valAttr = GetAttributeValue(spcPts, "val");
                if (int.TryParse(valAttr, out var ptsHundredths))
                    spcAft = ptsHundredths / 100.0;
            }
            var spcPct = spcAftEl.ChildElements.FirstOrDefault(e => e.LocalName == "spcPct");
            if (spcPct != null)
            {
                var valAttr = GetAttributeValue(spcPct, "val");
                if (int.TryParse(valAttr, out var pct))
                    spcAft = pct / 100000.0;
            }
        }

        return (spcBef, spcAft);
    }

    #endregion

    #region Line Spacing Helpers

    private static double? ExtractLineSpacingFromTextBodyLstStyle(OpenXmlElement? textBody, int level)
    {
        if (textBody == null) return null;

        var lstStyle = textBody.ChildElements.FirstOrDefault(e => e.LocalName == "lstStyle");
        if (lstStyle == null) return null;

        var levelName = $"lvl{level + 1}pPr";
        var lvlPpr = lstStyle.ChildElements.FirstOrDefault(e => e.LocalName == levelName);
        if (lvlPpr == null) return null;

        return ExtractLineSpacingFromElement(lvlPpr);
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
            var valAttr = GetAttributeValue(spcPts, "val");
            if (int.TryParse(valAttr, out var value))
                return value / 100.0;
        }

        // Try spcPct (percentage of line height)
        var spcPct = lnSpc.ChildElements.FirstOrDefault(e => e.LocalName == "spcPct");
        if (spcPct != null)
        {
            var valAttr = GetAttributeValue(spcPct, "val");
            if (int.TryParse(valAttr, out var value))
                return value / 100000.0; // spcPct is in 1/1000ths of a percent (100000 = 100% = 1.0)
        }

        return null;
    }

    #endregion

    private static int? GetPlaceholderIdx(Shape shape)
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

        var outerXml = ph.OuterXml;
        if (!string.IsNullOrEmpty(outerXml))
        {
            var match = System.Text.RegularExpressions.Regex.Match(outerXml, @"idx\s*=\s*""([^""]*)""");
            if (match.Success && int.TryParse(match.Groups[1].Value, out var idx))
                return idx;
        }
        return null;
    }

    private DefaultTextStyle ExtractLstStyleDefRPr(OpenXmlElement? textBody, int level)
    {
        if (textBody == null) return new DefaultTextStyle();

        var lstStyle = textBody.ChildElements.FirstOrDefault(e => e.LocalName == "lstStyle");
        if (lstStyle == null) return new DefaultTextStyle();

        var levelName = $"lvl{level + 1}pPr";
        var lvlPpr = lstStyle.ChildElements.FirstOrDefault(e => e.LocalName == levelName);
        if (lvlPpr == null) return new DefaultTextStyle();

        var defRPr = lvlPpr.Elements<Drawing.DefaultRunProperties>().FirstOrDefault();
        if (defRPr == null) return new DefaultTextStyle();

        var style = new DefaultTextStyle
        {
            FontSize = defRPr.FontSize?.Value != null ? defRPr.FontSize.Value / 100.0 : null,
            Bold = defRPr.Bold?.Value,
            Italic = defRPr.Italic?.Value,
            Underline = defRPr.Underline?.Value != null && defRPr.Underline.Value != Drawing.TextUnderlineValues.None,
            Color = ExtractColorFromRunProperties(defRPr),
            Caps = ExtractCapAttribute(defRPr)
        };

        var latinFont = defRPr.Elements<Drawing.LatinFont>().FirstOrDefault();
        if (latinFont?.Typeface != null)
            style.FontFamily = latinFont.Typeface.Value;

        return style;
    }

    private static string? ExtractCapAttribute(OpenXmlElement element)
    {
        var match = System.Text.RegularExpressions.Regex.Match(element.OuterXml, @"\bcap\s*=\s*""([^""]*)""");
        if (match.Success)
        {
            var value = match.Groups[1].Value;
            if (value == "none")
                return null;
            return value;
        }
        return null;
    }

    private static PlaceholderValues? GetPlaceholderType(Shape shape)
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

        var outerXml = ph.OuterXml;
        var match = System.Text.RegularExpressions.Regex.Match(outerXml, @"type\s*=\s*""([^""]*)""");
        if (match.Success)
        {
            return match.Groups[1].Value switch
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
                _ => null
            };
        }
        return null;
    }

    private Shape? FindLayoutPlaceholder(int? idx, PlaceholderValues? type)
    {
        if (_layoutPart?.SlideLayout?.CommonSlideData?.ShapeTree == null)
            return null;

        if (idx.HasValue)
        {
            foreach (var layoutShape in _layoutPart.SlideLayout.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>())
            {
                var shapeIdx = GetPlaceholderIdx(layoutShape);
                if (shapeIdx == idx.Value)
                {
                    if (type.HasValue)
                    {
                        var shapeType = GetPlaceholderType(layoutShape);
                        if (shapeType.HasValue && shapeType.Value != type.Value)
                            continue;
                    }
                    return layoutShape;
                }
            }
        }

        if (type.HasValue)
        {
            foreach (var layoutShape in _layoutPart.SlideLayout.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>())
            {
                var shapeType = GetPlaceholderType(layoutShape);
                if (shapeType == type.Value)
                    return layoutShape;
            }
        }

        return null;
    }

    private Shape? FindMasterPlaceholder(int? idx, PlaceholderValues? type)
    {
        if (_masterPart?.SlideMaster?.CommonSlideData?.ShapeTree == null)
            return null;

        if (idx.HasValue)
        {
            foreach (var masterShape in _masterPart.SlideMaster.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>())
            {
                var shapeIdx = GetPlaceholderIdx(masterShape);
                if (shapeIdx == idx.Value)
                {
                    if (type.HasValue)
                    {
                        var shapeType = GetPlaceholderType(masterShape);
                        if (shapeType.HasValue && shapeType.Value != type.Value)
                            continue;
                    }
                    return masterShape;
                }
            }
        }

        if (type.HasValue)
        {
            foreach (var masterShape in _masterPart.SlideMaster.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>())
            {
                var shapeType = GetPlaceholderType(masterShape);
                if (shapeType == type.Value)
                    return masterShape;
            }
        }

        return null;
    }

    public Drawing.BodyProperties? GetLayoutPlaceholderBodyPr(int? idx, PlaceholderValues? type)
    {
        var shape = FindLayoutPlaceholder(idx, type);
        return shape?.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
    }

    public Drawing.BodyProperties? GetMasterPlaceholderBodyPr(int? idx, PlaceholderValues? type)
    {
        var shape = FindMasterPlaceholder(idx, type);
        return shape?.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
    }

    /// <summary>
    /// Shape properties (&lt;p:spPr&gt;) of the layout placeholder matching idx/type, or
    /// null when the layout has no such placeholder. Used for placeholder fill/line
    /// inheritance (ECMA-376: a slide placeholder without its own fill/line takes them
    /// from the layout placeholder).
    /// </summary>
    public ShapeProperties? GetLayoutPlaceholderShapeProperties(int? idx, PlaceholderValues? type)
        => FindLayoutPlaceholder(idx, type)?.ShapeProperties;

    /// <summary>
    /// Shape properties (&lt;p:spPr&gt;) of the master placeholder matching idx/type —
    /// the final fallback of placeholder fill/line inheritance after the layout.
    /// </summary>
    public ShapeProperties? GetMasterPlaceholderShapeProperties(int? idx, PlaceholderValues? type)
        => FindMasterPlaceholder(idx, type)?.ShapeProperties;

    /// <summary>
    /// Resolves an &lt;a:solidFill&gt; to a hex color: explicit sRGB wins; scheme colors
    /// go through the theme with lumMod/lumOff luminance transforms applied (HSL space,
    /// per ECMA-376). Returns null when no usable color is present.
    /// </summary>
    public string? ResolveSolidFillColor(Drawing.SolidFill solidFill)
    {
        var rgb = solidFill.RgbColorModelHex;
        if (rgb?.Val != null)
            return $"#{rgb.Val.Value}";

        var schemeClr = solidFill.SchemeColor;
        if (schemeClr == null)
            return null;

        var schemeName = GetAttributeValue(schemeClr, "val");
        if (string.IsNullOrEmpty(schemeName))
            return null;

        var resolved = ResolveSchemeColor(schemeName);
        return resolved == null ? null : ApplyLumTransforms(schemeClr, resolved);
    }

    #endregion

    #region Theme/Scheme Color Resolution

    private Dictionary<string, string> LoadSchemeColors()
    {
        var colors = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        if (_themePart?.Theme?.ThemeElements?.ColorScheme == null)
            return colors;

        var scheme = _themePart.Theme.ThemeElements.ColorScheme;

        AddSchemeColor(colors, "dk1", scheme.Dark1Color);
        AddSchemeColor(colors, "lt1", scheme.Light1Color);
        AddSchemeColor(colors, "dk2", scheme.Dark2Color);
        AddSchemeColor(colors, "lt2", scheme.Light2Color);
        AddSchemeColor(colors, "accent1", scheme.Accent1Color);
        AddSchemeColor(colors, "accent2", scheme.Accent2Color);
        AddSchemeColor(colors, "accent3", scheme.Accent3Color);
        AddSchemeColor(colors, "accent4", scheme.Accent4Color);
        AddSchemeColor(colors, "accent5", scheme.Accent5Color);
        AddSchemeColor(colors, "accent6", scheme.Accent6Color);
        AddSchemeColor(colors, "hlink", scheme.Hyperlink);
        // FollowedHyperlink may not exist in all SDK versions, skip if not available
        try
        {
            var folHlink = scheme.Elements().FirstOrDefault(e => e.LocalName == "folHlink");
            if (folHlink != null)
                AddSchemeColor(colors, "folHlink", folHlink);
        }
        catch { /* folHlink may be absent in some themes */ }

        return colors;
    }

    private static void AddSchemeColor(Dictionary<string, string> colors, string name, OpenXmlElement? colorType)
    {
        if (colorType == null)
            return;

        // Check for sRGB
        var rgb = colorType.Elements<Drawing.RgbColorModelHex>().FirstOrDefault();
        if (rgb?.Val != null)
        {
            colors[name] = $"#{rgb.Val.Value}";
            return;
        }

        // Check for system color
        var sysClr = colorType.Elements<Drawing.SystemColor>().FirstOrDefault();
        if (sysClr?.LastColor?.Value != null)
        {
            colors[name] = $"#{sysClr.LastColor.Value}";
            return;
        }

        // Check for scheme color reference (rare but possible)
        var schemeClr = colorType.Elements<Drawing.SchemeColor>().FirstOrDefault();
        if (schemeClr?.Val != null)
        {
            colors[name] = "#000000";
        }
    }

    #endregion

    #region Helper Methods

    private SlideLayoutPart? GetSlideLayoutPart()
    {
        // Use the SDK's built-in property which correctly resolves the layout relationship
        return _slidePart.SlideLayoutPart;
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

    #endregion

    /// <summary>
    /// Default text style values from master
    /// </summary>
    public class DefaultTextStyle
    {
        public double? FontSize { get; set; }
        public bool? Bold { get; set; }
        public bool? Italic { get; set; }
        public bool? Underline { get; set; }
        public string? Color { get; set; }
        public string? FontFamily { get; set; }
        public string? Caps { get; set; }
    }

    /// <summary>
    /// Caches text styles by type and level
    /// </summary>
    private class TextStyleCache
    {
        private readonly Dictionary<string, DefaultTextStyle> _styles = new();

        public void AddStyle(string styleType, int level, DefaultTextStyle style)
        {
            var key = $"{styleType}:{level}";
            _styles[key] = style;
        }

        /// <summary>
        /// Exact-match lookup with no fallback — returns null when the styleType:level
        /// entry was never loaded (used by brand profile extraction).
        /// </summary>
        public DefaultTextStyle? TryGetStyle(string styleType, int level)
        {
            var key = $"{styleType}:{level}";
            return _styles.TryGetValue(key, out var style) ? style : null;
        }

        public DefaultTextStyle GetStyle(string styleType, int level = 0)
        {
            var key = $"{styleType}:{level}";
            if (_styles.TryGetValue(key, out var style))
                return style;

            // Fallback to level 0
            if (level > 0)
                return GetStyle(styleType, 0);

            // Fallback to body level 0
            if (styleType != "Body" && _styles.TryGetValue("Body:0", out var bodyStyle))
                return bodyStyle;

            return new DefaultTextStyle();
        }
    }
}
