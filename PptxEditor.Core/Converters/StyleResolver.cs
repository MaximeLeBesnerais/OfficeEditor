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
        else
            return new DefaultTextStyle();

        return _textStyles.GetStyle(key);
    }

    /// <summary>
    /// Resolves a scheme color to RGB
    /// </summary>
    public string? ResolveSchemeColor(string schemeColorName)
    {
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
        return null;
    }

    private string? GetLayoutBackground()
    {
        var bg = _layoutPart?.SlideLayout?.CommonSlideData?.Background;
        if (bg?.BackgroundProperties != null)
        {
            return ExtractBackgroundColor(bg.BackgroundProperties);
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
        return null;
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

    private static void LoadLevelStyles(TextStyleCache cache, string styleType, OpenXmlElement? styleList)
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
                Color = ExtractColorFromRunProperties(defRPr)
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

    private static string? ExtractColorFromRunProperties(Drawing.DefaultRunProperties defRPr)
    {
        var solidFill = defRPr.Elements<Drawing.SolidFill>().FirstOrDefault();
        if (solidFill == null)
            return null;

        var rgb = solidFill.RgbColorModelHex;
        if (rgb?.Val != null)
            return $"#{rgb.Val.Value}";

        return null;
    }

    public DefaultTextStyle GetLayoutPlaceholderLstStyle(int? idx, int level)
    {
        if (_layoutPart?.SlideLayout?.CommonSlideData?.ShapeTree == null || !idx.HasValue)
            return new DefaultTextStyle();

        foreach (var layoutShape in _layoutPart.SlideLayout.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>())
        {
            var shapeIdx = GetPlaceholderIdx(layoutShape);
            if (shapeIdx == idx.Value)
            {
                return ExtractLstStyleDefRPr(layoutShape.TextBody, level);
            }
        }
        return new DefaultTextStyle();
    }

    public DefaultTextStyle GetMasterPlaceholderLstStyle(int? idx, int level)
    {
        if (_masterPart?.SlideMaster?.CommonSlideData?.ShapeTree == null || !idx.HasValue)
            return new DefaultTextStyle();

        foreach (var masterShape in _masterPart.SlideMaster.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>())
        {
            var shapeIdx = GetPlaceholderIdx(masterShape);
            if (shapeIdx == idx.Value)
            {
                return ExtractLstStyleDefRPr(masterShape.TextBody, level);
            }
        }
        return new DefaultTextStyle();
    }

    #region Bullet Resolution

    public (string? BulletChar, string? AutoNumberType, bool HasBullet, bool HasBulletNone) GetLayoutPlaceholderBulletInfo(int? idx, int level)
    {
        if (_layoutPart?.SlideLayout?.CommonSlideData?.ShapeTree == null || !idx.HasValue)
            return (null, null, false, false);

        foreach (var layoutShape in _layoutPart.SlideLayout.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>())
        {
            var shapeIdx = GetPlaceholderIdx(layoutShape);
            if (shapeIdx == idx.Value)
            {
                return ExtractBulletInfoFromTextBodyLstStyle(layoutShape.TextBody, level);
            }
        }
        return (null, null, false, false);
    }

    public (string? BulletChar, string? AutoNumberType, bool HasBullet, bool HasBulletNone) GetMasterPlaceholderBulletInfo(int? idx, int level)
    {
        if (_masterPart?.SlideMaster?.CommonSlideData?.ShapeTree == null || !idx.HasValue)
            return (null, null, false, false);

        foreach (var masterShape in _masterPart.SlideMaster.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>())
        {
            var shapeIdx = GetPlaceholderIdx(masterShape);
            if (shapeIdx == idx.Value)
            {
                return ExtractBulletInfoFromTextBodyLstStyle(masterShape.TextBody, level);
            }
        }
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

    #endregion

    #region Line Spacing Resolution

    public double? GetLayoutPlaceholderLineSpacing(int? idx, int level)
    {
        if (_layoutPart?.SlideLayout?.CommonSlideData?.ShapeTree == null || !idx.HasValue)
            return null;

        foreach (var layoutShape in _layoutPart.SlideLayout.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>())
        {
            var shapeIdx = GetPlaceholderIdx(layoutShape);
            if (shapeIdx == idx.Value)
            {
                return ExtractLineSpacingFromTextBodyLstStyle(layoutShape.TextBody, level);
            }
        }
        return null;
    }

    public double? GetMasterPlaceholderLineSpacing(int? idx, int level)
    {
        if (_masterPart?.SlideMaster?.CommonSlideData?.ShapeTree == null || !idx.HasValue)
            return null;

        foreach (var masterShape in _masterPart.SlideMaster.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>())
        {
            var shapeIdx = GetPlaceholderIdx(masterShape);
            if (shapeIdx == idx.Value)
            {
                return ExtractLineSpacingFromTextBodyLstStyle(masterShape.TextBody, level);
            }
        }
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
            var valAttr = spcPts.GetAttribute("val", "");
            if (int.TryParse(valAttr.Value, out var value))
                return value / 100.0;
        }

        // Try spcPct (percentage of line height)
        var spcPct = lnSpc.ChildElements.FirstOrDefault(e => e.LocalName == "spcPct");
        if (spcPct != null)
        {
            var valAttr = spcPct.GetAttribute("val", "");
            if (int.TryParse(valAttr.Value, out var value))
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

    private static DefaultTextStyle ExtractLstStyleDefRPr(OpenXmlElement? textBody, int level)
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
            Color = ExtractColorFromRunProperties(defRPr)
        };

        var latinFont = defRPr.Elements<Drawing.LatinFont>().FirstOrDefault();
        if (latinFont?.Typeface != null)
            style.FontFamily = latinFont.Typeface.Value;

        return style;
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
        catch { }

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
