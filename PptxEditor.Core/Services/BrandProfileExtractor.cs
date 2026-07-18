using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeEditor.Core.Exceptions;
using PptxEditor.Core.Converters;
using PptxEditor.Core.Models;

namespace PptxEditor.Core.Services;

/// <summary>
/// Extracts an immutable <see cref="BrandProfile"/> from a presentation. Read-only:
/// no document part is ever written.
///
/// Scope assumptions:
/// - FIRST MASTER ONLY: theme colors, theme fonts and text-style defaults are resolved
///   from the slide master reached via the FIRST slide of the SlideIdList
///   (first slide part -> layout -> master -> theme). Decks with multiple masters only
///   report the first master's brand. Resolver-scoped data (background cascade,
///   per-level text styles) is likewise read through the first slide's part via
///   <see cref="StyleResolver"/>.
/// - Custom clrMap remapping is detected and surfaced in <see cref="BrandProfile.Warnings"/>
///   but NOT applied — ThemeColors always holds raw scheme slots.
/// - Backgrounds expressed as a theme bgRef (background style reference) yield a null
///   BackgroundColor; only explicit solid-fill backgrounds are resolved.
/// </summary>
public sealed class BrandProfileExtractor
{
    private const double EmuPerPoint = 12700.0;

    // PowerPoint default slide size (10in x 7.5in) used when p:sldSz is absent.
    private const long DefaultSlideWidthEmu = 9144000;
    private const long DefaultSlideHeightEmu = 6858000;

    // Identity clrMap: slide color slots map to the same-named scheme slots.
    private static readonly (string Attribute, string Identity)[] ColorMapSlots =
    [
        ("bg1", "lt1"), ("tx1", "dk1"), ("bg2", "lt2"), ("tx2", "dk2"),
        ("accent1", "accent1"), ("accent2", "accent2"), ("accent3", "accent3"),
        ("accent4", "accent4"), ("accent5", "accent5"), ("accent6", "accent6"),
        ("hlink", "hlink"), ("folHlink", "folHlink")
    ];

    /// <summary>
    /// Extracts the brand profile from the first slide master of the presentation.
    /// </summary>
    /// <exception cref="ArgumentNullException">document is null.</exception>
    /// <exception cref="OfficeEditorException">the presentation has no presentation part,
    /// no Presentation root, or no slides (a slide part is required to scope the master chain).</exception>
    public BrandProfile Extract(PresentationDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);

        var presentationPart = document.PresentationPart
            ?? throw new OfficeEditorException("Presentation has no presentation part.");
        var presentation = presentationPart.Presentation
            ?? throw new OfficeEditorException("Presentation part has no Presentation root element.");

        var firstSlidePart = GetFirstSlidePart(presentationPart);
        var masterPart = firstSlidePart.SlideLayoutPart?.SlideMasterPart;

        var resolver = new StyleResolver(document, firstSlidePart);

        var warnings = new List<string>();
        warnings.AddRange(DetectColorMapRemapping(masterPart));

        var (majorFont, minorFont) = ExtractThemeFonts(masterPart);

        var (widthPt, heightPt, slideSizeWarning) = ResolveSlideSizePt(presentation);
        if (slideSizeWarning != null)
            warnings.Add(slideSizeWarning);

        // Sorted by key (ordinal) for deterministic brand.json output.
        var themeColors = new SortedDictionary<string, string>(StringComparer.Ordinal);
        foreach (var kvp in resolver.SchemeColors)
            themeColors[kvp.Key] = kvp.Value;

        return new BrandProfile
        {
            ThemeColors = themeColors,
            MajorFont = majorFont,
            MinorFont = minorFont,
            Title = ToBrandTextStyle(resolver.GetTextStyle("Title", 0), majorFont, minorFont),
            Body = ToBrandTextStyle(resolver.GetTextStyle("Body", 0), majorFont, minorFont),
            Other = ToBrandTextStyle(resolver.GetTextStyle("Other", 0), majorFont, minorFont),
            BackgroundColor = resolver.ResolveBackgroundColor(),
            SlideWidthPt = widthPt,
            SlideHeightPt = heightPt,
            Warnings = warnings
        };
    }

    private static SlidePart GetFirstSlidePart(PresentationPart presentationPart)
    {
        var firstSlideId = presentationPart.Presentation!.SlideIdList?.ChildElements.OfType<SlideId>().FirstOrDefault()
            ?? throw new OfficeEditorException("Presentation contains no slides; brand extraction requires at least one slide to scope the first master.");

        return (SlidePart)presentationPart.GetPartById(firstSlideId.RelationshipId!);
    }

    // DEDUP-INTEGRATION: unify with PptxToTypstConverter.ExtractThemeFonts (the converter
    // keeps its own private copy, owned by another workstream, until integration time).
    private static (string? Major, string? Minor) ExtractThemeFonts(SlideMasterPart? masterPart)
    {
        var fontScheme = masterPart?.ThemePart?.Theme?.ThemeElements?.FontScheme;
        if (fontScheme == null)
            return (null, null);

        var major = fontScheme.MajorFont?.LatinFont?.Typeface?.Value;
        var minor = fontScheme.MinorFont?.LatinFont?.Typeface?.Value;

        return (string.IsNullOrWhiteSpace(major) ? null : major,
                string.IsNullOrWhiteSpace(minor) ? null : minor);
    }

    private static BrandTextStyle? ToBrandTextStyle(StyleResolver.DefaultTextStyle? style, string? majorFont, string? minorFont)
    {
        if (style == null)
            return null;

        return new BrandTextStyle(
            ResolveThemeFontReference(style.FontFamily, majorFont, minorFont),
            style.FontSize,
            style.Bold,
            style.Italic,
            style.Color);
    }

    private static string? ResolveThemeFontReference(string? fontFamily, string? majorFont, string? minorFont)
        => fontFamily switch
        {
            "+mj-lt" => majorFont ?? fontFamily,
            "+mn-lt" => minorFont ?? fontFamily,
            _ => fontFamily
        };

    private static (double WidthPt, double HeightPt, string? Warning) ResolveSlideSizePt(Presentation presentation)
    {
        var cx = presentation.SlideSize?.Cx?.Value;
        var cy = presentation.SlideSize?.Cy?.Value;

        if (cx == null || cy == null)
        {
            return (DefaultSlideWidthEmu / EmuPerPoint, DefaultSlideHeightEmu / EmuPerPoint,
                "Presentation has no explicit slide size; assuming the PowerPoint default of 720x540 pt.");
        }

        return (cx.Value / EmuPerPoint, cy.Value / EmuPerPoint, null);
    }

    private static List<string> DetectColorMapRemapping(SlideMasterPart? masterPart)
    {
        var colorMap = masterPart?.SlideMaster?.ColorMap;
        if (colorMap == null)
            return [];

        var outerXml = colorMap.OuterXml;
        var remapped = new List<string>();
        foreach (var (attribute, identity) in ColorMapSlots)
        {
            // Regex-on-OuterXml per AGENTS.pptx.md rule 1 (optional attributes).
            var value = ReadAttribute(outerXml, attribute);
            if (value != null && !string.Equals(value, identity, StringComparison.OrdinalIgnoreCase))
                remapped.Add($"{attribute}->{value}");
        }

        if (remapped.Count == 0)
            return [];

        return
        [
            $"Slide master uses a custom clrMap remapping ({string.Join(", ", remapped)}); " +
            "ThemeColors holds raw scheme slots which may not match the rendered colors."
        ];
    }

    private static string? ReadAttribute(string outerXml, string attributeName)
    {
        var match = System.Text.RegularExpressions.Regex.Match(
            outerXml,
            $@"\b{System.Text.RegularExpressions.Regex.Escape(attributeName)}\s*=\s*""([^""]*)""");

        return match.Success ? match.Groups[1].Value : null;
    }
}
