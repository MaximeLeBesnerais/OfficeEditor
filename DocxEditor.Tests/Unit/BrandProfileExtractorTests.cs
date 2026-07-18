using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeEditor.Core.Exceptions;
using PptxEditor.Core.Converters;
using PptxEditor.Core.Services;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace DocxEditor.Tests.Unit;

public sealed class BrandProfileExtractorTests : IDisposable
{
    private const string DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private readonly string _tempDir = Path.Combine(Path.GetTempPath(), "BrandProfileExtractorTests", Guid.NewGuid().ToString("N"));

    public BrandProfileExtractorTests() => Directory.CreateDirectory(_tempDir);

    [Fact]
    public void Extract_ThemeColors_ResolvesAllSchemeSlotsIncludingSysClrLastColor()
    {
        var path = CreatePresentation();
        using var document = PresentationDocument.Open(path, false);

        var profile = new BrandProfileExtractor().Extract(document);

        Assert.Equal(12, profile.ThemeColors.Count);
        Assert.Equal("#000000", profile.ThemeColors["dk1"]); // sysClr windowText -> LastColor
        Assert.Equal("#FFFFFF", profile.ThemeColors["lt1"]); // sysClr window -> LastColor
        Assert.Equal("#1F497D", profile.ThemeColors["dk2"]);
        Assert.Equal("#EEECE1", profile.ThemeColors["lt2"]);
        Assert.Equal("#4F81BD", profile.ThemeColors["accent1"]);
        Assert.Equal("#C0504D", profile.ThemeColors["accent2"]);
        Assert.Equal("#9BBB59", profile.ThemeColors["accent3"]);
        Assert.Equal("#8064A2", profile.ThemeColors["accent4"]);
        Assert.Equal("#4BACC6", profile.ThemeColors["accent5"]);
        Assert.Equal("#F79646", profile.ThemeColors["accent6"]);
        Assert.Equal("#0000FF", profile.ThemeColors["hlink"]);
        Assert.Equal("#800080", profile.ThemeColors["folHlink"]);
    }

    [Fact]
    public void Extract_ThemeFonts_ReturnsMajorMinorLatinTypefaces()
    {
        var path = CreatePresentation(majorFont: "Aptos Display", minorFont: "Aptos");
        using var document = PresentationDocument.Open(path, false);

        var profile = new BrandProfileExtractor().Extract(document);

        Assert.Equal("Aptos Display", profile.MajorFont);
        Assert.Equal("Aptos", profile.MinorFont);
    }

    [Fact]
    public void Extract_TextStyles_ResolveSchemeColoredFillsViaSchemeMap()
    {
        // Master txStyles overwhelmingly use <a:schemeClr> (e.g. tx1) for text fills;
        // before the StyleResolver fix these resolved to null.
        var title = new P.TitleStyle(ListStyleLevel(1, 4400, bold: true, fill: SchemeFill(Drawing.SchemeColorValues.Text1)));
        var body = new P.BodyStyle(ListStyleLevel(1, 2000, fill: SchemeFill(Drawing.SchemeColorValues.Accent1)));
        var other = new P.OtherStyle(ListStyleLevel(1, 1200, fill: RgbFill("123456")));
        var path = CreatePresentation(masterTextStyles: new TextStyles(title, body, other));
        using var document = PresentationDocument.Open(path, false);

        var profile = new BrandProfileExtractor().Extract(document);

        Assert.NotNull(profile.Title);
        Assert.Equal("#000000", profile.Title!.Color); // tx1 -> dk1 -> sysClr LastColor
        Assert.Equal(44, profile.Title.FontSizePt);
        Assert.True(profile.Title.Bold);

        Assert.NotNull(profile.Body);
        Assert.Equal("#4F81BD", profile.Body!.Color); // accent1 srgbClr
        Assert.Equal(20, profile.Body.FontSizePt);

        Assert.NotNull(profile.Other);
        Assert.Equal("#123456", profile.Other!.Color); // explicit srgbClr fast path preserved
        Assert.Equal(12, profile.Other.FontSizePt);
    }

    [Fact]
    public void Extract_TextStyles_ResolveThemeFontReferencesToMajorMinorFonts()
    {
        var title = new P.TitleStyle(ListStyleLevel(1, 4400, font: "+mj-lt"));
        var body = new P.BodyStyle(ListStyleLevel(1, 2000, font: "+mn-lt"));
        var other = new P.OtherStyle(ListStyleLevel(1, 1200, font: "Carlito"));
        var path = CreatePresentation(
            majorFont: "Aptos Display",
            minorFont: "Aptos",
            masterTextStyles: new TextStyles(title, body, other));
        using var document = PresentationDocument.Open(path, false);

        var profile = new BrandProfileExtractor().Extract(document);

        Assert.Equal("Aptos Display", profile.Title!.FontFamily);
        Assert.Equal("Aptos", profile.Body!.FontFamily);
        Assert.Equal("Carlito", profile.Other!.FontFamily); // explicit typeface untouched
    }

    [Fact]
    public void Extract_TextStyles_ApplyLumModAndLumOffTransforms()
    {
        var titleFill = SchemeFill(Drawing.SchemeColorValues.Light1, ("lumMod", "50000"));
        var bodyFill = SchemeFill(Drawing.SchemeColorValues.Dark1, ("lumOff", "25000"));
        var title = new P.TitleStyle(ListStyleLevel(1, 4400, fill: titleFill));
        var body = new P.BodyStyle(ListStyleLevel(1, 2000, fill: bodyFill));
        var path = CreatePresentation(masterTextStyles: new TextStyles(title, body, new P.OtherStyle()));
        using var document = PresentationDocument.Open(path, false);

        var profile = new BrandProfileExtractor().Extract(document);

        Assert.Equal("#808080", profile.Title!.Color); // white at 50% luminance
        Assert.Equal("#404040", profile.Body!.Color); // black + 25% luminance offset
    }

    [Fact]
    public void GetTextStyle_ExactLevelLookup_ReturnsNullOnMissWithoutFallback()
    {
        var title = new P.TitleStyle(ListStyleLevel(1, 4400));
        var body = new P.BodyStyle(ListStyleLevel(1, 2000), ListStyleLevel(2, 1600));
        var path = CreatePresentation(masterTextStyles: new TextStyles(title, body, new P.OtherStyle()));
        using var document = PresentationDocument.Open(path, false);
        var resolver = CreateResolver(document);

        Assert.Equal(20, resolver.GetTextStyle("Body", 0)!.FontSize);
        Assert.Equal(16, resolver.GetTextStyle("Body", 1)!.FontSize);
        Assert.Equal(44, resolver.GetTextStyle("Title", 0)!.FontSize);

        // No level-2 entry, no Other style, unknown key: exact lookup returns null
        // (unlike GetDefaultTextStyle which falls back to level 0 / Body).
        Assert.Null(resolver.GetTextStyle("Body", 2));
        Assert.Null(resolver.GetTextStyle("Other", 0));
        Assert.Null(resolver.GetTextStyle("nope", 0));
    }

    [Fact]
    public void SchemeColors_ExposesLoadedThemeColorMap()
    {
        var path = CreatePresentation();
        using var document = PresentationDocument.Open(path, false);
        var resolver = CreateResolver(document);

        IReadOnlyDictionary<string, string> colors = resolver.SchemeColors;

        Assert.Equal(12, colors.Count);
        Assert.Equal("#4F81BD", colors["accent1"]);
        Assert.Equal("#000000", colors["dk1"]);
    }

    [Fact]
    public void Extract_CustomClrMap_AddsRemappingWarning()
    {
        var remapped = new ColorMap { Background1 = Drawing.ColorSchemeIndexValues.Dark1, Text1 = Drawing.ColorSchemeIndexValues.Light1 };
        var remappedPath = CreatePresentation(colorMap: remapped);
        using (var document = PresentationDocument.Open(remappedPath, false))
        {
            var profile = new BrandProfileExtractor().Extract(document);

            var warning = Assert.Single(profile.Warnings);
            Assert.Contains("clrMap", warning);
            Assert.Contains("bg1->dk1", warning);
            Assert.Contains("tx1->lt1", warning);
        }

        var identityPath = CreatePresentation();
        using (var document = PresentationDocument.Open(identityPath, false))
        {
            Assert.Empty(new BrandProfileExtractor().Extract(document).Warnings);
        }
    }

    [Fact]
    public void Extract_SlideSize_ConvertsEmuToPoints()
    {
        var standardPath = CreatePresentation(slideWidthEmu: 9144000, slideHeightEmu: 6858000);
        using (var document = PresentationDocument.Open(standardPath, false))
        {
            var profile = new BrandProfileExtractor().Extract(document);
            Assert.Equal(720, profile.SlideWidthPt);
            Assert.Equal(540, profile.SlideHeightPt);
        }

        var widescreenPath = CreatePresentation(slideWidthEmu: 12192000, slideHeightEmu: 6858000);
        using (var document = PresentationDocument.Open(widescreenPath, false))
        {
            var profile = new BrandProfileExtractor().Extract(document);
            Assert.Equal(960, profile.SlideWidthPt);
            Assert.Equal(540, profile.SlideHeightPt);
        }
    }

    [Fact]
    public void Extract_MissingSlideSize_UsesPowerPointDefaultAndWarns()
    {
        var path = CreatePresentation(slideWidthEmu: null, slideHeightEmu: null);
        using var document = PresentationDocument.Open(path, false);

        var profile = new BrandProfileExtractor().Extract(document);

        Assert.Equal(720, profile.SlideWidthPt);
        Assert.Equal(540, profile.SlideHeightPt);
        var warning = Assert.Single(profile.Warnings);
        Assert.Contains("slide size", warning);
    }

    [Fact]
    public void Extract_NoSlides_ThrowsOfficeEditorException()
    {
        var path = CreatePresentation(includeSlide: false);
        using var document = PresentationDocument.Open(path, false);

        Assert.Throws<OfficeEditorException>(() => new BrandProfileExtractor().Extract(document));
    }

    [Fact]
    public void Extract_Background_ResolvesMasterSolidFillAndNullWhenAbsent()
    {
        var withBackgroundPath = CreatePresentation(masterBackground: new BackgroundProperties(
            new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = "102030" })));
        using (var document = PresentationDocument.Open(withBackgroundPath, false))
        {
            Assert.Equal("#102030", new BrandProfileExtractor().Extract(document).BackgroundColor);
        }

        var noBackgroundPath = CreatePresentation();
        using (var document = PresentationDocument.Open(noBackgroundPath, false))
        {
            Assert.Null(new BrandProfileExtractor().Extract(document).BackgroundColor);
        }
    }

    [Fact]
    public void Extract_SerializedShape_IsDeterministicWithSortedThemeColorKeys()
    {
        var path = CreatePresentation();
        using var document = PresentationDocument.Open(path, false);
        var profile = new BrandProfileExtractor().Extract(document);

        var first = BrandProfileSnapshotTests.Serialize(profile);
        var second = BrandProfileSnapshotTests.Serialize(profile);

        Assert.Equal(first, second);
        Assert.Contains("\"themeColors\"", first);
        // SortedDictionary ordering: accent1 sorts before dk1.
        Assert.True(first.IndexOf("\"accent1\"", StringComparison.Ordinal) < first.IndexOf("\"dk1\"", StringComparison.Ordinal));
    }

    private static StyleResolver CreateResolver(PresentationDocument document)
        => new(document, document.PresentationPart!.SlideParts.First());

    private string CreatePresentation(
        string? majorFont = null,
        string? minorFont = null,
        TextStyles? masterTextStyles = null,
        ColorMap? colorMap = null,
        BackgroundProperties? masterBackground = null,
        int? slideWidthEmu = 9144000,
        int? slideHeightEmu = 6858000,
        bool includeSlide = true)
    {
        var path = Path.Combine(_tempDir, $"brand-{Guid.NewGuid():N}.pptx");
        using var document = PresentationDocument.Create(path, PresentationDocumentType.Presentation);

        var presentationPart = document.AddPresentationPart();
        var presentation = new Presentation(new SlideMasterIdList(), new SlideIdList());
        if (slideWidthEmu.HasValue && slideHeightEmu.HasValue)
        {
            presentation.SlideSize = new SlideSize { Cx = slideWidthEmu.Value, Cy = slideHeightEmu.Value };
        }
        presentationPart.Presentation = presentation;

        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
        slideMasterPart.SlideMaster = new SlideMaster(
            CreateCommonSlideData(masterBackground),
            colorMap ?? new ColorMap(),
            new SlideLayoutIdList());
        if (masterTextStyles is not null)
        {
            slideMasterPart.SlideMaster.TextStyles = masterTextStyles;
        }

        var themePart = slideMasterPart.AddNewPart<ThemePart>();
        themePart.Theme = CreateOfficeTheme(majorFont, minorFont);

        var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
        slideLayoutPart.SlideLayout = new P.SlideLayout(CreateCommonSlideData(null));
        slideLayoutPart.AddPart(slideMasterPart);
        slideMasterPart.SlideMaster.SlideLayoutIdList!.Append(new SlideLayoutId
        {
            Id = 2147483649,
            RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart)
        });
        presentationPart.Presentation.SlideMasterIdList!.Append(new SlideMasterId
        {
            Id = 2147483648,
            RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
        });

        if (includeSlide)
        {
            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.Slide = new Slide(CreateCommonSlideData(null));
            slidePart.AddPart(slideLayoutPart);
            presentationPart.Presentation.SlideIdList!.Append(new SlideId
            {
                Id = 256,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });
        }

        return path;
    }

    private static CommonSlideData CreateCommonSlideData(BackgroundProperties? background)
        => background is null
            ? new CommonSlideData(CreateShapeTree())
            : new CommonSlideData(new Background(background), CreateShapeTree());

    private static ShapeTree CreateShapeTree()
        => new(
            new NonVisualGroupShapeProperties(
                new NonVisualDrawingProperties { Id = 0, Name = string.Empty },
                new NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(new Drawing.TransformGroup(
                new Drawing.Offset { X = 0, Y = 0 },
                new Drawing.Extents { Cx = 0, Cy = 0 },
                new Drawing.ChildOffset { X = 0, Y = 0 },
                new Drawing.ChildExtents { Cx = 0, Cy = 0 })));

    private static Drawing.Theme CreateOfficeTheme(string? majorFont, string? minorFont)
        => new(
            new Drawing.ThemeElements(
                new Drawing.ColorScheme(
                    new Drawing.Dark1Color(new Drawing.SystemColor { Val = Drawing.SystemColorValues.WindowText, LastColor = "000000" }),
                    new Drawing.Light1Color(new Drawing.SystemColor { Val = Drawing.SystemColorValues.Window, LastColor = "FFFFFF" }),
                    new Drawing.Dark2Color(new Drawing.RgbColorModelHex { Val = "1F497D" }),
                    new Drawing.Light2Color(new Drawing.RgbColorModelHex { Val = "EEECE1" }),
                    new Drawing.Accent1Color(new Drawing.RgbColorModelHex { Val = "4F81BD" }),
                    new Drawing.Accent2Color(new Drawing.RgbColorModelHex { Val = "C0504D" }),
                    new Drawing.Accent3Color(new Drawing.RgbColorModelHex { Val = "9BBB59" }),
                    new Drawing.Accent4Color(new Drawing.RgbColorModelHex { Val = "8064A2" }),
                    new Drawing.Accent5Color(new Drawing.RgbColorModelHex { Val = "4BACC6" }),
                    new Drawing.Accent6Color(new Drawing.RgbColorModelHex { Val = "F79646" }),
                    new Drawing.Hyperlink(new Drawing.RgbColorModelHex { Val = "0000FF" }),
                    new Drawing.FollowedHyperlinkColor(new Drawing.RgbColorModelHex { Val = "800080" })) { Name = "Office" },
                new Drawing.FontScheme(
                    new Drawing.MajorFont(new Drawing.LatinFont { Typeface = majorFont ?? "Calibri Light" }),
                    new Drawing.MinorFont(new Drawing.LatinFont { Typeface = minorFont ?? "Calibri" })) { Name = "Office" },
                new Drawing.FormatScheme(new Drawing.FillStyleList(), new Drawing.LineStyleList(), new Drawing.EffectStyleList(), new Drawing.BackgroundFillStyleList()) { Name = "Office" }),
            new Drawing.ObjectDefaults(),
            new Drawing.ExtraColorSchemeList()) { Name = "Office Theme" };

    private static OpenXmlUnknownElement ListStyleLevel(int level, int? fontSizeHundredths = null, bool? bold = null, string? font = null, Drawing.SolidFill? fill = null)
    {
        var lvl = new OpenXmlUnknownElement("a", $"lvl{level}pPr", DrawingNs);
        var rPr = new Drawing.DefaultRunProperties();
        if (fontSizeHundredths.HasValue) rPr.FontSize = fontSizeHundredths.Value;
        if (bold.HasValue) rPr.Bold = bold.Value;
        if (fill is not null) rPr.Append(fill);
        if (font is not null) rPr.Append(new Drawing.LatinFont { Typeface = font });
        lvl.Append(rPr);
        return lvl;
    }

    private static Drawing.SolidFill RgbFill(string rgb)
        => new(new Drawing.RgbColorModelHex { Val = rgb });

    private static Drawing.SolidFill SchemeFill(Drawing.SchemeColorValues scheme, params (string Name, string Value)[] transforms)
    {
        var schemeColor = new Drawing.SchemeColor { Val = scheme };
        foreach (var (name, value) in transforms)
        {
            schemeColor.Append(Unknown(name, ("val", value)));
        }
        return new Drawing.SolidFill(schemeColor);
    }

    private static OpenXmlUnknownElement Unknown(string localName, (string Name, string Value)? attribute = null)
    {
        var element = new OpenXmlUnknownElement("a", localName, DrawingNs);
        if (attribute.HasValue)
        {
            element.SetAttribute(new OpenXmlAttribute(string.Empty, attribute.Value.Name, string.Empty, attribute.Value.Value));
        }
        return element;
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDir)) Directory.Delete(_tempDir, true);
        }
        catch
        {
            // Best-effort cleanup.
        }
    }
}
