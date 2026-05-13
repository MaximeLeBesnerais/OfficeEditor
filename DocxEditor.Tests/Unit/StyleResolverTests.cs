using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PptxEditor.Core.Converters;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace DocxEditor.Tests.Unit;

public sealed class StyleResolverTests : IDisposable
{
    private const string DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private readonly string _tempDir = Path.Combine(Path.GetTempPath(), "StyleResolverTests", Guid.NewGuid().ToString("N"));

    public StyleResolverTests() => Directory.CreateDirectory(_tempDir);

    [Fact]
    public void ResolveBackgroundColor_PrefersSlideRgbOverLayoutAndMaster()
    {
        var path = CreatePresentation(
            slideBackground: RgbBackground("112233"),
            layoutBackground: SchemeBackground(Drawing.SchemeColorValues.Accent1),
            masterBackground: RgbBackground("AABBCC"));

        using var document = PresentationDocument.Open(path, false);
        var resolver = CreateResolver(document);

        Assert.Equal("#112233", resolver.ResolveBackgroundColor());
    }

    [Fact]
    public void ResolveBackgroundColor_FallsBackToLayoutSchemeColorThenMaster()
    {
        var path = CreatePresentation(
            slideBackground: null,
            layoutBackground: SchemeBackground(Drawing.SchemeColorValues.Accent2),
            masterBackground: RgbBackground("AABBCC"));

        using var document = PresentationDocument.Open(path, false);
        var resolver = CreateResolver(document);

        Assert.Equal("#C0504D", resolver.ResolveBackgroundColor());
    }

    [Fact]
    public void ResolveBackgroundColor_UsesMasterAndNullWhenNoSolidFillExists()
    {
        var withMasterPath = CreatePresentation(masterBackground: RgbBackground("ABCDEF"));
        using (var document = PresentationDocument.Open(withMasterPath, false))
        {
            Assert.Equal("#ABCDEF", CreateResolver(document).ResolveBackgroundColor());
        }

        var noFillPath = CreatePresentation(masterBackground: new BackgroundProperties(new Drawing.NoFill()));
        using (var document = PresentationDocument.Open(noFillPath, false))
        {
            Assert.Null(CreateResolver(document).ResolveBackgroundColor());
        }
    }

    [Theory]
    [InlineData("bg1", "#FFFFFF")]
    [InlineData("tx1", "#000000")]
    [InlineData("bg2", "#EEECE1")]
    [InlineData("tx2", "#1F497D")]
    [InlineData("accent1", "#4F81BD")]
    [InlineData("missing", null)]
    public void ResolveSchemeColor_HandlesAliasesSystemRgbAndMissingColors(string schemeName, string? expected)
    {
        var path = CreatePresentation();
        using var document = PresentationDocument.Open(path, false);
        var resolver = CreateResolver(document);

        Assert.Equal(expected, resolver.ResolveSchemeColor(schemeName));
    }

    [Fact]
    public void GetDefaultTextStyle_MapsPlaceholderTypesAndFallsBackToBody()
    {
        var path = CreatePresentation(masterTextStyles: CreateTextStyles());
        using var document = PresentationDocument.Open(path, false);
        var resolver = CreateResolver(document);

        var title = resolver.GetDefaultTextStyle(PlaceholderValues.CenteredTitle);
        Assert.Equal(44, title.FontSize);
        Assert.True(title.Bold);
        Assert.Equal("Aptos Display", title.FontFamily);

        var subtitle = resolver.GetDefaultTextStyle(PlaceholderValues.SubTitle);
        Assert.Equal(20, subtitle.FontSize);
        Assert.True(subtitle.Italic);

        var other = resolver.GetDefaultTextStyle(PlaceholderValues.Object);
        Assert.Equal(12, other.FontSize);
        Assert.True(other.Underline);

        Assert.Null(resolver.GetDefaultTextStyle(null).FontSize);
        Assert.Null(resolver.GetDefaultTextStyle(PlaceholderValues.Picture).FontSize);
    }

    [Fact]
    public void PlaceholderListStyle_ResolvesFormattingBulletsSpacingAndBodyPrViaIdx()
    {
        var layoutShape = PlaceholderShapeWithListStyle(10, PlaceholderValues.Body, CreateListStyleWithBranches(), anchor: Drawing.TextAnchoringTypeValues.Center);
        var path = CreatePresentation(layoutShapes: [layoutShape]);
        using var document = PresentationDocument.Open(path, false);
        var resolver = CreateResolver(document);

        var style = resolver.GetLayoutPlaceholderLstStyle(10, PlaceholderValues.Body, 0);
        Assert.Equal(18, style.FontSize);
        Assert.Equal("#123456", style.Color);
        Assert.Equal("Carlito", style.FontFamily);

        var bullet = resolver.GetLayoutPlaceholderBulletInfo(10, PlaceholderValues.Body, 0);
        Assert.True(bullet.HasBullet);
        Assert.Equal("→", bullet.BulletChar);

        var autoNumber = resolver.GetLayoutPlaceholderBulletInfo(10, PlaceholderValues.Body, 1);
        Assert.True(autoNumber.HasBullet);
        Assert.Equal("arabicPeriod", autoNumber.AutoNumberType);

        var bulletNone = resolver.GetLayoutPlaceholderBulletInfo(10, PlaceholderValues.Body, 2);
        Assert.True(bulletNone.HasBulletNone);

        Assert.Equal(1.2, resolver.GetLayoutPlaceholderLineSpacing(10, PlaceholderValues.Body, 0));
        Assert.Equal((6.0, 0.8), resolver.GetLayoutPlaceholderSpacing(10, PlaceholderValues.Body, 0));
        Assert.NotNull(resolver.GetLayoutPlaceholderBodyPr(10, PlaceholderValues.Body));
    }

    [Fact]
    public void MasterTextStyleBranches_ReturnBulletLineAndParagraphSpacingForSupportedPlaceholdersOnly()
    {
        var path = CreatePresentation(masterTextStyles: CreateTextStyles(includeParagraphBranches: true));
        using var document = PresentationDocument.Open(path, false);
        var resolver = CreateResolver(document);

        var bullet = resolver.GetMasterTxStyleBulletInfo(PlaceholderValues.Body, 0);
        Assert.True(bullet.HasBullet);
        Assert.Equal("•", bullet.BulletChar);
        Assert.Equal(14.5, resolver.GetMasterTxStyleLineSpacing(PlaceholderValues.Body, 0));
        Assert.Equal((2.4, 12.0), resolver.GetMasterTxStyleSpacing(PlaceholderValues.Body, 0));

        Assert.False(resolver.GetMasterTxStyleBulletInfo(null, 0).HasBullet);
        Assert.Null(resolver.GetMasterTxStyleLineSpacing(PlaceholderValues.Picture, 0));
        Assert.Equal((null, null), resolver.GetMasterTxStyleSpacing(PlaceholderValues.Picture, 0));
    }

    [Fact]
    public void MasterPlaceholderLookup_SkipsIdxWhenTypeDiffersThenFallsBackToType()
    {
        var mismatched = PlaceholderShapeWithListStyle(3, PlaceholderValues.DateAndTime, ListStyleWithFontSize(9));
        var bodyFallback = PlaceholderShapeWithListStyle(null, PlaceholderValues.Body, ListStyleWithFontSize(27));
        var path = CreatePresentation(masterShapes: [mismatched, bodyFallback]);
        using var document = PresentationDocument.Open(path, false);
        var resolver = CreateResolver(document);

        var style = resolver.GetMasterPlaceholderLstStyle(3, PlaceholderValues.Body, 0);

        Assert.Equal(27, style.FontSize);
    }

    [Fact]
    public void ResolveBackgroundColor_IgnoresUnsupportedAndInvalidFillsBeforeMasterFallback()
    {
        var path = CreatePresentation(
            slideBackground: new BackgroundProperties(new Drawing.GradientFill()),
            layoutBackground: RawSchemeBackground("notInTheme"),
            masterBackground: new BackgroundProperties(new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = "ABC123" })));
        using var document = PresentationDocument.Open(path, false);

        Assert.Equal("#ABC123", CreateResolver(document).ResolveBackgroundColor());
    }

    [Fact]
    public void ResolveBackgroundColor_ParsesSchemeValueFromOuterXmlWhenEnumIsMissing()
    {
        var path = CreatePresentation(slideBackground: RawSchemeBackground("accent3"));
        using var document = PresentationDocument.Open(path, false);

        Assert.Equal("#9BBB59", CreateResolver(document).ResolveBackgroundColor());
    }

    [Fact]
    public void ResolveSchemeColor_ReturnsNullWhenThemeOrThemeColorsAreMissing()
    {
        var noThemePath = CreatePresentation(includeTheme: false);
        using (var document = PresentationDocument.Open(noThemePath, false))
        {
            Assert.Null(CreateResolver(document).ResolveSchemeColor("accent1"));
        }

        var sparseThemePath = CreatePresentation(theme: new Drawing.Theme(
            new Drawing.ThemeElements(
                new Drawing.ColorScheme(new Drawing.Dark1Color(new Drawing.SchemeColor { Val = Drawing.SchemeColorValues.Accent1 })) { Name = "Sparse" },
                new Drawing.FontScheme(new Drawing.MajorFont(), new Drawing.MinorFont()) { Name = "Fonts" },
                new Drawing.FormatScheme(new Drawing.FillStyleList(), new Drawing.LineStyleList(), new Drawing.EffectStyleList(), new Drawing.BackgroundFillStyleList()) { Name = "Fmt" })) { Name = "Sparse" });
        using (var document = PresentationDocument.Open(sparseThemePath, false))
        {
            Assert.Null(CreateResolver(document).ResolveSchemeColor("dk1"));
            Assert.Null(CreateResolver(document).ResolveSchemeColor("accent1"));
        }
    }

    [Fact]
    public void MasterTextStyles_SkipMalformedLevelsAndPreserveFalseAndUnsetRunProperties()
    {
        var body = new P.BodyStyle(
            Unknown("defPPr"),
            Unknown("lvlXpPr", child: new Drawing.DefaultRunProperties { FontSize = 9900 }),
            Unknown("lvl2pPr", child: new Drawing.DefaultRunProperties { FontSize = 3300 }),
            Unknown("lvl1pPr", child: new Drawing.DefaultRunProperties
            {
                Bold = false,
                Italic = false,
                Underline = Drawing.TextUnderlineValues.None
            }));
        var path = CreatePresentation(masterTextStyles: new TextStyles(new P.TitleStyle(), body, new P.OtherStyle()));
        using var document = PresentationDocument.Open(path, false);
        var resolver = CreateResolver(document);

        var style = resolver.GetDefaultTextStyle(PlaceholderValues.Body);

        Assert.Null(style.FontSize);
        Assert.False(style.Bold);
        Assert.False(style.Italic);
        Assert.False(style.Underline);
        Assert.Null(style.Color);
        Assert.Null(style.FontFamily);
    }

    [Fact]
    public void PlaceholderLookup_HandlesMissingTextBodiesListStylesLevelsAndNullBodyProperties()
    {
        var noTextBody = PlaceholderShapeWithoutTextBody(1, PlaceholderValues.Body);
        var noListStyle = PlaceholderShapeWithTextBody(2, PlaceholderValues.Body, new TextBody(new Drawing.BodyProperties()));
        var noLevel = PlaceholderShapeWithTextBody(3, PlaceholderValues.Body, new TextBody(new Drawing.BodyProperties(), new Drawing.ListStyle(ListStyleLevel(2, 21))));
        var noRunProperties = PlaceholderShapeWithTextBody(4, PlaceholderValues.Body, new TextBody(new Drawing.BodyProperties(), new Drawing.ListStyle(Unknown("lvl1pPr"))));
        var path = CreatePresentation(layoutShapes: [noTextBody, noListStyle, noLevel, noRunProperties]);
        using var document = PresentationDocument.Open(path, false);
        var resolver = CreateResolver(document);

        Assert.Null(resolver.GetLayoutPlaceholderLstStyle(1, PlaceholderValues.Body, 0).FontSize);
        Assert.False(resolver.GetLayoutPlaceholderBulletInfo(2, PlaceholderValues.Body, 0).HasBullet);
        Assert.Null(resolver.GetLayoutPlaceholderLineSpacing(3, PlaceholderValues.Body, 0));
        Assert.Equal((null, null), resolver.GetLayoutPlaceholderSpacing(4, PlaceholderValues.Body, 0));
        Assert.Null(resolver.GetLayoutPlaceholderBodyPr(1, PlaceholderValues.Body));
        Assert.Null(resolver.GetLayoutPlaceholderBodyPr(99, PlaceholderValues.Body));
    }

    [Fact]
    public void PlaceholderLookup_UsesIdxWithoutTypeAndSkipsShapesWithoutRecognizedPlaceholderType()
    {
        var directPlaceholder = PlaceholderShapeWithListStyle(7, PlaceholderValues.DateAndTime, ListStyleWithFontSize(23));
        var unsupportedPlaceholder = PlaceholderShapeWithRawPlaceholder("media", 8, ListStyleWithFontSize(31));
        var typeFallback = PlaceholderShapeWithListStyle(null, PlaceholderValues.Body, ListStyleWithFontSize(19));
        var path = CreatePresentation(layoutShapes: [unsupportedPlaceholder, directPlaceholder, typeFallback]);
        using var document = PresentationDocument.Open(path, false);
        var resolver = CreateResolver(document);

        Assert.Equal(23, resolver.GetLayoutPlaceholderLstStyle(7, null, 0).FontSize);
        Assert.Equal(19, resolver.GetLayoutPlaceholderLstStyle(null, PlaceholderValues.Body, 0).FontSize);
        Assert.Equal(31, resolver.GetLayoutPlaceholderLstStyle(8, PlaceholderValues.Picture, 0).FontSize);
    }

    [Fact]
    public void LineAndParagraphSpacing_ReturnNullForInvalidValuesAndPreferPercentageWhenBothExist()
    {
        var invalid = Unknown("lvl1pPr");
        invalid.Append(Unknown("lnSpc", child: Unknown("spcPts", ("val", "bad"))));
        invalid.Append(Unknown("spcBef", child: Unknown("spcPts", ("val", "alsoBad"))));

        var both = Unknown("lvl2pPr");
        var line = Unknown("lnSpc");
        line.Append(Unknown("spcPts", ("val", "900")), Unknown("spcPct", ("val", "150000")));
        var before = Unknown("spcBef");
        before.Append(Unknown("spcPts", ("val", "100")), Unknown("spcPct", ("val", "250000")));
        both.Append(line, before, Unknown("spcAft", child: Unknown("spcPct", ("val", "50000"))));

        var path = CreatePresentation(layoutShapes: [PlaceholderShapeWithTextBody(1, PlaceholderValues.Body, new TextBody(new Drawing.BodyProperties(), new Drawing.ListStyle(invalid, both)))]);
        using var document = PresentationDocument.Open(path, false);
        var resolver = CreateResolver(document);

        Assert.Null(resolver.GetLayoutPlaceholderLineSpacing(1, PlaceholderValues.Body, 0));
        Assert.Equal((null, null), resolver.GetLayoutPlaceholderSpacing(1, PlaceholderValues.Body, 0));
        Assert.Equal(9, resolver.GetLayoutPlaceholderLineSpacing(1, PlaceholderValues.Body, 1));
        Assert.Equal((2.5, 0.5), resolver.GetLayoutPlaceholderSpacing(1, PlaceholderValues.Body, 1));
    }

    [Fact]
    public void MasterPlaceholderApis_ReturnMatchedStylesAndDefaultWhenNoMasterShapeMatches()
    {
        var master = PlaceholderShapeWithListStyle(5, PlaceholderValues.Object, CreateListStyleWithBranches(), Drawing.TextAnchoringTypeValues.Bottom);
        var path = CreatePresentation(masterShapes: [master]);
        using var document = PresentationDocument.Open(path, false);
        var resolver = CreateResolver(document);

        Assert.Equal(18, resolver.GetMasterPlaceholderLstStyle(5, PlaceholderValues.Object, 0).FontSize);
        Assert.Equal("→", resolver.GetMasterPlaceholderBulletInfo(5, PlaceholderValues.Object, 0).BulletChar);
        Assert.Equal(1.2, resolver.GetMasterPlaceholderLineSpacing(5, PlaceholderValues.Object, 0));
        Assert.Equal((6.0, 0.8), resolver.GetMasterPlaceholderSpacing(5, PlaceholderValues.Object, 0));
        Assert.NotNull(resolver.GetMasterPlaceholderBodyPr(5, PlaceholderValues.Object));
        Assert.Null(resolver.GetMasterPlaceholderLstStyle(6, PlaceholderValues.Picture, 0).FontSize);
    }

    [Fact]
    public void TextStyleCache_ReflectionCoversHitLevelFallbackBodyFallbackAndEmptyMiss()
    {
        var cacheType = typeof(StyleResolver).GetNestedType("TextStyleCache", System.Reflection.BindingFlags.NonPublic)!;
        var cache = Activator.CreateInstance(cacheType)!;
        var addStyle = cacheType.GetMethod("AddStyle")!;
        var getStyle = cacheType.GetMethod("GetStyle")!;

        var body = new StyleResolver.DefaultTextStyle { FontSize = 11 };
        var titleLevelTwo = new StyleResolver.DefaultTextStyle { FontSize = 22 };
        addStyle.Invoke(cache, ["Body", 0, body]);
        addStyle.Invoke(cache, ["Title", 2, titleLevelTwo]);

        Assert.Same(titleLevelTwo, getStyle.Invoke(cache, ["Title", 2]));
        Assert.Same(body, getStyle.Invoke(cache, ["Title", 1]));
        Assert.Same(body, getStyle.Invoke(cache, ["Other", 0]));

        var emptyCache = Activator.CreateInstance(cacheType)!;
        var missing = Assert.IsType<StyleResolver.DefaultTextStyle>(getStyle.Invoke(emptyCache, ["Body", 0]));
        Assert.Null(missing.FontSize);
    }

    [Fact]
    public void PlaceholderTypeMapping_CoversSupportedSdkPlaceholderValues()
    {
        var shapes = new[]
        {
            PlaceholderShapeWithListStyle(11, PlaceholderValues.Title, ListStyleWithFontSize(11)),
            PlaceholderShapeWithListStyle(12, PlaceholderValues.CenteredTitle, ListStyleWithFontSize(12)),
            PlaceholderShapeWithListStyle(13, PlaceholderValues.SubTitle, ListStyleWithFontSize(13)),
            PlaceholderShapeWithListStyle(14, PlaceholderValues.Picture, ListStyleWithFontSize(14)),
            PlaceholderShapeWithListStyle(15, PlaceholderValues.Chart, ListStyleWithFontSize(15)),
            PlaceholderShapeWithListStyle(16, PlaceholderValues.Table, ListStyleWithFontSize(16)),
            PlaceholderShapeWithListStyle(17, PlaceholderValues.SlideNumber, ListStyleWithFontSize(17)),
            PlaceholderShapeWithListStyle(18, PlaceholderValues.Footer, ListStyleWithFontSize(18)),
            PlaceholderShapeWithListStyle(19, PlaceholderValues.Header, ListStyleWithFontSize(19)),
            PlaceholderShapeWithListStyle(20, PlaceholderValues.Object, ListStyleWithFontSize(20)),
            PlaceholderShapeWithListStyle(21, PlaceholderValues.DateAndTime, ListStyleWithFontSize(21))
        };
        var path = CreatePresentation(layoutShapes: shapes);
        using var document = PresentationDocument.Open(path, false);
        var resolver = CreateResolver(document);

        Assert.Equal(11, resolver.GetLayoutPlaceholderLstStyle(11, PlaceholderValues.Title, 0).FontSize);
        Assert.Equal(12, resolver.GetLayoutPlaceholderLstStyle(12, PlaceholderValues.CenteredTitle, 0).FontSize);
        Assert.Equal(13, resolver.GetLayoutPlaceholderLstStyle(13, PlaceholderValues.SubTitle, 0).FontSize);
        Assert.Equal(14, resolver.GetLayoutPlaceholderLstStyle(14, PlaceholderValues.Picture, 0).FontSize);
        Assert.Equal(15, resolver.GetLayoutPlaceholderLstStyle(15, PlaceholderValues.Chart, 0).FontSize);
        Assert.Equal(16, resolver.GetLayoutPlaceholderLstStyle(16, PlaceholderValues.Table, 0).FontSize);
        Assert.Equal(17, resolver.GetLayoutPlaceholderLstStyle(17, PlaceholderValues.SlideNumber, 0).FontSize);
        Assert.Equal(18, resolver.GetLayoutPlaceholderLstStyle(18, PlaceholderValues.Footer, 0).FontSize);
        Assert.Equal(19, resolver.GetLayoutPlaceholderLstStyle(19, PlaceholderValues.Header, 0).FontSize);
        Assert.Equal(20, resolver.GetLayoutPlaceholderLstStyle(20, PlaceholderValues.Object, 0).FontSize);
        Assert.Equal(21, resolver.GetLayoutPlaceholderLstStyle(21, PlaceholderValues.DateAndTime, 0).FontSize);
    }

    [Fact]
    public void MasterTextStyleApis_MapTitleSubtitleAndObjectBranchesWithoutOptionalParagraphProperties()
    {
        var path = CreatePresentation(masterTextStyles: CreateTextStyles());
        using var document = PresentationDocument.Open(path, false);
        var resolver = CreateResolver(document);

        Assert.False(resolver.GetMasterTxStyleBulletInfo(PlaceholderValues.Title, 0).HasBullet);
        Assert.False(resolver.GetMasterTxStyleBulletInfo(PlaceholderValues.SubTitle, 0).HasBullet);
        Assert.False(resolver.GetMasterTxStyleBulletInfo(PlaceholderValues.Object, 0).HasBullet);
        Assert.False(resolver.GetMasterTxStyleBulletInfo(PlaceholderValues.Body, 4).HasBullet);

        Assert.Null(resolver.GetMasterTxStyleLineSpacing(PlaceholderValues.CenteredTitle, 0));
        Assert.Null(resolver.GetMasterTxStyleLineSpacing(PlaceholderValues.SubTitle, 0));
        Assert.Null(resolver.GetMasterTxStyleLineSpacing(PlaceholderValues.Object, 0));
        Assert.Null(resolver.GetMasterTxStyleLineSpacing(PlaceholderValues.Body, 4));

        Assert.Equal((null, null), resolver.GetMasterTxStyleSpacing(PlaceholderValues.CenteredTitle, 0));
        Assert.Equal((null, null), resolver.GetMasterTxStyleSpacing(PlaceholderValues.SubTitle, 0));
        Assert.Equal((null, null), resolver.GetMasterTxStyleSpacing(PlaceholderValues.Object, 0));
        Assert.Equal((null, null), resolver.GetMasterTxStyleSpacing(PlaceholderValues.Body, 4));
    }

    private string CreatePresentation(
        BackgroundProperties? slideBackground = null,
        BackgroundProperties? layoutBackground = null,
        BackgroundProperties? masterBackground = null,
        TextStyles? masterTextStyles = null,
        IReadOnlyCollection<OpenXmlElement>? layoutShapes = null,
        IReadOnlyCollection<OpenXmlElement>? masterShapes = null,
        bool includeTheme = true,
        Drawing.Theme? theme = null)
    {
        var path = Path.Combine(_tempDir, $"style-{Guid.NewGuid():N}.pptx");
        using var document = PresentationDocument.Create(path, PresentationDocumentType.Presentation);

        var presentationPart = document.AddPresentationPart();
        presentationPart.Presentation = new Presentation(new SlideMasterIdList(), new SlideIdList())
        {
            SlideSize = new SlideSize { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 }
        };

        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
        slideMasterPart.SlideMaster = new SlideMaster(
            CreateCommonSlideData(masterShapes, masterBackground),
            new ColorMap(),
            new SlideLayoutIdList());
        if (masterTextStyles is not null)
        {
            slideMasterPart.SlideMaster.TextStyles = masterTextStyles;
        }

        if (includeTheme)
        {
            var themePart = slideMasterPart.AddNewPart<ThemePart>();
            themePart.Theme = theme ?? CreateTheme();
        }

        var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
        slideLayoutPart.SlideLayout = new P.SlideLayout(CreateCommonSlideData(layoutShapes, layoutBackground));
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

        var slidePart = presentationPart.AddNewPart<SlidePart>();
        slidePart.Slide = new Slide(CreateCommonSlideData(null, slideBackground));
        slidePart.AddPart(slideLayoutPart);
        presentationPart.Presentation.SlideIdList!.Append(new SlideId
        {
            Id = 256,
            RelationshipId = presentationPart.GetIdOfPart(slidePart)
        });

        return path;
    }

    private static StyleResolver CreateResolver(PresentationDocument document)
        => new(document, document.PresentationPart!.SlideParts.First());

    private static CommonSlideData CreateCommonSlideData(IReadOnlyCollection<OpenXmlElement>? shapes, BackgroundProperties? background)
        => background is null
            ? new CommonSlideData(CreateShapeTree(shapes))
            : new CommonSlideData(new Background(background), CreateShapeTree(shapes));

    private static ShapeTree CreateShapeTree(IReadOnlyCollection<OpenXmlElement>? shapes = null)
    {
        var tree = new ShapeTree(
            new NonVisualGroupShapeProperties(
                new NonVisualDrawingProperties { Id = 0, Name = string.Empty },
                new NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(new Drawing.TransformGroup(
                new Drawing.Offset { X = 0, Y = 0 },
                new Drawing.Extents { Cx = 0, Cy = 0 },
                new Drawing.ChildOffset { X = 0, Y = 0 },
                new Drawing.ChildExtents { Cx = 0, Cy = 0 })));

        if (shapes is not null)
        {
            foreach (var shape in shapes)
            {
                tree.Append(shape);
            }
        }

        return tree;
    }

    private static BackgroundProperties RgbBackground(string rgb)
        => new(new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = rgb }));

    private static BackgroundProperties SchemeBackground(Drawing.SchemeColorValues scheme)
        => new(new Drawing.SolidFill(new Drawing.SchemeColor { Val = scheme }));

    private static BackgroundProperties RawSchemeBackground(string value)
    {
        var solidFill = new Drawing.SolidFill();
        var schemeColor = new OpenXmlUnknownElement("a", "schemeClr", DrawingNs);
        schemeColor.SetAttribute(new OpenXmlAttribute(string.Empty, "val", string.Empty, value));
        solidFill.Append(schemeColor);
        return new BackgroundProperties(solidFill);
    }

    private static Drawing.Theme CreateTheme()
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
                new Drawing.FontScheme(new Drawing.MajorFont(), new Drawing.MinorFont()) { Name = "Office" },
                new Drawing.FormatScheme(new Drawing.FillStyleList(), new Drawing.LineStyleList(), new Drawing.EffectStyleList(), new Drawing.BackgroundFillStyleList()) { Name = "Office" }),
            new Drawing.ObjectDefaults(),
            new Drawing.ExtraColorSchemeList()) { Name = "Office Theme" };

    private static TextStyles CreateTextStyles(bool includeParagraphBranches = false)
    {
        var title = new P.TitleStyle(ListStyleLevel(1, 44, bold: true, font: "Aptos Display"));
        var body = new P.BodyStyle(includeParagraphBranches ? BodyLevelWithBranches() : ListStyleLevel(1, 20, italic: true));
        var other = new P.OtherStyle(ListStyleLevel(1, 12, underline: true));
        return new TextStyles(title, body, other);
    }

    private static OpenXmlUnknownElement BodyLevelWithBranches()
    {
        var lvl = ListStyleLevel(1, 20, italic: true);
        lvl.Append(Unknown("buChar", ("char", "•")));
        lvl.Append(Unknown("lnSpc", child: Unknown("spcPts", ("val", "1450"))));
        lvl.Append(Unknown("spcBef", child: Unknown("spcPct", ("val", "240000"))));
        lvl.Append(Unknown("spcAft", child: Unknown("spcPts", ("val", "1200"))));
        return lvl;
    }

    private static Drawing.ListStyle CreateListStyleWithBranches()
    {
        var lvl1 = ListStyleLevel(1, 18, color: "123456", font: "Carlito");
        lvl1.Append(Unknown("buChar", ("char", "→")));
        lvl1.Append(Unknown("lnSpc", child: Unknown("spcPct", ("val", "120000"))));
        lvl1.Append(Unknown("spcBef", child: Unknown("spcPts", ("val", "600"))));
        lvl1.Append(Unknown("spcAft", child: Unknown("spcPct", ("val", "80000"))));

        var lvl2 = ListStyleLevel(2, 16);
        lvl2.Append(Unknown("buAutoNum", ("type", "arabicPeriod")));

        var lvl3 = ListStyleLevel(3, 14);
        lvl3.Append(Unknown("buNone"));

        return new Drawing.ListStyle(lvl1, lvl2, lvl3);
    }

    private static Drawing.ListStyle ListStyleWithFontSize(int fontSize)
        => new(ListStyleLevel(1, fontSize));

    private static OpenXmlUnknownElement ListStyleLevel(int level, int fontSize, bool? bold = null, bool? italic = null, bool underline = false, string? color = null, string? font = null)
    {
        var lvl = new OpenXmlUnknownElement("a", $"lvl{level}pPr", DrawingNs);
        var rPr = new Drawing.DefaultRunProperties { FontSize = fontSize * 100 };
        if (bold.HasValue) rPr.Bold = bold.Value;
        if (italic.HasValue) rPr.Italic = italic.Value;
        if (underline) rPr.Underline = Drawing.TextUnderlineValues.Single;
        if (color is not null) rPr.Append(new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = color }));
        if (font is not null) rPr.Append(new Drawing.LatinFont { Typeface = font });
        lvl.Append(rPr);
        return lvl;
    }

    private static P.Shape PlaceholderShapeWithListStyle(int? idx, PlaceholderValues type, Drawing.ListStyle listStyle, Drawing.TextAnchoringTypeValues? anchor = null)
    {
        var placeholder = new PlaceholderShape { Type = type };
        if (idx.HasValue) placeholder.Index = (uint)idx.Value;

        return new P.Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = (uint)(idx ?? 99), Name = type.ToString() },
                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(placeholder)),
            new ShapeProperties(),
            new TextBody(new Drawing.BodyProperties { Anchor = anchor }, listStyle));
    }

    private static P.Shape PlaceholderShapeWithTextBody(int? idx, PlaceholderValues type, TextBody textBody)
    {
        var placeholder = new PlaceholderShape { Type = type };
        if (idx.HasValue) placeholder.Index = (uint)idx.Value;

        return new P.Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = (uint)(idx ?? 99), Name = type.ToString() },
                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(placeholder)),
            new ShapeProperties(),
            textBody);
    }

    private static P.Shape PlaceholderShapeWithoutTextBody(int? idx, PlaceholderValues type)
    {
        var placeholder = new PlaceholderShape { Type = type };
        if (idx.HasValue) placeholder.Index = (uint)idx.Value;

        return new P.Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = (uint)(idx ?? 99), Name = type.ToString() },
                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(placeholder)),
            new ShapeProperties());
    }

    private static P.Shape DirectPlaceholderShape(int? idx, PlaceholderValues type, Drawing.ListStyle listStyle)
    {
        var placeholder = new PlaceholderShape { Type = type };
        if (idx.HasValue) placeholder.Index = (uint)idx.Value;

        return new P.Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = (uint)(idx ?? 99), Name = type.ToString() },
                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks { NoGrouping = true }),
                placeholder),
            new ShapeProperties(),
            new TextBody(new Drawing.BodyProperties(), listStyle));
    }

    private static P.Shape PlaceholderShapeWithRawPlaceholder(string rawType, int idx, Drawing.ListStyle listStyle)
    {
        var placeholder = new OpenXmlUnknownElement("p", "ph", "http://schemas.openxmlformats.org/presentationml/2006/main");
        placeholder.SetAttribute(new OpenXmlAttribute(string.Empty, "type", string.Empty, rawType));
        placeholder.SetAttribute(new OpenXmlAttribute(string.Empty, "idx", string.Empty, idx.ToString()));

        return new P.Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = (uint)idx, Name = rawType },
                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(placeholder)),
            new ShapeProperties(),
            new TextBody(new Drawing.BodyProperties(), listStyle));
    }

    private static OpenXmlUnknownElement Unknown(string localName, (string Name, string Value)? attribute = null, OpenXmlElement? child = null)
    {
        var element = new OpenXmlUnknownElement("a", localName, DrawingNs);
        if (attribute.HasValue)
        {
            element.SetAttribute(new OpenXmlAttribute(string.Empty, attribute.Value.Name, string.Empty, attribute.Value.Value));
        }

        if (child is not null)
        {
            element.Append(child);
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
