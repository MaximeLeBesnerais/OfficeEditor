using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PptxEditor.Core.Converters;
using Xunit;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Synthetic-XML tests for text/paragraph/list fidelity in the PPTX→Typst converter:
/// bullet glyph color (a:buClr / a:buClrTx), per-paragraph line spacing (a:lnSpc),
/// and per-level list indentation/marker resolution (marL/indent/lvl + buChar).
/// </summary>
public sealed class PptxToTypstConverterTextListTests : IDisposable
{
    private const long EmusPerPoint = 12700;

    private readonly string _tempDir;

    public PptxToTypstConverterTextListTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "PptxToTypstTextListTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, true);
            }
        }
        catch
        {
            // Best effort cleanup
        }
    }

    [Fact]
    public void GenerateTypstSource_ExplicitBulletColor_AppliesColorToMarker()
    {
        var shape = TextShape(2, "Cyan bullet", pPr: new Drawing.ParagraphProperties(
            new Drawing.BulletColor(new Drawing.RgbColorModelHex { Val = "22D3EE" }),
            new Drawing.CharacterBullet { Char = "•" }));
        var path = CreatePptx("bullet-color.pptx", customizeMaster: null, shape);

        var source = ConvertToTypstSource(path);

        Assert.Contains("#list(marker: [#text(fill: rgb(\"#22D3EE\"))[•]]", source);
    }

    [Fact]
    public void GenerateTypstSource_BulletColorFollowsText_UsesFirstRunColor()
    {
        var shape = TextShape(2, "Red bullet", pPr: new Drawing.ParagraphProperties(
                new Drawing.BulletColorText(),
                new Drawing.CharacterBullet { Char = "•" }),
            runProperties: new Drawing.RunProperties(
                new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = "FF0000" }))
            {
                FontSize = new Int32Value(1250)
            });
        var path = CreatePptx("bullet-color-text.pptx", customizeMaster: null, shape);

        var source = ConvertToTypstSource(path);

        Assert.Contains("#list(marker: [#text(fill: rgb(\"#FF0000\"))[•]]", source);
    }

    [Fact]
    public void GenerateTypstSource_LineSpacingPercent_EmitsPerParagraphLeading()
    {
        // Two paragraphs with different spcPct values: the second leading can only come
        // from per-paragraph emission (element-level fallback only knows the first).
        var shape = TextShape(2,
            new Drawing.Paragraph(
                new Drawing.ParagraphProperties(
                    new Drawing.LineSpacing(new Drawing.SpacingPercent { Val = 125000 })),
                new Drawing.Run(
                    new Drawing.RunProperties { FontSize = new Int32Value(1100) },
                    new Drawing.Text { Text = "First body line" })),
            new Drawing.Paragraph(
                new Drawing.ParagraphProperties(
                    new Drawing.LineSpacing(new Drawing.SpacingPercent { Val = 150000 })),
                new Drawing.Run(
                    new Drawing.RunProperties { FontSize = new Int32Value(1100) },
                    new Drawing.Text { Text = "Second body line" })));
        var path = CreatePptx("line-spacing-pct.pptx", customizeMaster: null, shape);

        var source = ConvertToTypstSource(path);

        // 11pt * 1.25 - 11pt * 0.65 = 6.60pt ; 11pt * 1.50 - 11pt * 0.65 = 9.35pt
        Assert.Contains("#set par(leading: 6.60pt)", source);
        Assert.Contains("#set par(leading: 9.35pt)", source);
    }

    [Fact]
    public void GenerateTypstSource_LineSpacingPoints_EmitsAbsoluteLeading()
    {
        var shape = TextShape(2,
            new Drawing.Paragraph(
                new Drawing.ParagraphProperties(
                    new Drawing.LineSpacing(new Drawing.SpacingPoints { Val = 2000 })),
                new Drawing.Run(
                    new Drawing.RunProperties { FontSize = new Int32Value(1100) },
                    new Drawing.Text { Text = "Absolute spaced line" })));
        var path = CreatePptx("line-spacing-pts.pptx", customizeMaster: null, shape);

        var source = ConvertToTypstSource(path);

        // spcPts 2000 = 20pt target line pitch; 20 - 11pt * 0.65 = 12.85pt
        Assert.Contains("#set par(leading: 12.85pt)", source);
    }

    [Fact]
    public void GenerateTypstSource_LineSpacingPercent_UsesRunFontSizeNotElementDefault()
    {
        // Mixed-formatting runs: the paragraph default falls back to 18pt, but leading
        // must be computed from the first run's 13pt (old behavior emitted 8.46pt).
        var shape = TextShape(2,
            new Drawing.Paragraph(
                new Drawing.ParagraphProperties(
                    new Drawing.LineSpacing(new Drawing.SpacingPercent { Val = 112000 })),
                new Drawing.Run(
                    new Drawing.RunProperties { FontSize = new Int32Value(1300), Bold = new BooleanValue(true) },
                    new Drawing.Text { Text = "Bold lead" })),
            new Drawing.Paragraph(
                new Drawing.ParagraphProperties(
                    new Drawing.LineSpacing(new Drawing.SpacingPercent { Val = 112000 })),
                new Drawing.Run(
                    new Drawing.RunProperties { FontSize = new Int32Value(1150) },
                    new Drawing.Text { Text = "Regular follow" })));
        var path = CreatePptx("line-spacing-mixed.pptx", customizeMaster: null, shape);

        var source = ConvertToTypstSource(path);

        // 13pt * 1.12 - 13pt * 0.65 = 6.11pt ; 11.5pt * 1.12 - 11.5pt * 0.65 = 5.41pt (rounded)
        Assert.Contains("#set par(leading: 6.11pt)", source);
        Assert.Contains("#set par(leading: 5.41pt)", source);
        Assert.DoesNotContain("#set par(leading: 8.46pt)", source);
    }

    [Fact]
    public void GenerateTypstSource_MixedListLevels_EmitsPerLevelMarkersAndIndents()
    {
        // Level signaled by marL/indent (+lvl) with a different buChar per level:
        // lvl0 "•" at marL 190500, lvl1 "–" at marL 381000, lvl2 "▪" at marL 571500,
        // all with indent -190500 (hanging 15pt).
        var shape = TextShape(2,
            BulletParagraph("Top level item", level: 0, marL: 190500, indent: -190500, bulletChar: "•"),
            BulletParagraph("Second level item", level: 1, marL: 381000, indent: -190500, bulletChar: "–"),
            BulletParagraph("Third level item", level: 2, marL: 571500, indent: -190500, bulletChar: "▪"));
        var path = CreatePptx("mixed-list-levels.pptx", customizeMaster: null, shape);

        var source = ConvertToTypstSource(path);

        // lvl0: marker at 15-15=0pt (indent omitted), body 15pt right of marker
        Assert.Contains("#list(marker: [•], body-indent: 15.00pt)", source);
        // lvl1: marker at 30-15=15pt, body at 30pt
        Assert.Contains("#list(marker: [–], indent: 15.00pt, body-indent: 15.00pt)", source);
        // lvl2: marker at 45-15=30pt, body at 45pt
        Assert.Contains("#list(marker: [▪], indent: 30.00pt, body-indent: 15.00pt)", source);
    }

    [Fact]
    public void GenerateTypstSource_PlaceholderListIndents_InheritFromMasterTxStyles()
    {
        // Body placeholder paragraph with lvl=1 and no local marL/indent/buChar:
        // everything must come from the master txStyles BodyStyle lvl2pPr.
        var shape = TextShape(2,
            new Drawing.Paragraph[]
            {
                new Drawing.Paragraph(
                    new Drawing.ParagraphProperties { Level = new Int32Value(1) },
                    new Drawing.Run(
                        new Drawing.RunProperties { FontSize = new Int32Value(1200) },
                        new Drawing.Text { Text = "Inherited level two" }))
            },
            placeholderType: PlaceholderValues.Body);

        void CustomizeMaster(SlideMaster master)
        {
            master.TextStyles = new TextStyles(
                new TitleStyle(),
                new BodyStyle(
                    new Drawing.Level1ParagraphProperties(new Drawing.CharacterBullet { Char = "•" })
                    {
                        LeftMargin = new Int32Value(342900),
                        Indent = new Int32Value(-342900)
                    },
                    new Drawing.Level2ParagraphProperties(new Drawing.CharacterBullet { Char = "–" })
                    {
                        LeftMargin = new Int32Value(742950),
                        Indent = new Int32Value(-285750)
                    }),
                new OtherStyle());
        }

        var path = CreatePptx("placeholder-indent-inherit.pptx", CustomizeMaster, shape);

        var source = ConvertToTypstSource(path);

        // lvl2pPr: marL 742950 EMU = 58.5pt, indent -285750 EMU = -22.5pt
        // -> marker at 36pt, body 22.5pt right of marker
        Assert.Contains("#list(marker: [–], indent: 36.00pt, body-indent: 22.50pt)", source);
    }

    private static Drawing.Paragraph BulletParagraph(string text, int level, int marL, int indent, string bulletChar)
    {
        var pPr = new Drawing.ParagraphProperties(new Drawing.CharacterBullet { Char = bulletChar })
        {
            LeftMargin = new Int32Value(marL),
            Indent = new Int32Value(indent)
        };
        if (level > 0)
        {
            pPr.Level = new Int32Value(level);
        }

        return new Drawing.Paragraph(
            pPr,
            new Drawing.Run(
                new Drawing.RunProperties { FontSize = new Int32Value(1300) },
                new Drawing.Text { Text = text }));
    }

    private string ConvertToTypstSource(string path)
    {
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);
        var presentation = converter.Convert();
        return converter.GenerateTypstSource(presentation);
    }

    private static P.Shape TextShape(uint id, string text, Drawing.ParagraphProperties pPr)
        => TextShape(id,
            new Drawing.Paragraph(pPr, new Drawing.Run(new Drawing.Text { Text = text })));

    private static P.Shape TextShape(uint id, string text, Drawing.ParagraphProperties pPr, Drawing.RunProperties runProperties)
        => TextShape(id,
            new Drawing.Paragraph(pPr, new Drawing.Run(runProperties, new Drawing.Text { Text = text })));

    private static P.Shape TextShape(uint id, params Drawing.Paragraph[] paragraphs)
        => TextShape(id, paragraphs, placeholderType: null);

    private static P.Shape TextShape(uint id, Drawing.Paragraph[] paragraphs, PlaceholderValues? placeholderType)
    {
        var appProps = new ApplicationNonVisualDrawingProperties();
        if (placeholderType != null)
        {
            appProps.Append(new PlaceholderShape { Type = placeholderType.Value });
        }

        var textBody = new TextBody(
            new Drawing.BodyProperties { LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0 },
            new Drawing.ListStyle());
        foreach (var paragraph in paragraphs)
        {
            textBody.Append(paragraph);
        }

        return new P.Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = id, Name = $"Text {id}" },
                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks { NoGrouping = true }),
                appProps),
            new ShapeProperties(
                new Drawing.Transform2D(
                    new Drawing.Offset { X = Pt(60), Y = Pt(60) },
                    new Drawing.Extents { Cx = Pt(400), Cy = Pt(200) })),
            textBody);
    }

    private string CreatePptx(string fileName, Action<SlideMaster>? customizeMaster, params OpenXmlElement[] slideElements)
    {
        var path = Path.Combine(_tempDir, fileName);

        using (var document = PresentationDocument.Create(path, PresentationDocumentType.Presentation))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new Presentation
            {
                SlideMasterIdList = new SlideMasterIdList(),
                SlideIdList = new SlideIdList(),
                SlideSize = new SlideSize { Cx = (int)Pt(720), Cy = (int)Pt(540), Type = SlideSizeValues.Screen4x3 }
            };

            var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
            slideMasterPart.SlideMaster = new SlideMaster(
                new CommonSlideData(CreateShapeTree()),
                new ColorMap
                {
                    Background1 = Drawing.ColorSchemeIndexValues.Light1,
                    Text1 = Drawing.ColorSchemeIndexValues.Dark1,
                    Background2 = Drawing.ColorSchemeIndexValues.Light2,
                    Text2 = Drawing.ColorSchemeIndexValues.Dark2,
                    Accent1 = Drawing.ColorSchemeIndexValues.Accent1,
                    Accent2 = Drawing.ColorSchemeIndexValues.Accent2,
                    Accent3 = Drawing.ColorSchemeIndexValues.Accent3,
                    Accent4 = Drawing.ColorSchemeIndexValues.Accent4,
                    Accent5 = Drawing.ColorSchemeIndexValues.Accent5,
                    Accent6 = Drawing.ColorSchemeIndexValues.Accent6,
                    Hyperlink = Drawing.ColorSchemeIndexValues.Hyperlink,
                    FollowedHyperlink = Drawing.ColorSchemeIndexValues.FollowedHyperlink
                },
                new SlideLayoutIdList());
            customizeMaster?.Invoke(slideMasterPart.SlideMaster);

            var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
            slideLayoutPart.SlideLayout = new P.SlideLayout(new CommonSlideData(CreateShapeTree()));
            slideLayoutPart.AddPart(slideMasterPart);
            slideMasterPart.SlideMaster.SlideLayoutIdList!.Append(new SlideLayoutId
            {
                Id = 2147483649,
                RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart)
            });

            presentationPart.Presentation.SlideMasterIdList.Append(new SlideMasterId
            {
                Id = 2147483648,
                RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
            });

            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.Slide = new Slide(new CommonSlideData(CreateShapeTree(slideElements)));
            slidePart.AddPart(slideLayoutPart);

            presentationPart.Presentation.SlideIdList.Append(new SlideId
            {
                Id = 256,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });
        }

        return path;
    }

    private static ShapeTree CreateShapeTree(params OpenXmlElement[] elements)
    {
        var shapeTree = new ShapeTree(
            new NonVisualGroupShapeProperties(
                new NonVisualDrawingProperties { Id = 0, Name = "" },
                new NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(
                new Drawing.TransformGroup(
                    new Drawing.Offset { X = 0, Y = 0 },
                    new Drawing.Extents { Cx = 0, Cy = 0 },
                    new Drawing.ChildOffset { X = 0, Y = 0 },
                    new Drawing.ChildExtents { Cx = 0, Cy = 0 })));

        foreach (var element in elements)
        {
            shapeTree.Append(element);
        }

        return shapeTree;
    }

    private static long Pt(double points) => (long)Math.Round(points * EmusPerPoint);
}
