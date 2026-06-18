using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Converters;
using OfficeEditor.Core.Services;
using System.Globalization;
using System.Text.RegularExpressions;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace DocxEditor.Tests.Unit;

public sealed class DocxToTypstConverterTests : IDisposable
{
    private readonly string tempDirectory = Path.Combine(Path.GetTempPath(), "DocxToTypstConverterTests", Guid.NewGuid().ToString("N"));

    public DocxToTypstConverterTests()
    {
        Directory.CreateDirectory(tempDirectory);
    }

    public void Dispose()
    {
        if (Directory.Exists(tempDirectory))
        {
            Directory.Delete(tempDirectory, recursive: true);
        }
    }

    [Fact]
    public void GenerateTypstSource_WithPlainParagraph_RendersParagraphText()
    {
        string path = CreateDocx("plain.docx", body => body.Append(CreateParagraph("Hello world")));

        string typst = ConvertToTypst(path);

        Assert.Contains("#set page(width: 8.27in, height: 11.69in", typst);
        Assert.Contains("#set text(font: \"Liberation Serif\", size: 11pt)", typst);
        Assert.Contains($"{Environment.NewLine}Hello world{Environment.NewLine}", typst);
    }

    [Fact]
    public void GenerateTypstSource_WithTypstSpecialCharacters_EscapesTextContent()
    {
        string path = CreateDocx("escaping.docx", body => body.Append(CreateParagraph(@"# $ % & @ * _ ^ ~ ` < > [ ] \")));

        string typst = ConvertToTypst(path);

        Assert.Contains(@"\# \$ \% \& \@ \* \_ \^ \~ \` \< \> \[ \] \", typst);
    }

    [Fact]
    public void GenerateTypstSource_WithFormattedTypstDelimiters_EscapesNestedMarkupContent()
    {
        W.Run run = new(
            new RunProperties(new Bold()),
            new Text(@"Value [A] #1 @x <tag> * ) % \"));
        string path = CreateDocx("formatted-escaping.docx", body => body.Append(new W.Paragraph(run)));

        string typst = ConvertToTypst(path);

        Assert.Contains(@"#strong[Value \[A\] \#1 \@x \<tag\> \* ) \% \\]", typst);
    }

    [Fact]
    public void GenerateTypstSource_WithTableCellTypstDelimiters_EscapesCellContent()
    {
        W.Table table = new(new TableRow(new TableCell(CreateParagraph(@"Cell [x] @ref <x> # * % \"))));
        string path = CreateDocx("table-escaping.docx", body => body.Append(table));

        string typst = ConvertToTypst(path);

        Assert.Contains(@"[Cell \[x\] \@ref \<x\> \# \* \% \\]", typst);
    }

    [Fact]
    public void GenerateTypstSource_WithEscapedFormattedAndTableText_Compiles()
    {
        W.Run run = new(new RunProperties(new Bold()), new Text(@"Formatted [x] @ref <x> # * % \"));
        W.Table table = new(new TableRow(new TableCell(CreateParagraph(@"Cell [x] @ref <x> # * % \"))));
        string path = CreateDocx("compile-escaping.docx", body =>
        {
            body.Append(new W.Paragraph(run));
            body.Append(table);
        });

        string typst = ConvertToTypst(path);
        using TypstCompilerService compiler = new();

        CompileResult result = compiler.Compile(typst, new CompileOptions
        {
            Format = OutputFormat.Pdf,
            WorkingDirectory = tempDirectory
        });

        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.Pages);
    }

    [Fact]
    public void GenerateTypstSource_WithRunFormatting_RendersBoldItalicUnderlineSizeColorAndFont()
    {
        W.Run run = new(
            new RunProperties(
                new Bold(),
                new Italic(),
                new Underline { Val = UnderlineValues.Single },
                new FontSize { Val = "28" },
                new Color { Val = "1a2b3c" },
                new RunFonts { Ascii = "Aptos" }),
            new Text("Styled"));
        string path = CreateDocx("formatted.docx", body => body.Append(new W.Paragraph(run)));

        string typst = ConvertToTypst(path);

        Assert.Contains("#text(fill: rgb(\"#1A2B3C\"), size: 14pt, font: \"Aptos\")[#underline[#emph[#strong[Styled]]]]", typst);
    }

    [Fact]
    public void GenerateTypstSource_WithBasedOnParagraphStyle_AppliesInheritedRunFormatting()
    {
        W.Paragraph paragraph = new(
            new ParagraphProperties(new ParagraphStyleId { Val = "Derived" }),
            new W.Run(new RunProperties(new Italic()), new Text("Inherited")));
        string path = CreateDocx("style-inheritance.docx", body => body.Append(paragraph), mainPart =>
        {
            AddStyles(mainPart,
                new Style(new StyleRunProperties(new Bold(), new Color { Val = "336699" })) { Type = StyleValues.Paragraph, StyleId = "Base" },
                new Style(new BasedOn { Val = "Base" }, new StyleRunProperties(new FontSize { Val = "24" })) { Type = StyleValues.Paragraph, StyleId = "Derived" });
        });

        string typst = ConvertToTypst(path);

        Assert.Contains("#text(fill: rgb(\"#336699\"), size: 12pt)[#emph[#strong[Inherited]]]", typst);
    }

    [Fact]
    public void GenerateTypstSource_WithThemeColorAndFont_ResolvesThemeValues()
    {
        W.Paragraph paragraph = new(
            new ParagraphProperties(new ParagraphStyleId { Val = "Themed" }),
            new W.Run(new Text("Themed text")));
        string path = CreateDocx("theme.docx", body => body.Append(paragraph), mainPart =>
        {
            AddTheme(mainPart);
            AddStyles(mainPart, new Style(new StyleRunProperties(
                new Color { ThemeColor = ThemeColorValues.Accent1 },
                new RunFonts { AsciiTheme = ThemeFontValues.MinorHighAnsi })) { Type = StyleValues.Paragraph, StyleId = "Themed" });
        });

        string typst = ConvertToTypst(path);

        Assert.Contains("#text(fill: rgb(\"#112233\"), font: \"Carlito\")[Themed text]", typst);
    }

    [Fact]
    public void GenerateTypstSource_WithSectionProperties_RendersPageSizeAndMargins()
    {
        string path = CreateDocx("page-setup.docx", body =>
        {
            body.Append(CreateParagraph("Sized page"));
            body.Append(new SectionProperties(
                new PageSize { Width = 12240, Height = 15840 },
                new PageMargin { Left = 720, Right = 1440, Top = 2160, Bottom = 2880 }));
        });

        string typst = ConvertToTypst(path);

        Assert.Contains("#set page(width: 8.5in, height: 11in, margin: (left: 0.5in, right: 1in, top: 1.5in, bottom: 2in))", typst);
    }

    [Fact]
    public void GenerateTypstSource_WithPageBreak_RendersTypstPageBreak()
    {
        string path = CreateDocx("page-break.docx", body =>
        {
            body.Append(CreateParagraph("Before"));
            body.Append(new W.Paragraph(new W.Run(new Break { Type = BreakValues.Page })));
            body.Append(CreateParagraph("After"));
        });

        string typst = ConvertToTypst(path);

        Assert.Contains("Before", typst);
        Assert.Contains("#pagebreak()", typst);
        Assert.Contains("After", typst);
    }

    [Fact]
    public void GenerateTypstSource_WithParagraphSpacingAndAlignment_RendersCompileSafeTypst()
    {
        W.Paragraph paragraph = new(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { Before = "240", After = "120" }),
            new W.Run(new Text("Centered spaced text")));
        string path = CreateDocx("paragraph-spacing-alignment.docx", body => body.Append(paragraph));

        string typst = ConvertToTypst(path);
        using TypstCompilerService compiler = new();

        CompileResult result = compiler.Compile(typst, new CompileOptions
        {
            Format = OutputFormat.Pdf,
            WorkingDirectory = tempDirectory
        });

        Assert.DoesNotContain("#par(above", typst);
        Assert.DoesNotContain("#par(below", typst);
        Assert.Contains("#block(width: 100%, above: 12pt, below: 6pt)[#align(center)[Centered spaced text]]", typst);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.Pages);
    }

    [Fact]
    public void GenerateTypstSource_WithEmptyStyledParagraph_PreservesSpacingBetweenHeadings()
    {
        string path = CreateDocx("empty-styled-paragraph.docx", body =>
        {
            body.Append(new W.Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
                new W.Run(new Text("Primary Title"))));
            body.Append(new W.Paragraph(
                new ParagraphProperties(
                    new ParagraphStyleId { Val = "Heading2" },
                    new SpacingBetweenLines { Before = "240", After = "120" })));
            body.Append(new W.Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "Heading2" }),
                new W.Run(new Text("Secondary Title"))));
        });

        string typst = ConvertToTypst(path);

        int primaryIndex = typst.IndexOf("Primary Title", StringComparison.Ordinal);
        int secondaryIndex = typst.IndexOf("Secondary Title", StringComparison.Ordinal);
        Assert.True(primaryIndex >= 0);
        Assert.True(secondaryIndex > primaryIndex);
        string between = typst[primaryIndex..secondaryIndex];
        Assert.Contains("#block(", between);
        Assert.Contains("above: 12pt", between);
        Assert.Contains("below: 6pt", between);
        Assert.Contains("#box(height:", between);
    }

    [Fact]
    public void GenerateTypstSource_WithLeadingManualLineBreaks_AddsWordLineBoxSpacing()
    {
        W.Paragraph paragraph = new(
            new ParagraphProperties(new ParagraphStyleId { Val = "TitleStyle" }),
            new W.Run(new W.Break()),
            new W.Run(new W.Break()),
            new W.Run(new Text("Lower title")));
        string path = CreateDocx("leading-linebreaks.docx", body => body.Append(paragraph), mainPart =>
        {
            AddStyles(mainPart, new Style(
                new StyleParagraphProperties(new SpacingBetweenLines { Before = "240" }),
                new StyleRunProperties(new FontSize { Val = "32" }))
            {
                Type = StyleValues.Paragraph,
                StyleId = "TitleStyle"
            });
        });

        string typst = ConvertToTypst(path);

        Assert.Contains("#block(above: 12pt)[#v(24pt)#linebreak()#linebreak()", typst);
    }

    [Fact]
    public void GenerateTypstSource_WithAutoLineSpacing_UsesReadableTypstLeading()
    {
        W.Paragraph paragraph = new(
            new ParagraphProperties(new SpacingBetweenLines { Line = "276", LineRule = LineSpacingRuleValues.Auto }),
            new W.Run(new Text("Line one Line two Line three")));
        string path = CreateDocx("auto-line-spacing.docx", body => body.Append(paragraph));

        string typst = ConvertToTypst(path);

        Assert.Contains("#par(leading: 7.15pt)", typst);
        Assert.DoesNotContain("leading: 2.8pt", typst);
    }

    [Fact]
    public void GenerateTypstSource_WithTrailingManualLineBreak_AddsEmptyLineAfterParagraph()
    {
        W.Paragraph paragraph = new(
            new ParagraphProperties(new SpacingBetweenLines { After = "200", Line = "276", LineRule = LineSpacingRuleValues.Auto }),
            new W.Run(new Text("Body")),
            new W.Run(new W.Break()));
        string path = CreateDocx("trailing-linebreak.docx", body =>
        {
            body.Append(paragraph);
            body.Append(CreateParagraph("Following"));
        });

        string typst = ConvertToTypst(path);

        Assert.Contains("#block(below: 28.15pt)[#par(leading: 7.15pt)[Body#linebreak()]]", typst);
    }

    [Fact]
    public void GenerateTypstSource_WithOrderedAndBulletLists_RendersSeparateListItems()
    {
        string path = CreateDocx("lists.docx", body =>
        {
            body.Append(CreateListParagraph("First", 1));
            body.Append(CreateListParagraph("Second", 1));
            body.Append(CreateParagraph("Between lists"));
            body.Append(CreateListParagraph("Bullet", 2));
        }, ConfigureNumbering);

        string typst = ConvertToTypst(path);

        Assert.Contains("#enum[First][Second]", typst);
        Assert.Contains("#list[Bullet]", typst);
    }

    [Fact]
    public void GenerateTypstSource_WithSimpleTable_RendersCellsAndColumnCount()
    {
        W.Table table = new(
            new TableProperties(new TableBorders(new TopBorder { Val = BorderValues.Single })),
            new TableRow(new TableCell(CreateParagraph("A1")), new TableCell(CreateParagraph("B1"))),
            new TableRow(new TableCell(CreateParagraph("A2")), new TableCell(CreateParagraph("B2"))));
        string path = CreateDocx("table.docx", body => body.Append(table));

        string typst = ConvertToTypst(path);

        Assert.Contains("#table(columns: 2, inset: 4pt, stroke: 0.5pt", typst);
        Assert.Contains("[A1]", typst);
        Assert.Contains("[B1]", typst);
        Assert.Contains("[A2]", typst);
        Assert.Contains("[B2]", typst);
    }

    [Fact]
    public void GenerateTypstSource_WithTableStyleFirstRow_AppliesHeaderShadingAndFormatting()
    {
        W.Table table = new(
            new TableProperties(new TableStyle { Val = "HeaderTable" }),
            new TableRow(new TableCell(CreateParagraph("Head"))),
            new TableRow(new TableCell(CreateParagraph("Body"))));
        string path = CreateDocx("table-style.docx", body => body.Append(table), mainPart =>
        {
            AddStyles(mainPart, new Style(
                new TableStyleProperties(
                    new TableCellProperties(new Shading { Fill = "D9EAF7" }),
                    new RunProperties(new Bold(), new Color { Val = "FFFFFF" }))
                { Type = TableStyleOverrideValues.FirstRow }) { Type = StyleValues.Table, StyleId = "HeaderTable" });
        });

        string typst = ConvertToTypst(path);

        Assert.Contains("table.cell(fill: rgb(\"#D9EAF7\"))[#text(fill: rgb(\"#FFFFFF\"))[#strong[Head]]]", typst);
        Assert.Contains("[Body]", typst);
    }

    [Fact]
    public void GenerateTypstSource_WithEmbeddedImage_ExtractsImageAndRendersImageCall()
    {
        string path = CreateDocx("image.docx", body => body.Append(new W.Paragraph()), AddPngImageToDocument);

        using WordprocessingDocument document = WordprocessingDocument.Open(path, false);
        using DocxToTypstConverter converter = new(document);
        var typstDocument = converter.Convert();
        string typst = converter.GenerateTypstSource(typstDocument);

        Assert.Contains("#image(\"assets/image-1.png\", width: 1in, height: 0.5in)", typst);
        Assert.True(File.Exists(Path.Combine(typstDocument.AssetsDirectory, "image-1.png")));
    }

    [Fact]
    public void GenerateTypstSource_WithMissingOrLinkedImageRelationship_SkipsImageWithoutThrowing()
    {
        string path = CreateDocx("missing-linked-image.docx", body => body.Append(new W.Paragraph()), mainPart =>
        {
            mainPart.Document!.Body!.Append(
                new W.Paragraph(new W.Run(CreateDrawing("rIdMissing"))),
                new W.Paragraph(new W.Run(CreateDrawing(mainPart.AddExternalRelationship(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                    new Uri("https://example.com/image.png")).Id!, linked: true))));
        });

        string typst = ConvertToTypst(path);

        Assert.DoesNotContain("#image(", typst);
    }

    [Fact]
    public void GenerateTypstSource_WithVmlTextbox_RendersPlacedFilledShapeWithoutDuplicateBodyText()
    {
        string path = CreateDocx("vml-textbox.docx", body =>
        {
            body.Append(new W.Paragraph(new W.Run(CreateVmlPicture("""
                <v:shape xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" id="TextBox1" style="position:absolute;margin-left:72pt;margin-top:36pt;width:144pt;height:54pt" fillcolor="#D9EAF7" strokecolor="#1A2B3C">
                  <v:textbox>
                    <w:txbxContent>
                      <w:p><w:r><w:t>Box #1 [safe]</w:t></w:r></w:p>
                    </w:txbxContent>
                  </v:textbox>
                </v:shape>
                """))));
        });

        string typst = ConvertToTypst(path);

        Assert.Contains("#place(dx: 72pt, dy: 36pt)[#rect(width: 144pt, height: 54pt, fill: rgb(\"#D9EAF7\"), stroke: rgb(\"#1A2B3C\"), inset: 4pt)", typst);
        Assert.DoesNotContain("#place(dx: 72pt, dy: 36pt, #rect", typst);
        Assert.Contains(@"Box \#1 \[safe\]", typst);
        Assert.Equal(1, CountOccurrences(typst, "Box"));
    }

    [Fact]
    public void GenerateTypstSource_WithUnpositionedVmlTextbox_RendersCompileSafeRectBody()
    {
        string path = CreateDocx("vml-textbox-unpositioned.docx", body =>
        {
            body.Append(new W.Paragraph(new W.Run(CreateVmlPicture("""
                <v:shape xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" id="TextBox2" style="width:72pt;height:36pt" fillcolor="#FFFFFF">
                  <v:textbox>
                    <w:txbxContent>
                      <w:p><w:r><w:t>Plain box</w:t></w:r></w:p>
                    </w:txbxContent>
                  </v:textbox>
                </v:shape>
                """))));
        });

        string typst = ConvertToTypst(path);

        Assert.Contains("#rect(width: 72pt, height: 36pt, fill: rgb(\"#FFFFFF\"), stroke: none, inset: 4pt)[Plain box]", typst);
    }

    [Fact]
    public void GenerateTypstSource_WithFilledVmlRectangle_RendersShapeWithoutText()
    {
        string path = CreateDocx("vml-rect.docx", body =>
        {
            body.Append(new W.Paragraph(new W.Run(CreateVmlPicture("""
                <v:shape xmlns:v="urn:schemas-microsoft-com:vml" id="Rect1" style="width:90pt;height:18pt" fillcolor="FFCC00" />
                """))));
        });

        string typst = ConvertToTypst(path);

        Assert.Contains("#rect(width: 90pt, height: 18pt, fill: rgb(\"#FFCC00\"), stroke: none)", typst);
    }

    [Fact]
    public void GenerateTypstSource_WithLargeUnitlessVmlGeometry_ConvertsHundredthsOfPoint()
    {
        string path = CreateDocx("vml-hundredths-unitless.docx", body =>
        {
            body.Append(new W.Paragraph(new W.Run(CreateVmlPicture("""
                <v:shape xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" id="TextBoxHundredths" style="position:absolute;margin-left:15959;margin-top:14760;width:61201;height:4680" fillcolor="#D9EAF7">
                  <v:textbox>
                    <w:txbxContent>
                      <w:p><w:r><w:t>Converted hundredths box</w:t></w:r></w:p>
                    </w:txbxContent>
                  </v:textbox>
                </v:shape>
                """))));
        });

        string typst = ConvertToTypst(path);

        Assert.Contains("#place(dx: 159.59pt, dy: 147.6pt)[#rect(width: 612.01pt, height: 46.8pt", typst);
        Assert.DoesNotContain("#place(dx: 1.257pt", typst);
        Assert.DoesNotContain("width: 4.819pt", typst);
        Assert.DoesNotContain("61201pt", typst);
        Assert.Equal(1, CountOccurrences(typst, "Converted hundredths box"));
    }

    [Fact]
    public void GenerateTypstSource_WithImpossibleVmlGeometry_RendersTextFallbackWithoutHugePlacement()
    {
        string path = CreateDocx("vml-impossible.docx", body =>
        {
            body.Append(new W.Paragraph(new W.Run(CreateVmlPicture("""
                <v:shape xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" id="TextBoxHuge" style="position:absolute;margin-left:18189pt;margin-top:36pt;width:61201pt;height:54pt" fillcolor="#FFFFFF">
                  <v:textbox>
                    <w:txbxContent>
                      <w:p><w:r><w:t>Fallback box</w:t></w:r></w:p>
                    </w:txbxContent>
                  </v:textbox>
                </v:shape>
                """))));
        });

        string typst = ConvertToTypst(path);

        Assert.DoesNotContain("#place(dx: 18189pt", typst);
        Assert.DoesNotContain("61201pt", typst);
        Assert.Contains("#rect(fill: rgb(\"#FFFFFF\"), stroke: none, inset: 4pt)[Fallback box]", typst);
        Assert.Equal(1, CountOccurrences(typst, "Fallback box"));
    }

    [Fact]
    public void GenerateTypstSource_WithDuplicateTextboxText_PrefersPositionedShapeOnce()
    {
        string path = CreateDocx("vml-duplicate-textbox.docx", body =>
        {
            body.Append(new W.Paragraph(new W.Run(CreateVmlPicture("""
                <v:shape xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" id="TextBoxA" style="width:72pt;height:36pt" fillcolor="#FFFFFF">
                  <v:textbox><w:txbxContent><w:p><w:r><w:t>Duplicate box</w:t></w:r></w:p></w:txbxContent></v:textbox>
                </v:shape>
                <v:shape xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" id="TextBoxB" style="position:absolute;margin-left:72pt;margin-top:36pt;width:72pt;height:36pt" fillcolor="#FFFFFF">
                  <v:textbox><w:txbxContent><w:p><w:r><w:t>Duplicate box</w:t></w:r></w:p></w:txbxContent></v:textbox>
                </v:shape>
                """))));
        });

        string typst = ConvertToTypst(path);

        Assert.Contains("#place(dx: 72pt, dy: 36pt)[#rect(width: 72pt, height: 36pt", typst);
        Assert.Equal(1, CountOccurrences(typst, "Duplicate box"));
    }

    [Fact]
    public void GenerateTypstSource_WithHeaderAndFooterText_RendersPageHeaderAndFooter()
    {
        string path = CreateDocx("header-footer-text.docx", body => body.Append(CreateParagraph("Body text")), mainPart =>
        {
            HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new Header(CreateParagraph("Header text"));
            headerPart.Header.Save();

            FooterPart footerPart = mainPart.AddNewPart<FooterPart>();
            footerPart.Footer = new Footer(CreateParagraph("Footer text"));
            footerPart.Footer.Save();

            mainPart.Document!.Body!.Append(new SectionProperties(
                new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(headerPart) },
                new FooterReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(footerPart) }));
        });

        string typst = ConvertToTypst(path);

        Assert.Contains("header: [Header text]", typst);
        Assert.Contains("footer: [Footer text]", typst);
        Assert.Contains("Body text", typst);
    }

    [Fact]
    public void GenerateTypstSource_WithHeaderImage_ExtractsImageFromHeaderPart()
    {
        string path = CreateDocx("header-image.docx", body => body.Append(CreateParagraph("Body text")), mainPart =>
        {
            HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
            ImagePart imagePart = headerPart.AddImagePart(ImagePartType.Png);
            using MemoryStream stream = new(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII="));
            imagePart.FeedData(stream);

            headerPart.Header = new Header(new W.Paragraph(new W.Run(CreateDrawing(headerPart.GetIdOfPart(imagePart)))));
            headerPart.Header.Save();

            mainPart.Document!.Body!.Append(new SectionProperties(
                new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(headerPart) }));
        });

        using WordprocessingDocument document = WordprocessingDocument.Open(path, false);
        using DocxToTypstConverter converter = new(document);
        var typstDocument = converter.Convert();
        string typst = converter.GenerateTypstSource(typstDocument);

        Assert.Contains("header: [#image(\"assets/image-1.png\", width: 1in, height: 0.5in)]", typst);
        Assert.True(File.Exists(Path.Combine(typstDocument.AssetsDirectory, "image-1.png")));
    }

    [Fact]
    public void GenerateTypstSource_WithTitlePageAndDefaultHeader_DoesNotApplyDefaultHeaderToFirstPage()
    {
        string path = CreateDocx("title-page-default-header.docx", body =>
        {
            body.Append(CreateParagraph("Cover"));
            body.Append(new W.Paragraph(new W.Run(new W.Break { Type = BreakValues.Page })));
            body.Append(CreateParagraph("Body"));
        }, mainPart =>
        {
            HeaderPart defaultHeaderPart = mainPart.AddNewPart<HeaderPart>();
            defaultHeaderPart.Header = new Header(CreateParagraph("Default header"));
            defaultHeaderPart.Header.Save();

            mainPart.Document!.Body!.Append(new SectionProperties(
                new TitlePage(),
                new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(defaultHeaderPart) }));
        });

        string typst = ConvertToTypst(path);

        int firstPageSet = typst.IndexOf("#set page(", StringComparison.Ordinal);
        int pageBreak = typst.IndexOf("#pagebreak()", StringComparison.Ordinal);
        int defaultHeader = typst.IndexOf("header: [Default header]", StringComparison.Ordinal);
        Assert.True(firstPageSet >= 0);
        Assert.True(pageBreak > firstPageSet);
        Assert.True(defaultHeader > pageBreak);
    }

    [Fact]
    public void GenerateTypstSource_WithLaterSectionDefaultHeaderAndFooter_DelaysUntilAfterPageBreak()
    {
        string path = CreateDocx("later-section-default-header-footer.docx", body =>
        {
            body.Append(CreateParagraph("Cover"));
            body.Append(new W.Paragraph(
                new W.ParagraphProperties(new SectionProperties()),
                new W.Run(new W.Break { Type = BreakValues.Page })));
            body.Append(CreateParagraph("Body"));
        }, mainPart =>
        {
            HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new Header(CreateParagraph("Later header"));
            headerPart.Header.Save();

            FooterPart footerPart = mainPart.AddNewPart<FooterPart>();
            footerPart.Footer = new Footer(CreateParagraph("Later footer"));
            footerPart.Footer.Save();

            mainPart.Document!.Body!.Append(new SectionProperties(
                new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(headerPart) },
                new FooterReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(footerPart) }));
        });

        string typst = ConvertToTypst(path);

        string firstSet = typst[..typst.IndexOf(Environment.NewLine, StringComparison.Ordinal)];
        int pageBreak = typst.IndexOf("#pagebreak()", StringComparison.Ordinal);
        int laterHeader = typst.IndexOf("header: [Later header]", StringComparison.Ordinal);
        int laterFooter = typst.IndexOf("footer: [Later footer]", StringComparison.Ordinal);
        Assert.DoesNotContain("Later header", firstSet);
        Assert.DoesNotContain("Later footer", firstSet);
        Assert.True(pageBreak >= 0);
        Assert.True(laterHeader > pageBreak);
        Assert.True(laterFooter > pageBreak);
    }

    [Fact]
    public void GenerateTypstSource_WithTitlePageFirstHeaderAndFooter_RendersFirstVariantsOnFirstPage()
    {
        string path = CreateDocx("title-page-first-header-footer.docx", body => body.Append(CreateParagraph("Cover")), mainPart =>
        {
            HeaderPart firstHeaderPart = mainPart.AddNewPart<HeaderPart>();
            firstHeaderPart.Header = new Header(CreateParagraph("First header"));
            firstHeaderPart.Header.Save();

            FooterPart firstFooterPart = mainPart.AddNewPart<FooterPart>();
            firstFooterPart.Footer = new Footer(CreateParagraph("First footer"));
            firstFooterPart.Footer.Save();

            HeaderPart defaultHeaderPart = mainPart.AddNewPart<HeaderPart>();
            defaultHeaderPart.Header = new Header(CreateParagraph("Default header"));
            defaultHeaderPart.Header.Save();

            mainPart.Document!.Body!.Append(new SectionProperties(
                new TitlePage(),
                new HeaderReference { Type = HeaderFooterValues.First, Id = mainPart.GetIdOfPart(firstHeaderPart) },
                new FooterReference { Type = HeaderFooterValues.First, Id = mainPart.GetIdOfPart(firstFooterPart) },
                new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(defaultHeaderPart) }));
        });

        string typst = ConvertToTypst(path);

        string firstSet = typst[..typst.IndexOf(Environment.NewLine, StringComparison.Ordinal)];
        Assert.Contains("First header", firstSet);
        Assert.Contains("First footer", firstSet);
        Assert.Contains("#context if counter(page).get().first() == 1", firstSet);
        Assert.Contains("else [Default header]", firstSet);
    }

    [Fact]
    public void GenerateTypstSource_WithFooterText_RendersFooterContent()
    {
        string path = CreateDocx("footer-text.docx", body => body.Append(CreateParagraph("Body text")), mainPart =>
        {
            FooterPart footerPart = mainPart.AddNewPart<FooterPart>();
            footerPart.Footer = new Footer(new SdtBlock(new SdtContentBlock(new W.Paragraph(
                new W.Run(new Text("Page footer text ")),
                new SimpleField(new W.Run(new Text("1"))) { Instruction = "PAGE" }))));
            footerPart.Footer.Save();

            mainPart.Document!.Body!.Append(new SectionProperties(
                new FooterReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(footerPart) }));
        });

        string typst = ConvertToTypst(path);

        Assert.Contains("footer: [Page footer text #context counter(page).display()]", typst);
    }

    [Fact]
    public void GenerateTypstSource_WithComplexPageFields_RendersTypstCountersAndSuppressesCachedValues()
    {
        string path = CreateDocx("page-fields.docx", body =>
        {
            W.Paragraph paragraph = new(new W.Run(new Text("Page ")));
            paragraph.Append(CreateComplexFieldRun("PAGE", "9"));
            paragraph.Append(new W.Run(new Text(" of ")));
            paragraph.Append(CreateComplexFieldRun("NUMPAGES", "99"));
            body.Append(paragraph);
        });

        string typst = ConvertToTypst(path);

        Assert.Contains("Page #context counter(page).display() of #context counter(page).final().at(0)", typst);
        Assert.DoesNotContain("99", typst);
        Assert.DoesNotContain("Page 9", typst);
    }

    [Fact]
    public void GenerateTypstSource_WithWingdingsCheckbox_RendersDrawnCheckbox()
    {
        W.Run run = new(
            new RunProperties(new RunFonts { Ascii = "Wingdings", HighAnsi = "Wingdings" }),
            new Text("\u00FE"));
        string path = CreateDocx("checkbox.docx", body => body.Append(new W.Paragraph(run)));

        string typst = ConvertToTypst(path);

        Assert.Contains("#box(width: 8pt, height: 8pt, stroke: 0.7pt)", typst);
        Assert.DoesNotContain("font: \"Wingdings\"", typst);
    }

    [Fact]
    public void GenerateTypstSource_WithSymbolBullet_MapsToUnicodeBullet()
    {
        string path = CreateDocx("symbol-bullet.docx", body =>
        {
            body.Append(new W.Paragraph(new W.Run(new SymbolChar { Font = "Symbol", Char = "F0B7" })));
        });

        string typst = ConvertToTypst(path);

        Assert.Contains("•", typst);
        Assert.DoesNotContain("F0B7", typst);
    }

    [Fact]
    public void GenerateTypstSource_WithUnsupportedEmfImage_RendersPlaceholderAndDiagnostic()
    {
        string path = CreateDocx("emf-image.docx", body => body.Append(new W.Paragraph()), mainPart =>
        {
            ImagePart imagePart = mainPart.AddImagePart("image/x-emf");
            using MemoryStream stream = new([1, 2, 3, 4]);
            imagePart.FeedData(stream);
            W.Paragraph paragraph = mainPart.Document!.Body!.Elements<W.Paragraph>().Single();
            paragraph.Append(new W.Run(CreateDrawing(mainPart.GetIdOfPart(imagePart))));
        });

        using WordprocessingDocument document = WordprocessingDocument.Open(path, false);
        using DocxToTypstConverter converter = new(document);
        var typstDocument = converter.Convert();
        string typst = converter.GenerateTypstSource(typstDocument);

        Assert.Contains("Exclusive Windows image format", typst);
        Assert.DoesNotContain("#image(\"assets/image-1.emf\"", typst);
        Assert.Contains(typstDocument.Diagnostics, diagnostic => diagnostic.Contains("unsupported image format", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void GenerateTypstSource_WithAnchoredHeaderImage_RendersPlacedImageWithDimensions()
    {
        string path = CreateDocx("anchored-header-image.docx", body => body.Append(CreateParagraph("Body text")), mainPart =>
        {
            HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
            ImagePart imagePart = headerPart.AddImagePart(ImagePartType.Png);
            using MemoryStream stream = new(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII="));
            imagePart.FeedData(stream);

            headerPart.Header = new Header(new W.Paragraph(new W.Run(CreateAnchorDrawing(headerPart.GetIdOfPart(imagePart)))));
            headerPart.Header.Save();

            mainPart.Document!.Body!.Append(new SectionProperties(
                new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(headerPart) }));
        });

        string typst = ConvertToTypst(path);

        Assert.Contains("header: [#place(dx: 36pt, dy: 18pt)[#image(\"assets/image-1.png\", width: 1in, height: 0.5in)]]", typst);
    }

    [Fact]
    public void GenerateTypstSource_WithHeaderTable_RendersTableInHeader()
    {
        string path = CreateDocx("header-table.docx", body => body.Append(CreateParagraph("Body text")), mainPart =>
        {
            HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new Header(new W.Table(
                new TableRow(new TableCell(CreateParagraph("Logo")), new TableCell(CreateParagraph("Report")))));
            headerPart.Header.Save();

            mainPart.Document!.Body!.Append(new SectionProperties(
                new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(headerPart) }));
        });

        string typst = ConvertToTypst(path);

        Assert.Contains("header: [#table(columns: 2", typst);
        Assert.Contains("[Logo]", typst);
        Assert.Contains("[Report]", typst);
    }

    [Fact]
    public void GenerateTypstSource_WithMultipleSectionHeaders_EmitsPageSettingsAtSectionBoundary()
    {
        string path = CreateDocx("section-headers.docx", body =>
        {
            body.Append(CreateParagraph("First section"));
            body.Append(new W.Paragraph(
                new W.ParagraphProperties(new SectionProperties()),
                new W.Run(new W.Break { Type = BreakValues.Page })));
            body.Append(CreateParagraph("Second section"));
        }, mainPart =>
        {
            HeaderPart secondHeader = mainPart.AddNewPart<HeaderPart>();
            secondHeader.Header = new Header(CreateParagraph("Second header"));
            secondHeader.Header.Save();

            mainPart.Document!.Body!.Append(new SectionProperties(
                new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(secondHeader) }));
        });

        string typst = ConvertToTypst(path);

        int pageBreak = typst.IndexOf("#pagebreak()", StringComparison.Ordinal);
        int header = typst.IndexOf("header: [Second header]", StringComparison.Ordinal);
        Assert.True(header > pageBreak);
    }

    [Fact]
    public void Header_WithMultipleParagraphs_DoesNotEmbedNewlineInContentArray()
    {
        string path = CreateDocx("multi-para-header-no-newline.docx", body => body.Append(CreateParagraph("Body")), mainPart =>
        {
            HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new Header(
                CreateParagraph("Header line A"),
                CreateParagraph("Header line B"));
            headerPart.Header.Save();

            mainPart.Document!.Body!.Append(new SectionProperties(
                new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(headerPart) }));
        });

        string typst = ConvertToTypst(path);

        int setPageStart = typst.IndexOf("#set page(", StringComparison.Ordinal);
        Assert.True(setPageStart >= 0, "Expected the generated source to start with `#set page(`.");
        int nextNewline = typst.IndexOf('\n', setPageStart);
        string setPageLine = nextNewline > setPageStart
            ? typst[setPageStart..nextNewline]
            : typst[setPageStart..];

        Assert.DoesNotContain("\n", setPageLine);
        Assert.DoesNotContain("\r", setPageLine);
    }

    [Fact]
    public void Header_WithMultipleParagraphs_JoinsBlocksWithParagraphBreak()
    {
        string path = CreateDocx("multi-para-header-wrapped.docx", body => body.Append(CreateParagraph("Body")), mainPart =>
        {
            HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new Header(
                CreateParagraph("Header line A"),
                CreateParagraph("Header line B"));
            headerPart.Header.Save();

            mainPart.Document!.Body!.Append(new SectionProperties(
                new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(headerPart) }));
        });

        string typst = ConvertToTypst(path);

        Assert.Contains("header: [Header line A#parbreak()Header line B]", typst);
        Assert.DoesNotContain("[[", typst);
        Assert.DoesNotContain("]]", typst);
    }

    [Fact]
    public void Footer_WithTextAndAnchoredImage_SplitsImageAsDecorativeBottom()
    {
        string path = CreateDocx("footer-text-anchored-image.docx", body => body.Append(CreateParagraph("Body")), mainPart =>
        {
            FooterPart footerPart = mainPart.AddNewPart<FooterPart>();
            ImagePart imagePart = footerPart.AddImagePart(ImagePartType.Png);
            using MemoryStream stream = new(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII="));
            imagePart.FeedData(stream);

            footerPart.Footer = new Footer(
                new W.Paragraph(new W.Run(new Text("Footer note"))),
                new W.Paragraph(new W.Run(CreateAnchorDrawing(footerPart.GetIdOfPart(imagePart)))));
            footerPart.Footer.Save();

            mainPart.Document!.Body!.Append(new SectionProperties(
                new FooterReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(footerPart) }));
        });

        string typst = ConvertToTypst(path);

        Assert.Contains("#place(bottom + center, [#place(dx: 36pt, dy: 18pt)[#image(", typst);

        int setPageStart = typst.IndexOf("#set page(", StringComparison.Ordinal);
        int nextNewline = typst.IndexOf('\n', setPageStart);
        string setPageLine = nextNewline > setPageStart
            ? typst[setPageStart..nextNewline]
            : typst[setPageStart..];

        Assert.Contains("footer: [", setPageLine);
        Assert.DoesNotContain("#image(", setPageLine);
    }

    [Fact]
    public void Header_WithSingleParagraph_RemainsUnwrappedInContentArray()
    {
        string path = CreateDocx("single-para-header.docx", body => body.Append(CreateParagraph("Body")), mainPart =>
        {
            HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new Header(CreateParagraph("Header text"));
            headerPart.Header.Save();

            mainPart.Document!.Body!.Append(new SectionProperties(
                new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(headerPart) }));
        });

        string typst = ConvertToTypst(path);

        Assert.Contains("header: [Header text]", typst);
        Assert.DoesNotContain("[[Header text]]", typst);
    }

    [Theory]
    [InlineData("w:before=\"100\" w:after=\"200\" w:afterAutospacing=\"1\"", "above: 5pt", "below: 10pt")]
    [InlineData("w:before=\"200\" w:after=\"100\"", "above: 10pt", "below: 5pt")]
    [InlineData("w:after=\"200\" w:afterAutospacing=\"1\"", null, "below: 10pt")]
    public void Paragraph_WithAfterAutospacing_FallsBackToRawBeforeValue(string spacingAttributes, string? expectedAbove, string expectedBelow)
    {
        W.Paragraph paragraph = new();
        ParagraphProperties pPr = new();
        pPr.InnerXml = $"<w:spacing xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" {spacingAttributes}/>";
        paragraph.PrependChild(pPr);
        paragraph.Append(new W.Run(new Text("Spaced")));

        string path = CreateDocx("spacing-after-autospacing.docx", body => body.Append(paragraph));

        string typst = ConvertToTypst(path);

        if (expectedAbove is not null)
        {
            Assert.Contains(expectedAbove, typst);
        }
        else
        {
            Assert.DoesNotContain("above:", typst);
        }

        Assert.Contains(expectedBelow, typst);
    }

    [Fact]
    public void GenerateTypstSource_WithMultiParagraphHeaderAndFooter_DoesNotContainDoubleBrackets()
    {
        string path = CreateDocx("header-footer-no-double-brackets.docx", body => body.Append(CreateParagraph("Body")), mainPart =>
        {
            HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new Header(CreateParagraph("Header line A"), CreateParagraph("Header line B"));
            headerPart.Header.Save();

            FooterPart footerPart = mainPart.AddNewPart<FooterPart>();
            footerPart.Footer = new Footer(CreateParagraph("Footer line A"), CreateParagraph("Footer line B"));
            footerPart.Footer.Save();

            mainPart.Document!.Body!.Append(new SectionProperties(
                new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(headerPart) },
                new FooterReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(footerPart) }));
        });

        string typst = ConvertToTypst(path);

        string headerValue = ExtractPageOptionValue(typst, "header");
        string footerValue = ExtractPageOptionValue(typst, "footer");

        Assert.False(headerValue.Contains("[[", StringComparison.Ordinal), $"Header value contains literal double brackets: {headerValue}");
        Assert.False(headerValue.Contains("]]", StringComparison.Ordinal), $"Header value contains literal double brackets: {headerValue}");
        Assert.False(footerValue.Contains("[[", StringComparison.Ordinal), $"Footer value contains literal double brackets: {footerValue}");
        Assert.False(footerValue.Contains("]]", StringComparison.Ordinal), $"Footer value contains literal double brackets: {footerValue}");
    }

    [Fact]
    public void Footer_WithRightAlignedParagraphAndLeading_RendersAlignOutsidePar()
    {
        string path = CreateDocx("footer-right-align-leading.docx", body => body.Append(CreateParagraph("Body")), mainPart =>
        {
            FooterPart footerPart = mainPart.AddNewPart<FooterPart>();
            W.Paragraph footerParagraph = new(
                new ParagraphProperties(
                    new Justification { Val = JustificationValues.Right },
                    new SpacingBetweenLines { Line = "240", LineRule = LineSpacingRuleValues.Auto }),
                new W.Run(new Text("Right aligned footer")));
            footerPart.Footer = new Footer(footerParagraph);
            footerPart.Footer.Save();

            mainPart.Document!.Body!.Append(new SectionProperties(
                new FooterReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(footerPart) }));
        });

        string typst = ConvertToTypst(path);

        string footerValue = ExtractPageOptionValue(typst, "footer");
        Assert.Contains("#align(right)[#par(leading: 1pt)[Right aligned footer]]", footerValue);
    }

    [Fact]
    public void GenerateTypstSource_WithIntermediateSectionBreak_InheritsEffectiveHeaderFooterSet()
    {
        string path = Path.Combine(tempDirectory, "intermediate-section-inherits.docx");
        using (WordprocessingDocument document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            MainDocumentPart mainPart = document.AddMainDocumentPart();
            FooterPart firstFooterPart = mainPart.AddNewPart<FooterPart>();
            firstFooterPart.Footer = new Footer(CreateParagraph("First footer"));
            firstFooterPart.Footer.Save();

            FooterPart finalFooterPart = mainPart.AddNewPart<FooterPart>();
            finalFooterPart.Footer = new Footer(CreateParagraph("Final footer"));
            finalFooterPart.Footer.Save();

            Body body = new();
            mainPart.Document = new W.Document(body);

            body.Append(new W.Paragraph(
                new W.ParagraphProperties(
                    new SectionProperties(
                        new FooterReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(firstFooterPart) })),
                new W.Run(new Text("Section 1"))));

            body.Append(new W.Paragraph(
                new W.ParagraphProperties(new SectionProperties()),
                new W.Run(new W.Break { Type = BreakValues.Page })));

            body.Append(CreateParagraph("Section 2"));

            body.Append(new SectionProperties(
                new FooterReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(finalFooterPart) }));

            mainPart.Document.Save();
        }

        string typst = ConvertToTypst(path);

        int firstSetPage = typst.IndexOf("#set page(", StringComparison.Ordinal);
        int firstFooterInFirstSet = typst.IndexOf("footer: [First footer]", firstSetPage, StringComparison.Ordinal);
        Assert.True(firstFooterInFirstSet >= 0);

        int pageBreak = typst.IndexOf("#pagebreak()", firstSetPage, StringComparison.Ordinal);
        Assert.True(pageBreak > firstFooterInFirstSet);

        int intermediateSetPage = typst.IndexOf("#set page(", pageBreak, StringComparison.Ordinal);
        Assert.True(intermediateSetPage > pageBreak);
        int firstFooterInIntermediateSet = typst.IndexOf("footer: [First footer]", intermediateSetPage, StringComparison.Ordinal);
        Assert.True(firstFooterInIntermediateSet > intermediateSetPage, "Expected intermediate section to inherit the effective footer.");

        int finalFooterIndex = typst.IndexOf("footer: [Final footer]", StringComparison.Ordinal);
        Assert.True(finalFooterIndex > firstFooterInIntermediateSet);
    }

    [Fact]
    public void GenerateTypstSource_WithFooterFrameHorizontalAlignment_PreservesRightAlignment()
    {
        string path = Path.Combine(tempDirectory, "footer-frame-align.docx");
        using (WordprocessingDocument document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            MainDocumentPart mainPart = document.AddMainDocumentPart();
            FooterPart footerPart = mainPart.AddNewPart<FooterPart>();
            ParagraphProperties pPr = new();
            pPr.InnerXml = "<w:framePr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" w:xAlign=\"right\"/>";
            W.Paragraph paragraph = new(pPr);
            paragraph.Append(new W.Run(new Text("Page ")));
            paragraph.Append(new SimpleField(new W.Run(new Text("1"))) { Instruction = "PAGE" });
            footerPart.Footer = new Footer(paragraph);
            footerPart.Footer.Save();

            Body body = new();
            mainPart.Document = new W.Document(body);
            body.Append(CreateParagraph("Body"));
            body.Append(new SectionProperties(
                new FooterReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(footerPart) }));
            mainPart.Document.Save();
        }

        string typst = ConvertToTypst(path);

        string footerValue = ExtractPageOptionValue(typst, "footer");
        Assert.Contains("#align(right)", footerValue);
        Assert.Contains("counter(page).display()", footerValue);
    }

    [Fact]
    public void Footer_WithStyledPageFields_RendersCountersWithInheritedFormatting()
    {
        string path = CreateDocx("footer-styled-page-fields.docx", body => body.Append(CreateParagraph("Body")), mainPart =>
        {
            FooterPart footerPart = mainPart.AddNewPart<FooterPart>();
            W.Paragraph paragraph = new(new W.Run(new Text("Page ")));
            paragraph.Append(CreateStyledComplexFieldRun("PAGE", "1", "006EB6", "18"));
            paragraph.Append(new W.Run(new Text(" of ")));
            paragraph.Append(CreateStyledComplexFieldRun("NUMPAGES", "5", "006EB6", "18"));
            footerPart.Footer = new Footer(paragraph);
            footerPart.Footer.Save();

            mainPart.Document!.Body!.Append(new SectionProperties(
                new FooterReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(footerPart) }));
        });

        string typst = ConvertToTypst(path);

        string footerValue = ExtractPageOptionValue(typst, "footer");
        Assert.Contains(
            "Page #text(fill: rgb(\"#006EB6\"), size: 9pt)[#context counter(page).display()] of #text(fill: rgb(\"#006EB6\"), size: 9pt)[#context counter(page).final().at(0)]",
            footerValue);
    }

    [Fact]
    public void Footer_WithBorderedInlineImage_RendersLineAboveImageInsideFooterContent()
    {
        string path = CreateDocx("footer-bordered-inline-image.docx", body => body.Append(CreateParagraph("Body")), mainPart =>
        {
            FooterPart footerPart = mainPart.AddNewPart<FooterPart>();
            ImagePart imagePart = footerPart.AddImagePart(ImagePartType.Png);
            using MemoryStream stream = new(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII="));
            imagePart.FeedData(stream);

            W.Paragraph borderedImageParagraph = new(
                new ParagraphProperties(new ParagraphBorders(new TopBorder { Val = BorderValues.Single, Size = 6, Color = "00257D" })),
                new W.Run(CreateDrawing(footerPart.GetIdOfPart(imagePart))));

            footerPart.Footer = new Footer(
                new W.Paragraph(new W.Run(new Text("NOTE text"))),
                borderedImageParagraph);
            footerPart.Footer.Save();

            mainPart.Document!.Body!.Append(new SectionProperties(
                new FooterReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(footerPart) }));
        });

        string typst = ConvertToTypst(path);

        string footerValue = ExtractPageOptionValue(typst, "footer");
        Assert.Contains("NOTE text", footerValue);
        Assert.Contains("#line(length: 100%, stroke: 0.75pt + rgb(\"#00257D\"))#image(", footerValue);
        Assert.DoesNotContain("[[", footerValue);
        Assert.DoesNotContain("]]", footerValue);
        Assert.DoesNotContain("#place(bottom + center", typst);
    }

    [Fact]
    public void Footer_WithBorderedAnchoredImage_RendersLineAbovePlacedDecorativeImage()
    {
        string path = CreateDocx("footer-bordered-anchored-image.docx", body => body.Append(CreateParagraph("Body")), mainPart =>
        {
            FooterPart footerPart = mainPart.AddNewPart<FooterPart>();
            ImagePart imagePart = footerPart.AddImagePart(ImagePartType.Png);
            using MemoryStream stream = new(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII="));
            imagePart.FeedData(stream);

            W.Paragraph borderedImageParagraph = new(
                new ParagraphProperties(new ParagraphBorders(new TopBorder { Val = BorderValues.Single, Size = 6, Color = "00257D" })),
                new W.Run(CreateAnchorDrawing(footerPart.GetIdOfPart(imagePart))));

            footerPart.Footer = new Footer(
                new W.Paragraph(new W.Run(new Text("NOTE text"))),
                borderedImageParagraph);
            footerPart.Footer.Save();

            mainPart.Document!.Body!.Append(new SectionProperties(
                new FooterReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(footerPart) }));
        });

        string typst = ConvertToTypst(path);

        Assert.Contains("#place(bottom + center, [#line(length: 100%, stroke: 0.75pt + rgb(\"#00257D\")) #place(dx: 36pt, dy: 18pt)[#image(\"assets/image-1.png\", width: 1in, height: 0.5in)]])", typst);

        string footerValue = ExtractPageOptionValue(typst, "footer");
        Assert.Contains("NOTE text", footerValue);
        Assert.DoesNotContain("[[", footerValue);
        Assert.DoesNotContain("]]", footerValue);
        Assert.DoesNotContain("#image(", footerValue);
        Assert.DoesNotContain("#line(", footerValue);
    }

    [Fact]
    public void GenerateTypstSource_WithImageOnlyParagraph_PreservesStyleAlignmentAndSpacing()
    {
        string path = CreateDocx("image-only-paragraph.docx", body => body.Append(new W.Paragraph()), mainPart =>
        {
            AddPngImageToDocument(mainPart);
            W.Paragraph paragraph = mainPart.Document!.Body!.Elements<W.Paragraph>().Single();
            paragraph.PrependChild(new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { Before = "480" }));
        });

        string typst = ConvertToTypst(path);

        Assert.Contains("#align(center)", typst);
        Assert.Contains("#block(", typst);
        Assert.Contains("above: 24pt", typst);
        Assert.Contains("width: 100%", typst);
        Assert.Contains("#image(", typst);
    }

    [Fact]
    public void GenerateTypstSource_WithNormalStyleFontButNoDocDefaults_UsesNormalStyleFontAsDefault()
    {
        W.Paragraph paragraph = new(new W.Run(new Text("Sample text")));
        string path = CreateDocx("normal-style-font.docx", body => body.Append(paragraph), mainPart =>
        {
            AddStyles(mainPart, new Style(
                new StyleRunProperties(new RunFonts { Ascii = "Segoe UI" }))
            {
                Type = StyleValues.Paragraph,
                StyleId = "Normal",
                Default = true
            });
        });

        string typst = ConvertToTypst(path);

        // Segoe UI is mapped to a sans-serif fallback on systems where it is not installed.
        Assert.Contains("\"Liberation Sans\"", typst);
        Assert.DoesNotContain("\"Liberation Serif\"", typst);
    }

    [Fact]
    public void GenerateTypstSource_WithStyleInheritedBottomBorder_RendersBottomBorderInFooter()
    {
        string path = CreateDocx("style-bottom-border.docx", body => body.Append(CreateParagraph("Body text")), mainPart =>
        {
            AddStyles(mainPart, new Style(
                new StyleParagraphProperties(new ParagraphBorders(new BottomBorder { Val = BorderValues.Single, Size = 12, Color = "FF0000" })),
                new StyleRunProperties(new Bold()))
            {
                Type = StyleValues.Paragraph,
                StyleId = "Bordered"
            });

            FooterPart footerPart = mainPart.AddNewPart<FooterPart>();
            footerPart.Footer = new Footer(new W.Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "Bordered" }),
                new W.Run(new Text("Footer text"))));
            footerPart.Footer.Save();

            mainPart.Document!.Body!.Append(new SectionProperties(
                new FooterReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(footerPart) }));
        });

        string typst = ConvertToTypst(path);

        string footerValue = ExtractPageOptionValue(typst, "footer");
        Assert.Contains("#line(", footerValue);
        Assert.Contains("rgb(\"#FF0000\")", footerValue);
    }

    [Fact]
    public void GenerateTypstSource_WithLargeFooterDistance_ReservesFooterSpaceInPageMargin()
    {
        string path = CreateDocx("footer-distance.docx", body =>
        {
            body.Append(CreateParagraph("Body text"));
            body.Append(new SectionProperties(
                new PageSize { Width = 12240, Height = 15840 },
                new PageMargin { Left = 720, Right = 720, Top = 720, Bottom = 720, Footer = 2880 }));
        });

        string typst = ConvertToTypst(path);

        double bottomMarginInches = ExtractBottomMarginInches(typst);
        Assert.True(bottomMarginInches >= 2.0, $"Expected bottom margin >= 2in to reserve footer space, but got {bottomMarginInches}in.");
    }

    private string ConvertToTypst(string path)
    {
        using WordprocessingDocument document = WordprocessingDocument.Open(path, false);
        using DocxToTypstConverter converter = new(document);
        return converter.GenerateTypstSource(converter.Convert());
    }

    private string CreateDocx(string fileName, Action<Body> configureBody, Action<MainDocumentPart>? configureMainPart = null)
    {
        string path = Path.Combine(tempDirectory, fileName);
        using WordprocessingDocument document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        Body body = new();
        mainPart.Document = new W.Document(body);
        configureBody(body);
        configureMainPart?.Invoke(mainPart);
        mainPart.Document.Save();
        return path;
    }

    private static W.Paragraph CreateParagraph(string text) => new(new W.Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));

    private static W.Run[] CreateComplexFieldRun(string instruction, string cachedValue) =>
    [
        new W.Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
        new W.Run(new FieldCode(instruction) { Space = SpaceProcessingModeValues.Preserve }),
        new W.Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
        new W.Run(new Text(cachedValue)),
        new W.Run(new FieldChar { FieldCharType = FieldCharValues.End })
    ];

    private static W.Run[] CreateStyledComplexFieldRun(string instruction, string cachedValue, string color, string fontSize) =>
    [
        new W.Run(new RunProperties(new Color { Val = color }, new FontSize { Val = fontSize }), new FieldChar { FieldCharType = FieldCharValues.Begin }),
        new W.Run(new RunProperties(new Color { Val = color }, new FontSize { Val = fontSize }), new FieldCode(instruction) { Space = SpaceProcessingModeValues.Preserve }),
        new W.Run(new RunProperties(new Color { Val = color }, new FontSize { Val = fontSize }), new FieldChar { FieldCharType = FieldCharValues.Separate }),
        new W.Run(new RunProperties(new Color { Val = color }, new FontSize { Val = fontSize }), new Text(cachedValue)),
        new W.Run(new RunProperties(new Color { Val = color }, new FontSize { Val = fontSize }), new FieldChar { FieldCharType = FieldCharValues.End })
    ];

    private static int CountOccurrences(string value, string search)
    {
        int count = 0;
        int index = 0;
        while ((index = value.IndexOf(search, index, StringComparison.Ordinal)) >= 0)
        {
            count++;
            index += search.Length;
        }

        return count;
    }

    private static string ExtractPageOptionValue(string typst, string optionName)
    {
        string prefix = $"{optionName}: [";
        int start = typst.IndexOf(prefix, StringComparison.Ordinal);
        Assert.True(start >= 0, $"Expected generated Typst source to contain '{prefix}'.");
        start += prefix.Length - 1;
        int depth = 0;
        for (int i = start; i < typst.Length; i++)
        {
            if (typst[i] == '[')
            {
                depth++;
            }
            else if (typst[i] == ']')
            {
                depth--;
            }

            if (depth == 0)
            {
                return typst[start..(i + 1)];
            }
        }

        throw new InvalidOperationException($"Could not find closing bracket for '{prefix}'.");
    }

    private static double ExtractBottomMarginInches(string typst)
    {
        Match match = Regex.Match(typst, @"#set page\([^)]*bottom:\s*(\d+(?:\.\d+)?)in", RegexOptions.Singleline);
        Assert.True(match.Success, "Expected generated Typst source to set a bottom page margin.");
        return double.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
    }

    private static Picture CreateVmlPicture(string innerXml)
    {
        Picture picture = new();
        picture.InnerXml = innerXml;
        return picture;
    }

    private static W.Paragraph CreateListParagraph(string text, int numberingId) => new(
        new ParagraphProperties(new NumberingProperties(new NumberingLevelReference { Val = 0 }, new NumberingId { Val = numberingId })),
        new W.Run(new Text(text)));

    private static void ConfigureNumbering(MainDocumentPart mainPart)
    {
        NumberingDefinitionsPart numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering = new Numbering(
            new AbstractNum(new Level(new NumberingFormat { Val = NumberFormatValues.Decimal }) { LevelIndex = 0 }) { AbstractNumberId = 1 },
            new AbstractNum(new Level(new NumberingFormat { Val = NumberFormatValues.Bullet }) { LevelIndex = 0 }) { AbstractNumberId = 2 },
            new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 },
            new NumberingInstance(new AbstractNumId { Val = 2 }) { NumberID = 2 });
        numberingPart.Numbering.Save();
    }

    private static void AddStyles(MainDocumentPart mainPart, params Style[] styles)
    {
        StyleDefinitionsPart stylesPart = mainPart.StyleDefinitionsPart ?? mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles(styles);
        stylesPart.Styles.Save();
    }

    private static void AddTheme(MainDocumentPart mainPart)
    {
        ThemePart themePart = mainPart.ThemePart ?? mainPart.AddNewPart<ThemePart>();
        themePart.Theme = new A.Theme(
            new A.ThemeElements(
                new A.ColorScheme(
                    new A.Dark1Color(new A.RgbColorModelHex { Val = "000000" }),
                    new A.Light1Color(new A.RgbColorModelHex { Val = "FFFFFF" }),
                    new A.Dark2Color(new A.RgbColorModelHex { Val = "1F1F1F" }),
                    new A.Light2Color(new A.RgbColorModelHex { Val = "F2F2F2" }),
                    new A.Accent1Color(new A.RgbColorModelHex { Val = "112233" }),
                    new A.Accent2Color(new A.RgbColorModelHex { Val = "445566" }),
                    new A.Accent3Color(new A.RgbColorModelHex { Val = "778899" }),
                    new A.Accent4Color(new A.RgbColorModelHex { Val = "AABBCC" }),
                    new A.Accent5Color(new A.RgbColorModelHex { Val = "DDEEFF" }),
                    new A.Accent6Color(new A.RgbColorModelHex { Val = "123456" }),
                    new A.Hyperlink(new A.RgbColorModelHex { Val = "0000FF" }),
                    new A.FollowedHyperlinkColor(new A.RgbColorModelHex { Val = "800080" })) { Name = "Test" },
                new A.FontScheme(
                    new A.MajorFont(new A.LatinFont { Typeface = "Aptos Display" }),
                    new A.MinorFont(new A.LatinFont { Typeface = "Carlito" })) { Name = "Test" },
                new A.FormatScheme { Name = "Test" })) { Name = "Test" };
        themePart.Theme.Save();
    }

    private static void AddPngImageToDocument(MainDocumentPart mainPart)
    {
        ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
        using MemoryStream stream = new(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII="));
        imagePart.FeedData(stream);
        string relationshipId = mainPart.GetIdOfPart(imagePart);
        W.Paragraph paragraph = mainPart.Document!.Body!.Elements<W.Paragraph>().Single();
        paragraph.Append(new W.Run(CreateDrawing(relationshipId)));
    }

    private static Drawing CreateDrawing(string relationshipId, bool linked = false) => new(
        new DW.Inline(
            new DW.Extent { Cx = 914400, Cy = 457200 },
            new DW.DocProperties { Id = 1U, Name = "Tiny image" },
            new A.Graphic(new A.GraphicData(
                new Pic.Picture(
                    new Pic.NonVisualPictureProperties(
                        new Pic.NonVisualDrawingProperties { Id = 1U, Name = "tiny.png" },
                        new Pic.NonVisualPictureDrawingProperties()),
                    new Pic.BlipFill(linked ? new A.Blip { Link = relationshipId } : new A.Blip { Embed = relationshipId }, new A.Stretch(new A.FillRectangle())),
                    new Pic.ShapeProperties()))
            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })));

    private static Drawing CreateAnchorDrawing(string relationshipId) => new(
        new DW.Anchor(
            new DW.SimplePosition { X = 0L, Y = 0L },
            new DW.HorizontalPosition(new DW.PositionOffset("457200")) { RelativeFrom = DW.HorizontalRelativePositionValues.Page },
            new DW.VerticalPosition(new DW.PositionOffset("228600")) { RelativeFrom = DW.VerticalRelativePositionValues.Page },
            new DW.Extent { Cx = 914400, Cy = 457200 },
            new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
            new DW.WrapNone(),
            new DW.DocProperties { Id = 2U, Name = "Anchored image" },
            new A.Graphic(new A.GraphicData(
                new Pic.Picture(
                    new Pic.NonVisualPictureProperties(
                        new Pic.NonVisualDrawingProperties { Id = 2U, Name = "anchored.png" },
                        new Pic.NonVisualPictureDrawingProperties()),
                    new Pic.BlipFill(new A.Blip { Embed = relationshipId }, new A.Stretch(new A.FillRectangle())),
                    new Pic.ShapeProperties()))
            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
        {
            BehindDoc = false,
            LayoutInCell = true,
            AllowOverlap = true,
            SimplePos = false,
            RelativeHeight = 0U
        });
}
