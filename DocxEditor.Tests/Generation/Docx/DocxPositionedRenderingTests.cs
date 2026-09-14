using System.Text.Json;
using DocxEditor.Core.Builders;
using OfficeEditor.Core.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocxEditor.Core.Converters;
using DocxEditor.Core.Generation;
using DocxEditor.Core.Models;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace DocxEditor.Tests.Generation.Docx;

public sealed class DocxPositionedRenderingTests
{
    [Theory]
    [InlineData("page", 24, 36)]
    [InlineData("margin", 60, 90)]
    public void JsonRoundTrip_AnchoredText_UsesAnchorAndReferenceFrame(string anchor, int dx, int dy)
    {
        var bytes = new DocxGenerator().GenerateToBytes(Document(anchor)).Content;
        using var stream = new MemoryStream(bytes);
        using var docx = WordprocessingDocument.Open(stream, false);
        using var converter = new DocxToTypstConverter(docx);
        var model = converter.Convert();
        var text = Assert.Single(model.Blocks.OfType<TypstShapeBlock>(), s => s.Paragraphs.Count > 0);
        Assert.Equal(60, text.XPt);
        Assert.Equal(90, text.YPt);
        var source = converter.GenerateTypstSource(model);
        Assert.Contains($"#place(top + left, dx: {dx}pt, dy: {dy}pt)[#rect", source);
    }

    [Fact]
    public void JsonRoundTrip_ZeroMarginsAndEmptyHeaders_PreservesFullPageBackground()
    {
        var json = Document("page", 0).Replace("\"x\": 60, \"y\": 90", "\"x\": 0, \"y\": 0");
        var source = Convert(json);
        Assert.Contains("margin: (left: 0in, right: 0in, top: 0in, bottom: 0in)", source);
        Assert.Contains("#place(top + left, dx: 0pt, dy: 0pt)", source);
    }

    [Fact]
    public void JsonRoundTrip_TextBox_UsesEmittedInsets()
    {
        var source = Convert(Document("page"));
        Assert.Contains("inset: (left: 7.2pt, right: 7.2pt, top: 3.6pt, bottom: 3.6pt)", source);
        Assert.DoesNotContain("inset: 4pt", source);
    }

    [Fact]
    public void TextBox_ExplicitZeroAndAsymmetricInsets_ArePreserved()
    {
        var bytes = new DocxGenerator().GenerateToBytes(Document("page")).Content;
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using var docx = WordprocessingDocument.Open(stream, true);
        var bodyPr = docx.MainDocumentPart!.Document!.Descendants().Single(e => e.LocalName == "bodyPr");
        foreach (var (name, value) in new[] { ("lIns", "0"), ("rIns", "12700"), ("tIns", "25400"), ("bIns", "0") })
            bodyPr.SetAttribute(new OpenXmlAttribute(name, string.Empty, value));
        using var converter = new DocxToTypstConverter(docx);
        var source = converter.GenerateTypstSource(converter.Convert());
        Assert.Contains("inset: (left: 0pt, right: 1pt, top: 2pt, bottom: 0pt)", source);
    }

    [Fact]
    public void JsonRoundTrip_ImageAndShape_UseTheSamePageOrigin()
    {
        var image = TestImages.DataUriBase64(TestImages.Png(16, 16));
        var json = Document("page").Replace("\"positioned\": [", $$"""
            "positioned": [
                { "type": "image", "src": "{{image}}", "anchor": "page", "wrap": "inFrontOfText",
                  "x": 60, "y": 90, "width": 120, "height": 40 },
            """);
        var source = Convert(json);
        Assert.Contains("#place(top + left, dx: 24pt, dy: 36pt)[#image", source);
        Assert.Contains("#place(top + left, dx: 24pt, dy: 36pt)[#rect", source);
    }

    [Theory]
    [InlineData("English text", "Arial")]
    [InlineData("Expérience française", "Arial")]
    [InlineData("中文简历", "Noto Sans CJK SC")]
    public void JsonRoundTrip_Languages_PreserveRequestedFontAndText(string text, string font)
    {
        var source = Convert(Document("page", text: text, font: font));
        Assert.Contains(text, source);
        Assert.Contains(font == "Arial" ? "font: (\"Arial\", \"Liberation Sans\", \"Noto Sans\", \"Aptos\")" : "font: \"Noto Sans CJK SC\"", source);
    }

    [Fact]
    public void GroupedShapes_ComposeNestedTranslationScaleAndAnchor()
    {
        // Local child (15,25) -> inner group (20,30) at scale 2 -> (30,40).
        // Outer group maps (10,20) to (100,200), scale 3 -> (160,260).
        // Its top-level off is local to the anchor, not an additional page offset:
        // subtract (100,200), then anchor (60,90) -> page (120,150).
        var xml = """
            <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
              xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
              <wp:positionH relativeFrom="page"><wp:posOffset>762000</wp:posOffset></wp:positionH>
              <wp:positionV relativeFrom="page"><wp:posOffset>1143000</wp:posOffset></wp:positionV>
              <wp:extent cx="3810000" cy="3810000"/>
              <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">
                <wpg:wgp><wpg:grpSpPr><a:xfrm>
                  <a:off x="1270000" y="2540000"/><a:ext cx="3810000" cy="3810000"/>
                  <a:chOff x="127000" y="254000"/><a:chExt cx="1270000" cy="1270000"/>
                </a:xfrm></wpg:grpSpPr>
                  <wpg:grpSp><wpg:grpSpPr><a:xfrm>
                    <a:off x="254000" y="381000"/><a:ext cx="1270000" cy="1270000"/>
                    <a:chOff x="127000" y="254000"/><a:chExt cx="635000" cy="635000"/>
                  </a:xfrm></wpg:grpSpPr>
                    <wps:wsp><wps:spPr><a:xfrm>
                      <a:off x="190500" y="317500"/><a:ext cx="127000" cy="127000"/>
                    </a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="17245C"/></a:solidFill></wps:spPr></wps:wsp>
                  </wpg:grpSp>
                </wpg:wgp>
              </a:graphicData></a:graphic>
            </wp:anchor>
            """;
        using var stream = new MemoryStream();
        using var docx = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = docx.AddMainDocumentPart();
        main.Document = new W.Document(new W.Body(new W.Paragraph(new W.Run(new W.Drawing { InnerXml = xml })),
            new W.SectionProperties(new W.PageMargin { Left = 720, Top = 1080, Right = 720, Bottom = 720 })));
        using var converter = new DocxToTypstConverter(docx);
        var model = converter.Convert();
        var shape = Assert.Single(model.Blocks.OfType<TypstShapeBlock>());
        Assert.Equal(120, shape.XPt);
        Assert.Equal(150, shape.YPt);
        Assert.Equal(60, shape.WidthPt);
        Assert.Equal(60, shape.HeightPt);
        Assert.Contains("#place(top + left, dx: 84pt, dy: 96pt)", converter.GenerateTypstSource(model));
    }

    [Fact]
    public void PageAnchoredHeaderShape_RepeatsOnEveryRenderedPage()
    {
        var bytes = new DocxGenerator().GenerateToBytes(Document("page").Replace("\"type\": \"textBox\"", "\"type\": \"textBox\", \"fill\": \"#17245C\"")).Content;
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using var docx = WordprocessingDocument.Open(stream, true);
        var main = docx.MainDocumentPart!;
        var body = main.Document!.Body!;
        var drawing = body.Descendants<W.Drawing>().Single();
        drawing.Remove();
        var header = main.AddNewPart<HeaderPart>();
        header.Header = new W.Header(new W.Paragraph(new W.Run(drawing)));
        body.Elements<W.SectionProperties>().Single().RemoveAllChildren<W.HeaderReference>();
        body.Elements<W.SectionProperties>().Single().Append(new W.HeaderReference
        {
            Type = W.HeaderFooterValues.Default, Id = main.GetIdOfPart(header)
        });
        body.PrependChild(new W.Paragraph(new W.Run(new W.Text("Page one"), new W.Break { Type = W.BreakValues.Page }, new W.Text("Page two"))));
        using var converter = new DocxToTypstConverter(docx);
        var model = converter.Convert();
        var source = converter.GenerateTypstSource(model);
        Assert.Contains("foreground: [", source);
        if (Environment.GetEnvironmentVariable("OE_RUN_TYPST_COMPILE_TESTS") != "1") return;
        using var compiler = new TypstCompilerService();
        var result = compiler.Compile(source, new CompileOptions { Format = OutputFormat.Svg, WorkingDirectory = model.TempDirectory });
        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(2, result.Pages.Length);
        // The colored header rectangle must appear on both pages.
        foreach (var page in result.Pages)
        {
            var svg = System.Text.Encoding.UTF8.GetString(page);
            Assert.Contains("#17245c", svg, StringComparison.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public void JsonRoundTrip_LanguageDocuments_RenderAndExposeFontDiagnostics()
    {
        if (Environment.GetEnvironmentVariable("OE_RUN_TYPST_COMPILE_TESTS") != "1") return;
        foreach (var (text, font) in new[] { ("English text", "Arial"), ("Expérience française", "Helvetica"), ("中文简历", "Noto Sans CJK SC") })
        {
            using var builder = (DocumentBuilder)DocumentBuilder.Open(new DocxGenerator().GenerateToBytes(Document("page", text: text, font: font)).Content);
            var fontPath = Environment.GetEnvironmentVariable("OE_DOCX_TEST_FONT_PATH");
            var result = builder.ExportWithDiagnostics(new CompileOptions { Format = OutputFormat.Svg, FontDirectory = fontPath });
            Assert.True(result.Success, result.ErrorMessage);
            Assert.Single(result.Pages);
            if (font != "Noto Sans CJK SC" || !string.IsNullOrWhiteSpace(fontPath))
                Assert.DoesNotContain(result.Warnings, warning => warning.Contains("Missing glyphs", StringComparison.Ordinal));
            else
                Assert.Contains(result.Warnings, warning => warning.Contains("Missing glyphs", StringComparison.Ordinal));
        }
        using var missing = (DocumentBuilder)DocumentBuilder.Open(new DocxGenerator().GenerateToBytes(Document("page", text: "Missing \U0010FFFF", font: "OfficeEditor Missing Font 9281")).Content);
        var diagnosticResult = missing.ExportWithDiagnostics();
        Assert.True(diagnosticResult.Success, diagnosticResult.ErrorMessage);
        Assert.Contains(diagnosticResult.Warnings, warning => warning.Contains("unknown font family", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(diagnosticResult.Warnings, warning => warning.Contains("U+10FFFF", StringComparison.Ordinal));
    }

    internal static string Document(string anchor, int margin = 36, string text = "Positioned text", string font = "Arial") => $$"""
        {
          "version": "1.0",
          "design": { "page": { "size": "a4", "margins": {
            "left": {{margin}}, "right": {{margin}}, "top": {{(margin == 0 ? 0 : 54)}}, "bottom": {{margin}}
          } } },
          "sections": [{ "blocks": [{ "type": "paragraph", "text": " " }], "positioned": [
            { "type": "textBox", "anchor": "{{anchor}}", "wrap": "inFrontOfText",
              "x": 60, "y": 90, "width": 120, "height": 40,
              "runs": [{ "text": {{JsonSerializer.Serialize(text)}}, "font": {{JsonSerializer.Serialize(font)}}, "size": 12 }] }
          ] }]
        }
        """;

    private static string Convert(string json)
    {
        using var stream = new MemoryStream(new DocxGenerator().GenerateToBytes(json).Content);
        using var docx = WordprocessingDocument.Open(stream, false);
        using var converter = new DocxToTypstConverter(docx);
        return converter.GenerateTypstSource(converter.Convert());
    }
}
