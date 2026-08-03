using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Builders;
using DocxEditor.Core.Markdown.Rendering;
using DocxEditor.Tests.Generation.Docx;
using A = DocumentFormat.OpenXml.Drawing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace DocxEditor.Tests.Unit.Markdown;

/// <summary>
/// Markdown v2: rich recursive IR rendered to OOXML through <see cref="RichMarkdownRenderer"/>
/// and the <see cref="DocumentBuilder"/> integration. Focuses on the constructs the legacy
/// flat renderer could not express (nested formatting, real hyperlinks, images, nested
/// numbering, footnotes, tables, YAML core properties) and on package validity after rendering
/// and in-place replacement.
/// </summary>
public class RichMarkdownRendererTests
{
    // ---- Inline formatting -------------------------------------------------

    [Fact]
    public void Heading_RendersInlineFormattingAndStyle()
    {
        var bytes = Build("# Title *italic*");

        using var doc = Open(bytes);
        var body = doc.MainDocumentPart!.Document!.Body!;
        var heading = body.Elements<W.Paragraph>().First();
        Assert.Equal("Heading1", heading.ParagraphProperties!.ParagraphStyleId!.Val!.Value);
        Assert.Contains("Title", heading.InnerText);
        Assert.Contains("italic", heading.InnerText);
        Assert.Single(heading.Descendants<Italic>());
    }

    [Fact]
    public void Paragraph_NestedEmphasis_CombinesRunProperties()
    {
        var bytes = Build("**bold _italic_**");

        using var doc = Open(bytes);
        var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().Single();
        var runs = paragraph.Elements<W.Run>().ToList();
        Assert.Equal(2, runs.Count);

        Assert.NotNull(runs[0].RunProperties!.Bold);
        Assert.Null(runs[0].RunProperties!.Italic);

        Assert.NotNull(runs[1].RunProperties!.Bold);
        Assert.NotNull(runs[1].RunProperties!.Italic);
    }

    [Fact]
    public void Paragraph_EmphasisKinds_MapToMatchingRunProperties()
    {
        var bytes = Build("~~s~~ ~sub~ ^sup^ ++ins++ ==mark==");

        using var doc = Open(bytes);
        var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().Single();

        Assert.Single(paragraph.Descendants<Strike>());
        Assert.Contains(paragraph.Descendants<VerticalTextAlignment>(), v => v.Val!.Value == VerticalPositionValues.Subscript);
        Assert.Contains(paragraph.Descendants<VerticalTextAlignment>(), v => v.Val!.Value == VerticalPositionValues.Superscript);
        Assert.Contains(paragraph.Descendants<Underline>(), u => u.Val!.Value == UnderlineValues.Single);
        Assert.Contains(paragraph.Descendants<Highlight>(), h => h.Val!.Value == HighlightColorValues.Yellow);
    }

    [Fact]
    public void Paragraph_InlineCode_IsExactAndMonospace()
    {
        var bytes = Build("Use `code` here");

        using var doc = Open(bytes);
        var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().Single();
        var codeRun = paragraph.Descendants<W.Run>().Single(r =>
            r.RunProperties?.RunFonts?.Ascii?.Value == "Consolas");

        Assert.Equal("code", codeRun.InnerText);
    }

    [Fact]
    public void Paragraph_InlineCode_AppliesCommonMarkPaddingNormalization()
    {
        // A code span padded with one space on each side keeps the remaining padding verbatim.
        var bytes = Build("Use `  padded  ` and ``two ` ticks``");

        using var doc = Open(bytes);
        var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().Single();
        var codeRuns = paragraph.Descendants<W.Run>()
            .Where(r => r.RunProperties?.RunFonts?.Ascii?.Value == "Consolas")
            .Select(r => r.InnerText)
            .ToList();

        Assert.Contains(" padded ", codeRuns);
        Assert.Contains("two ` ticks", codeRuns);
    }

    // ---- Links and images --------------------------------------------------

    [Fact]
    public void Paragraph_Links_ProduceRealHyperlinkRelationships()
    {
        var bytes = Build("""
        [GitHub](https://github.com) and [ref][id] and <https://example.com>

        [id]: http://example.com
        """);

        using var doc = Open(bytes);
        var mainPart = doc.MainDocumentPart!;
        var body = mainPart.Document!.Body!;

        var hyperlinks = body.Descendants<W.Hyperlink>().ToList();
        Assert.Equal(3, hyperlinks.Count);

        var uris = hyperlinks
            .Select(h => mainPart.HyperlinkRelationships.First(r => r.Id == h.Id!.Value).Uri)
            .Select(u => u.AbsoluteUri)
            .ToList();
        Assert.Contains("https://github.com/", uris);
        Assert.Contains("http://example.com/", uris); // reference link resolved by the parser
        Assert.Contains("https://example.com/", uris); // autolink
    }

    [Fact]
    public void Paragraph_NonAbsoluteLink_RendersAsPlainText()
    {
        var bytes = Build("[x](not a url)");

        using var doc = Open(bytes);
        var mainPart = doc.MainDocumentPart!;
        var body = mainPart.Document!.Body!;

        Assert.DoesNotContain(body.Descendants<W.Hyperlink>(), h => true);
        Assert.Empty(mainPart.HyperlinkRelationships);
        Assert.Contains("x", body.InnerText);
    }

    [Fact]
    public void Paragraph_Breaks_HardAndSoft()
    {
        var hard = Build("line  \nhard");
        using (var doc = Open(hard))
        {
            var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().Single();
            Assert.Single(paragraph.Descendants<W.Break>());
        }

        var soft = Build("soft\nbreak");
        using (var doc = Open(soft))
        {
            var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().Single();
            Assert.DoesNotContain(paragraph.Descendants<W.Break>(), b => true);
            Assert.Equal("soft break", paragraph.InnerText);
        }

        var softAsBreak = Build("soft\nbreak", new MarkdownRenderOptions { SoftBreakMode = MarkdownSoftBreakMode.LineBreak });
        using (var doc = Open(softAsBreak))
        {
            var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().Single();
            Assert.Single(paragraph.Descendants<W.Break>());
        }
    }

    [Fact]
    public void Paragraph_EntitiesEmojiAndUnicode_AreDecoded()
    {
        var bytes = Build("a &amp; b :smile: 😀");

        using var doc = Open(bytes);
        var text = doc.MainDocumentPart!.Document!.Body!.InnerText;

        Assert.Contains("a & b", text);
        Assert.Contains("😄", text);
        Assert.Contains("😀", text);
    }

    [Fact]
    public void Image_DataUri_EmbedsPartWithAltText()
    {
        var png = TestImages.DataUriBase64(TestImages.Png(16, 16));
        var bytes = Build($"![alt text]({png})");

        using var doc = Open(bytes);
        var mainPart = doc.MainDocumentPart!;
        var body = mainPart.Document!.Body!;

        Assert.Single(mainPart.ImageParts);

        var docProperties = body.Descendants<Wp.DocProperties>().Single();
        Assert.Equal("alt text", docProperties.Description!.Value);

        var blip = body.Descendants<A.Blip>().Single();
        Assert.NotNull(blip.Embed);
        Assert.NotNull(mainPart.GetPartById(blip.Embed!.Value!));
    }

    [Fact]
    public void Image_RemoteUrl_FallsBackToVisibleAltWithoutFetching()
    {
        var bytes = Build("![logo](https://example.com/x.png)", new MarkdownRenderOptions { Strict = true });
        using var doc = Open(bytes);
        var mainPart = doc.MainDocumentPart!;

        Assert.Empty(mainPart.ImageParts);
        Assert.Contains("[image: logo]", mainPart.Document!.Body!.InnerText);
    }

    [Fact]
    public void Image_MissingLocalFile_FallsBackToVisibleAlt()
    {
        var bytes = Build("![x](does-not-exist.png)");

        using var doc = Open(bytes);
        var mainPart = doc.MainDocumentPart!;

        Assert.Empty(mainPart.ImageParts);
        Assert.Contains("[image: x]", mainPart.Document!.Body!.InnerText);
    }

    [Fact]
    public void Image_SmallIntrinsicSizing_RespectsNaturalSize()
    {
        // 16x16 px at 96 DPI = 12pt; well under the display cap.
        var png = TestImages.DataUriBase64(TestImages.Png(16, 16));
        var bytes = Build($"![x]({png})");

        using var doc = Open(bytes);
        var extent = doc.MainDocumentPart!.Document!.Body!.Descendants<Wp.Extent>().Single();
        Assert.Equal(12L * 12700, extent.Cx!.Value); // 12pt in EMU
        Assert.Equal(12L * 12700, extent.Cy!.Value);
    }

    // ---- Lists -------------------------------------------------------------

    [Fact]
    public void List_NestedOrdered_PreservesStartValueAndIndent()
    {
        var bytes = Build("""
        3. three
           1. nested one
           2. nested two
        4. four
        """);

        using var doc = Open(bytes);
        var mainPart = doc.MainDocumentPart!;
        var body = mainPart.Document!.Body!;
        var numbering = mainPart.NumberingDefinitionsPart!.Numbering!;

        Assert.NotEmpty(numbering.Elements<NumberingInstance>());

        // The top-level list's abstract definition starts at the authored value 3.
        Assert.Contains(
            numbering.Elements<AbstractNum>(),
            a => a.Elements<Level>().First().StartNumberingValue!.Val!.Value == 3);

        var paragraphs = body.Elements<W.Paragraph>().ToList();
        var three = paragraphs[0];
        var nestedOne = paragraphs[1];
        var four = paragraphs[3];

        Assert.Equal("360", three.ParagraphProperties!.Indentation!.Hanging!.Value);
        Assert.Equal("720", three.ParagraphProperties.Indentation.Left!.Value);
        Assert.Equal("1440", nestedOne.ParagraphProperties!.Indentation!.Left!.Value); // depth 1

        // Sibling "four" binds to the same top-level instance as "three".
        Assert.Equal(
            three.ParagraphProperties.NumberingProperties!.NumberingId!.Val!.Value,
            four.ParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value);
    }

    [Fact]
    public void List_EightLevelsOfNesting_RendersValidDocument()
    {
        var md = new System.Text.StringBuilder();
        for (var i = 0; i < 8; i++)
        {
            md.Append(' ', i * 2);
            md.AppendLine($"- level {i}");
        }

        using var temp = new TempDirectory();
        var path = temp.File("deep.docx");
        using (var builder = DocumentBuilder.Create(path))
        {
            builder.AddRichMarkdown(md.ToString());
            builder.Save();
        }
        OpenXmlAssert.NoDocxValidationErrors(path);

        using var doc = WordprocessingDocument.Open(path, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        Assert.Equal(8, body.Elements<W.Paragraph>().Count());
        Assert.Contains("level 7", body.InnerText);
    }

    [Fact]
    public void List_TaskCheckboxes_RenderGlyphs()
    {
        var bytes = Build("- [x] done\n- [ ] todo");

        using var doc = Open(bytes);
        var text = doc.MainDocumentPart!.Document!.Body!.InnerText;

        Assert.Contains("☑", text);
        Assert.Contains("☐", text);
    }

    // ---- Code / tables / quotes / definitions ------------------------------

    [Fact]
    public void CodeBlock_PreservesWhitespaceAndNewlines()
    {
        var bytes = Build("""
        ```csharp
        var x = 1;
          indented;
        last
        ```
        """);

        using var doc = Open(bytes);
        var mainPart = doc.MainDocumentPart!;
        var body = mainPart.Document!.Body!;

        var codeParagraphs = body.Elements<W.Paragraph>()
            .Where(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "Code")
            .ToList();
        Assert.Equal(3, codeParagraphs.Count);
        Assert.Equal("Code", codeParagraphs[0].ParagraphProperties!.ParagraphStyleId!.Val!.Value);

        var indented = codeParagraphs[1].Descendants<W.Text>().Single();
        Assert.Equal("  indented;", indented.Text);
        Assert.Equal(SpaceProcessingModeValues.Preserve, indented.Space!.Value);
    }

    [Fact]
    public void Table_HeaderShadingAlignmentAndFormattedCells()
    {
        var bytes = Build("""
        | Left | Center | Right |
        |:-----|:------:|------:|
        | a    | b      | **c** |
        """);

        using var doc = Open(bytes);
        var mainPart = doc.MainDocumentPart!;
        var body = mainPart.Document!.Body!;

        var table = body.Descendants<W.Table>().Single();
        var rows = table.Elements<W.TableRow>().ToList();
        Assert.Equal(2, rows.Count);

        var header = rows[0].Elements<W.TableCell>().ToList();
        Assert.Equal(3, header.Count);
        Assert.All(header, cell => Assert.NotNull(cell.TableCellProperties!.Shading));
        Assert.All(header, cell => Assert.NotNull(cell.Descendants<Bold>().FirstOrDefault()));

        var headerJustifications = header.Select(cell =>
            cell.Descendants<W.Paragraph>().First().ParagraphProperties!.Justification!.Val!.Value).ToList();
        Assert.Equal(JustificationValues.Left, headerJustifications[0]);
        Assert.Equal(JustificationValues.Center, headerJustifications[1]);
        Assert.Equal(JustificationValues.Right, headerJustifications[2]);

        var data = rows[1].Elements<W.TableCell>().ToList();
        Assert.Contains("c", data[2].InnerText);
        Assert.NotNull(data[2].Descendants<Bold>().FirstOrDefault());
    }

    [Fact]
    public void Blockquote_Recursive_RendersQuoteStyleAndIndent()
    {
        var bytes = Build("""
        > outer
        > > inner
        """);

        using var doc = Open(bytes);
        var body = doc.MainDocumentPart!.Document!.Body!;
        var paragraphs = body.Elements<W.Paragraph>().ToList();

        Assert.Equal("Quote", paragraphs[0].ParagraphProperties!.ParagraphStyleId!.Val!.Value);
        Assert.Equal("720", paragraphs[0].ParagraphProperties!.Indentation!.Left!.Value);

        Assert.Equal("Quote", paragraphs[1].ParagraphProperties!.ParagraphStyleId!.Val!.Value);
        Assert.Equal("1440", paragraphs[1].ParagraphProperties!.Indentation!.Left!.Value);
    }

    [Fact]
    public void DefinitionList_TermBoldDefinitionIndented()
    {
        var bytes = Build("""
        Term 1
        :   Definition one
        """);

        using var doc = Open(bytes);
        var body = doc.MainDocumentPart!.Document!.Body!;
        var paragraphs = body.Elements<W.Paragraph>().ToList();

        Assert.Equal("Term 1", paragraphs[0].InnerText);
        Assert.NotNull(paragraphs[0].Descendants<Bold>().FirstOrDefault());

        Assert.Equal("Definition one", paragraphs[1].InnerText);
        Assert.Equal("1080", paragraphs[1].ParagraphProperties!.Indentation!.Left!.Value);
    }

    [Fact]
    public void ThematicBreak_EmitsBorderedParagraph()
    {
        var bytes = Build("before\n\n---\n\nafter");

        using var doc = Open(bytes);
        var body = doc.MainDocumentPart!.Document!.Body!;
        var paragraphs = body.Elements<W.Paragraph>().ToList();

        Assert.Contains(paragraphs, p => p.ParagraphProperties?.ParagraphBorders?.BottomBorder is not null);
    }

    // ---- Footnotes / YAML / HTML / unknown ---------------------------------

    [Fact]
    public void Footnotes_ProduceValidFootnotesPart()
    {
        var bytes = Build("""
        A ref[^1].

        [^1]: The note with *emphasis*.
        """);

        using var doc = Open(bytes);
        var mainPart = doc.MainDocumentPart!;
        var body = mainPart.Document!.Body!;

        var footnotesPart = mainPart.FootnotesPart;
        Assert.NotNull(footnotesPart);

        var footnotes = footnotesPart!.Footnotes!.Elements<Footnote>().ToList();
        Assert.Contains(footnotes, f => f.Id!.Value == 0 && f.Type!.Value == FootnoteEndnoteValues.Separator);
        Assert.Contains(footnotes, f => f.Id!.Value == 1 && f.Type!.Value == FootnoteEndnoteValues.ContinuationSeparator);

        var note = footnotes.Single(f => f.Id!.Value == 2);
        Assert.Contains(note.Descendants<FootnoteReferenceMark>(), m => true);
        Assert.Contains("The note", note.InnerText);
        Assert.Contains(note.Descendants<Italic>(), i => true);

        var reference = body.Descendants<FootnoteReference>().Single();
        Assert.Equal(2u, reference.Id!.Value);
    }

    [Fact]
    public void YamlFrontMatter_MapsRecognizedKeysToCoreProperties()
    {
        var bytes = Build("""
        ---
        title: My Title
        author: Jane Doe
        subject: Test Subject
        keywords: one, two
        description: A description
        language: en-US
        ---
        # Heading
        """);

        using var doc = Open(bytes);
        var mainPart = doc.MainDocumentPart!;
        var corePart = doc.CoreFilePropertiesPart!;

        XDocument xml;
        using (var stream = corePart.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.Read))
        {
            xml = XDocument.Load(stream);
        }

        var dc = "http://purl.org/dc/elements/1.1/";
        var cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        Assert.Equal("My Title", xml.Root!.Element(XNamespace.Get(dc) + "title")!.Value);
        Assert.Equal("Jane Doe", xml.Root.Element(XNamespace.Get(dc) + "creator")!.Value);
        Assert.Equal("Test Subject", xml.Root.Element(XNamespace.Get(cp) + "subject")!.Value);
        Assert.Equal("one, two", xml.Root.Element(XNamespace.Get(cp) + "keywords")!.Value);
        Assert.Equal("A description", xml.Root.Element(XNamespace.Get(dc) + "description")!.Value);
        Assert.Equal("en-US", xml.Root.Element(XNamespace.Get(dc) + "language")!.Value);

        // Front matter is metadata, not body content.
        Assert.DoesNotContain("My Title", mainPart.Document!.Body!.InnerText);
        Assert.Contains("Heading", mainPart.Document.Body.InnerText);
    }

    [Fact]
    public void HtmlBlock_RendersAsEscapedVisibleText()
    {
        var bytes = Build("<div>\nraw\n</div>\n\npara <span>inline</span>");

        using var doc = Open(bytes);
        var body = doc.MainDocumentPart!.Document!.Body!;

        Assert.Contains("raw", body.InnerText);
        Assert.Contains("inline", body.InnerText);
    }

    // ---- Validator / relationships -----------------------------------------

    [Fact]
    public void ComprehensiveDocument_PassesOpenXmlValidator()
    {
        var png = TestImages.DataUriBase64(TestImages.Png(16, 16));
        var markdown = """
        ---
        title: Comprehensive
        author: OfficeEditor
        ---

        # Title *with emphasis*

        Paragraph with **bold _nested_** and [link](https://example.com), `code`, :rocket:.

        ![chart](__PNG__)

        3. three
           - nested bullet
        4. four

        > A quote with *italic*.

        ```json
        {
          "a": 1,
          "indented": true
        }
        ```

        | Name | Value |
        |:-----|------:|
        | A    | **B** |

        Term
        :   Definition

        A note[^1].

        ---

        [^1]: The footnote.

        <div>raw html</div>
        """.Replace("__PNG__", png);

        using var temp = new TempDirectory();
        var path = temp.File("comprehensive.docx");
        using (var builder = DocumentBuilder.Create(path))
        {
            builder.AddRichMarkdown(markdown);
            builder.Save();
        }
        OpenXmlAssert.NoDocxValidationErrors(path);
    }

    [Fact]
    public void ReplaceWithRichMarkdown_RemovesTargetAndKeepsRelationshipsValid()
    {
        var png = TestImages.DataUriBase64(TestImages.Png(16, 16));

        using var temp = new TempDirectory();
        var path = temp.File("replace.docx");
        using (var builder = DocumentBuilder.Create(path))
        {
            builder.AddParagraph("Placeholder");
            builder.AddParagraph("TARGET to replace");
            builder.AddParagraph("Tail");
            builder.ReplaceWithRichMarkdown("TARGET", $"""
                ## Replaced

                [link](https://example.com)

                ![img]({png})

                A note[^1].

                [^1]: Definition.
                """);
            builder.Save();
        }

        using (var doc = WordprocessingDocument.Open(path, false))
        {
            var mainPart = doc.MainDocumentPart!;
            var body = mainPart.Document!.Body!;

            var paragraphs = body.Elements<W.Paragraph>().Select(p => p.InnerText).ToList();
            Assert.Equal("Placeholder", paragraphs[0]);
            Assert.Equal("Replaced", paragraphs[1]);
            Assert.Contains("link", paragraphs[2]);
            Assert.Equal(string.Empty, paragraphs[3]); // image paragraph
            Assert.Contains("A note", paragraphs[4]);
            Assert.Equal("Tail", paragraphs[5]);
            Assert.DoesNotContain("TARGET", body.InnerText);

            // Hyperlink, image, and footnote relationships all survive the replace.
            Assert.NotEmpty(mainPart.HyperlinkRelationships);
            Assert.Single(mainPart.ImageParts);
            var blip = body.Descendants<A.Blip>().Single();
            Assert.NotNull(mainPart.GetPartById(blip.Embed!.Value!));
            Assert.NotNull(mainPart.FootnotesPart);
            Assert.Single(body.Descendants<FootnoteReference>());
        }

        OpenXmlAssert.NoDocxValidationErrors(path);
    }

    [Fact]
    public void ReplaceWithRichMarkdown_MissingTarget_Throws()
    {
        using var builder = DocumentBuilder.Create();
        builder.AddParagraph("Only");

        Assert.Throws<InvalidOperationException>(() => builder.ReplaceWithRichMarkdown("Nope", "# X"));
    }

    // ---- Result / diagnostics plumbing -------------------------------------

    [Fact]
    public void AddRichMarkdown_RecordsResultAndStrictWarnings()
    {
        using var builder = DocumentBuilder.Create();
        builder.AddRichMarkdown("<div>raw</div>");
        Assert.NotNull(builder.LastRichMarkdownResult);
        Assert.False(builder.LastRichMarkdownResult!.HasWarnings);
        Assert.False(builder.LastRichMarkdownResult.HasErrors);

        builder.AddRichMarkdown("<div>raw</div>", new MarkdownRenderOptions { Strict = true });
        Assert.True(builder.LastRichMarkdownResult!.HasWarnings);
        Assert.Contains(
            builder.LastRichMarkdownResult.Diagnostics,
            d => d.Message.Contains("HTML", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void AddMarkdown_RoutesThroughRichRendering()
    {
        // The existing overload must produce rich-only output: a real hyperlink relationship.
        using var builder = DocumentBuilder.Create();
        builder.AddMarkdown("[x](https://example.com)");
        var bytes = builder.SaveToBytes();

        using var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        Assert.NotEmpty(doc.MainDocumentPart!.HyperlinkRelationships);
    }

    // ---- Helpers -----------------------------------------------------------

    private static byte[] Build(string markdown, MarkdownRenderOptions? options = null)
    {
        using var builder = (DocumentBuilder)DocumentBuilder.Create();
        builder.AddRichMarkdown(markdown, options);
        return builder.SaveToBytes();
    }

    private static WordprocessingDocument Open(byte[] bytes) =>
        WordprocessingDocument.Open(new MemoryStream(bytes), false);
}
