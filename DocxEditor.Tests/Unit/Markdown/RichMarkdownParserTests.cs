using System.Text;
using DocxEditor.Core.Markdown;
using DocxEditor.Core.Markdown.Model;

namespace DocxEditor.Tests.Unit.Markdown;

public class RichMarkdownParserTests
{
    private static RichMarkdownParser Parser() => new();

    private static RichMarkdownParser StrictParser() =>
        new(new MarkdownParseOptions { Strict = true });

    // ---- Basic / lifecycle ----------------------------------------------

    [Fact]
    public void Parse_NullMarkdown_Throws()
    {
        var parser = Parser();

        Assert.Throws<ArgumentNullException>(() => parser.Parse(null!));
    }

    [Fact]
    public void Parse_EmptyMarkdown_ReturnsEmptyDocument()
    {
        var result = Parser().Parse(string.Empty);

        Assert.Empty(result.Document.Blocks);
        Assert.Empty(result.Diagnostics);
    }

    [Fact]
    public void ParseDocument_StaticHelper_ParsesWithDefaultOptions()
    {
        var result = RichMarkdownParser.ParseDocument("# Hello");

        var heading = Assert.IsType<MarkdownHeading>(Assert.Single(result.Document.Blocks));
        Assert.Equal(1, heading.Level);
    }

    // ---- Nested emphasis ------------------------------------------------

    [Fact]
    public void Parse_NestedEmphasis_KeepsNesting()
    {
        var paragraph = ParagraphOf("**bold _italic_**");

        var bold = Assert.IsType<MarkdownEmphasis>(paragraph.Inlines[0]);
        Assert.Equal(EmphasisKind.Bold, bold.Kind);
        Assert.Equal("bold ", Assert.IsType<MarkdownText>(bold.Children[0]).Text);

        var italic = Assert.IsType<MarkdownEmphasis>(bold.Children[1]);
        Assert.Equal(EmphasisKind.Italic, italic.Kind);
        Assert.Equal("italic", Assert.IsType<MarkdownText>(italic.Children[0]).Text);
    }

    [Fact]
    public void Parse_TripleEmphasis_ProducesItalicInsideBold()
    {
        var paragraph = ParagraphOf("***both***");

        var italic = Assert.IsType<MarkdownEmphasis>(paragraph.Inlines[0]);
        Assert.Equal(EmphasisKind.Italic, italic.Kind);

        var bold = Assert.IsType<MarkdownEmphasis>(italic.Children[0]);
        Assert.Equal(EmphasisKind.Bold, bold.Kind);
        Assert.Equal("both", Assert.IsType<MarkdownText>(bold.Children[0]).Text);
    }

    [Fact]
    public void Parse_EmphasisKinds_MapsEachDelimiterToItsKind()
    {
        var paragraph = ParagraphOf("**b** *i* ~~s~~ ~sub~ ^sup^ ++ins++ ==mark==");

        var kinds = paragraph.Inlines.OfType<MarkdownEmphasis>().Select(e => e.Kind).ToList();
        Assert.Equal(
            new[]
            {
                EmphasisKind.Bold,
                EmphasisKind.Italic,
                EmphasisKind.Strikethrough,
                EmphasisKind.Subscript,
                EmphasisKind.Superscript,
                EmphasisKind.Inserted,
                EmphasisKind.Marked
            },
            kinds);
    }

    [Fact]
    public void Parse_UnderscoreEmphasis_IsBoldOrItalic()
    {
        var paragraph = ParagraphOf("__strong__ _em_");

        Assert.Equal(EmphasisKind.Bold, ((MarkdownEmphasis)paragraph.Inlines[0]).Kind);
        Assert.Equal(EmphasisKind.Italic, ((MarkdownEmphasis)paragraph.Inlines[2]).Kind);
    }

    // ---- Links and images -----------------------------------------------

    [Fact]
    public void Parse_InlineLink_RetainsUrlTitleAndText()
    {
        var paragraph = ParagraphOf("[GitHub](https://github.com \"home\")");

        var link = Assert.IsType<MarkdownLink>(Assert.Single(paragraph.Inlines));
        Assert.Equal("https://github.com", link.Url);
        Assert.Equal("home", link.Title);
        Assert.Equal("GitHub", PlainText(link.Children));
    }

    [Fact]
    public void Parse_Autolink_IsMarkedAsAutoLink()
    {
        var paragraph = ParagraphOf("Visit <https://example.com> or www.example.org");

        Assert.IsType<MarkdownText>(paragraph.Inlines[0]);

        var angle = Assert.IsType<MarkdownLink>(paragraph.Inlines[1]);
        Assert.True(angle.IsAutoLink);
        Assert.Equal("https://example.com", angle.Url);

        var bare = Assert.IsType<MarkdownLink>(paragraph.Inlines[3]);
        Assert.True(bare.IsAutoLink);
        Assert.Equal("http://www.example.org", bare.Url);
    }

    [Fact]
    public void Parse_Image_RetainsUrlTitleAndAltText()
    {
        var paragraph = ParagraphOf("![alt text](img.png \"tooltip\")");

        var image = Assert.IsType<MarkdownImage>(Assert.Single(paragraph.Inlines));
        Assert.Equal("img.png", image.Url);
        Assert.Equal("tooltip", image.Title);
        Assert.Equal("alt text", PlainText(image.Children));
    }

    [Fact]
    public void Parse_ReferenceLink_ResolvesUrlAndRetainsDefinition()
    {
        var result = Parser().Parse("[text][id]\n\n[id]: http://example.com \"A title\"\n");

        var paragraph = Assert.IsType<MarkdownParagraph>(result.Document.Blocks[0]);
        var link = Assert.IsType<MarkdownLink>(Assert.Single(paragraph.Inlines));
        Assert.Equal("http://example.com", link.Url);
        Assert.Equal("text", PlainText(link.Children));

        var definitions = Assert.IsType<MarkdownLinkReferenceDefinitions>(result.Document.Blocks[1]);
        var definition = Assert.Single(definitions.Definitions);
        Assert.Equal("id", definition.Label);
        Assert.Equal("http://example.com", definition.Url);
        Assert.Equal("A title", definition.Title);
    }

    [Fact]
    public void Parse_ShortcutReference_IsMarkedAsShortcut()
    {
        var result = Parser().Parse("[text]\n\n[text]: http://x.org");
        var paragraph = Assert.IsType<MarkdownParagraph>(result.Document.Blocks[0]);
        var link = Assert.IsType<MarkdownLink>(Assert.Single(paragraph.Inlines));

        Assert.True(link.IsShortcut);
        Assert.Equal("http://x.org", link.Url);
    }

    // ---- Lists, nesting, starts, tasks ----------------------------------

    [Fact]
    public void Parse_NestedList_PreservesHierarchy()
    {
        var result = Parser().Parse("""
        - item one
          - nested one
          - nested two
        - item two
        """);

        var list = Assert.IsType<MarkdownList>(result.Document.Blocks[0]);
        Assert.False(list.Ordered);
        Assert.Equal('-', list.BulletType);
        Assert.Equal(2, list.Items.Count);

        var first = list.Items[0];
        var firstParagraph = Assert.IsType<MarkdownParagraph>(first.Blocks[0]);
        Assert.Equal("item one", PlainText(firstParagraph.Inlines));

        var nested = Assert.IsType<MarkdownList>(first.Blocks[1]);
        Assert.Equal(2, nested.Items.Count);
        Assert.Equal("nested one", PlainText(Assert.IsType<MarkdownParagraph>(nested.Items[0].Blocks[0]).Inlines));
        Assert.Equal("nested two", PlainText(Assert.IsType<MarkdownParagraph>(nested.Items[1].Blocks[0]).Inlines));
    }

    [Fact]
    public void Parse_StartedOrderedList_PreservesStartAndItemOrder()
    {
        var result = Parser().Parse("3. three\n4. four\n");

        var list = Assert.IsType<MarkdownList>(result.Document.Blocks[0]);
        Assert.True(list.Ordered);
        Assert.Equal("3", list.OrderedStart);
        Assert.Equal(3, list.Items[0].Order);
        Assert.Equal(4, list.Items[1].Order);
    }

    [Fact]
    public void Parse_TaskList_PreservesCheckedState()
    {
        var result = Parser().Parse("- [x] done\n- [ ] todo\n");

        var list = Assert.IsType<MarkdownList>(result.Document.Blocks[0]);
        var done = Assert.IsType<MarkdownParagraph>(list.Items[0].Blocks[0]);
        var doneCheckbox = Assert.IsType<MarkdownTaskCheckbox>(done.Inlines[0]);
        Assert.True(doneCheckbox.Checked);

        var todo = Assert.IsType<MarkdownParagraph>(list.Items[1].Blocks[0]);
        var todoCheckbox = Assert.IsType<MarkdownTaskCheckbox>(todo.Inlines[0]);
        Assert.False(todoCheckbox.Checked);
    }

    // ---- Tables ----------------------------------------------------------

    [Fact]
    public void Parse_PipeTable_RetainsHeaderAlignmentAndFormattedCells()
    {
        var result = Parser().Parse("""
        | Left | Center | Right |
        |:-----|:------:|------:|
        | a    | b      | **c** |
        """);

        var table = Assert.IsType<MarkdownTable>(result.Document.Blocks[0]);
        Assert.Equal(3, table.Columns.Count);
        Assert.Equal(MarkdownTableAlignment.Left, table.Columns[0].Alignment);
        Assert.Equal(MarkdownTableAlignment.Center, table.Columns[1].Alignment);
        Assert.Equal(MarkdownTableAlignment.Right, table.Columns[2].Alignment);
        Assert.Equal(0f, table.Columns[0].Width);

        Assert.Equal(2, table.Rows.Count);

        var header = table.Rows[0];
        Assert.True(header.IsHeader);
        Assert.Equal("Left", PlainText(Assert.IsType<MarkdownParagraph>(header.Cells[0].Blocks[0]).Inlines));

        var data = table.Rows[1];
        Assert.False(data.IsHeader);
        Assert.Equal("a", PlainText(Assert.IsType<MarkdownParagraph>(data.Cells[0].Blocks[0]).Inlines));

        var formatted = Assert.IsType<MarkdownParagraph>(data.Cells[2].Blocks[0]);
        var emphasis = Assert.IsType<MarkdownEmphasis>(formatted.Inlines[0]);
        Assert.Equal(EmphasisKind.Bold, emphasis.Kind);
        Assert.Equal("c", PlainText(emphasis.Children));
    }

    [Fact]
    public void Parse_GridTable_RetainsRowsAndCells()
    {
        var result = Parser().Parse("+---+---+\n| A | B |\n+===+---+\n| 1 | 2 |\n+---+---+\n");

        var table = Assert.IsType<MarkdownTable>(result.Document.Blocks[0]);
        Assert.Equal(2, table.Rows.Count);
        Assert.True(table.Rows[0].IsHeader);
        Assert.Equal(2, table.Rows[0].Cells.Count);
        Assert.Equal("A", PlainText(Assert.IsType<MarkdownParagraph>(table.Rows[0].Cells[0].Blocks[0]).Inlines));
        Assert.Equal("1", PlainText(Assert.IsType<MarkdownParagraph>(table.Rows[1].Cells[0].Blocks[0]).Inlines));
    }

    // ---- Breaks and entities ---------------------------------------------

    [Fact]
    public void Parse_HardBreak_PreservesHardFlag()
    {
        var paragraph = ParagraphOf("line  \nhard");

        Assert.Equal(3, paragraph.Inlines.Count);
        var lineBreak = Assert.IsType<MarkdownLineBreak>(paragraph.Inlines[1]);
        Assert.True(lineBreak.IsHard);
        Assert.False(lineBreak.IsBackslash);
    }

    [Fact]
    public void Parse_SoftBreak_IsNotHard()
    {
        var paragraph = ParagraphOf("soft\nbreak");

        var lineBreak = Assert.IsType<MarkdownLineBreak>(paragraph.Inlines[1]);
        Assert.False(lineBreak.IsHard);
    }

    [Fact]
    public void Parse_BackslashBreak_IsHardAndBackslash()
    {
        var paragraph = ParagraphOf("back\\\nslash");

        var lineBreak = Assert.IsType<MarkdownLineBreak>(paragraph.Inlines[1]);
        Assert.True(lineBreak.IsHard);
        Assert.True(lineBreak.IsBackslash);
    }

    [Fact]
    public void Parse_Entities_AreDecodedSafely()
    {
        var paragraph = ParagraphOf("a &amp; b &copy;");

        var first = Assert.IsType<MarkdownEntity>(paragraph.Inlines[1]);
        Assert.Equal("&amp;", first.Original);
        Assert.Equal("&", first.Decoded);

        var second = Assert.IsType<MarkdownEntity>(paragraph.Inlines[3]);
        Assert.Equal("&copy;", second.Original);
        Assert.Equal("©", second.Decoded);
    }

    [Fact]
    public void Parse_UnknownEntity_FallsBackToLiteralText()
    {
        var paragraph = ParagraphOf("nope &notanentity; done");

        Assert.Contains(paragraph.Inlines, i => i is MarkdownText t && t.Text.Contains("&notanentity;", StringComparison.Ordinal));
    }

    // ---- Code ------------------------------------------------------------

    [Fact]
    public void Parse_FencedCode_PreservesNewlinesAndIndentation()
    {
        var result = Parser().Parse("```csharp\nvar x = 1;\n  indented;\nlast\n```");

        var code = Assert.IsType<MarkdownCodeBlock>(result.Document.Blocks[0]);
        Assert.Equal("csharp", code.Language);
        Assert.Equal("var x = 1;\n  indented;\nlast", code.Text);
    }

    [Fact]
    public void Parse_IndentedCode_HasNoLanguage()
    {
        var result = Parser().Parse("    indented code\n    second line");

        var code = Assert.IsType<MarkdownCodeBlock>(result.Document.Blocks[0]);
        Assert.Null(code.Language);
        Assert.Equal("indented code\nsecond line", code.Text);
    }

    [Fact]
    public void Parse_InlineCode_PreservesWhitespaceAndInnerBackticks()
    {
        var paragraph = ParagraphOf("Use `code` and ``two ` ticks`` and `  padded  `");

        var simple = Assert.IsType<MarkdownCode>(paragraph.Inlines[1]);
        Assert.Equal("code", simple.Content);

        var inner = Assert.IsType<MarkdownCode>(paragraph.Inlines[3]);
        Assert.Equal("two ` ticks", inner.Content);
        Assert.Equal(2, inner.DelimiterCount);

        var padded = Assert.IsType<MarkdownCode>(paragraph.Inlines[5]);
        Assert.Equal(" padded ", padded.Content);
    }

    // ---- Footnotes -------------------------------------------------------

    [Fact]
    public void Parse_Footnotes_RetainsReferenceAndDefinition()
    {
        var result = Parser().Parse("A ref[^1].\n\n[^1]: The note with *emphasis*.\n");

        var paragraph = Assert.IsType<MarkdownParagraph>(result.Document.Blocks[0]);
        var reference = paragraph.Inlines.OfType<MarkdownFootnoteReference>().Single();
        Assert.Equal("^1", reference.Label);
        Assert.Equal(1, reference.Index);
        Assert.False(reference.IsBackLink);

        var footnotes = Assert.IsType<MarkdownFootnotesBlock>(result.Document.Blocks.OfType<MarkdownFootnotesBlock>().Single());
        var footnote = Assert.Single(footnotes.Footnotes);
        Assert.Equal("^1", footnote.Label);
        Assert.Equal(1, footnote.Order);

        var note = Assert.IsType<MarkdownParagraph>(footnote.Blocks[0]);
        Assert.Contains(note.Inlines, i => i is MarkdownText t && t.Text.Contains("The note with", StringComparison.Ordinal));
        Assert.Contains(note.Inlines, i => i is MarkdownEmphasis e && e.Kind == EmphasisKind.Italic);
        Assert.Contains(note.Inlines, i => i is MarkdownFootnoteReference b && b.IsBackLink && b.Index == 1);
    }

    // ---- Definition lists ------------------------------------------------

    [Fact]
    public void Parse_DefinitionList_RetainsTermsAndDefinitions()
    {
        var result = Parser().Parse("Term 1\n:   Definition one\n:   Definition two\n");

        var list = Assert.IsType<MarkdownDefinitionList>(result.Document.Blocks[0]);
        Assert.Equal(2, list.Items.Count);

        var first = list.Items[0];
        var term = Assert.Single(first.Terms);
        Assert.Equal("Term 1", PlainText(term.Inlines));
        var definition = Assert.IsType<MarkdownParagraph>(Assert.Single(first.Definitions));
        Assert.Equal("Definition one", PlainText(definition.Inlines));

        var second = list.Items[1];
        Assert.Equal("Definition two", PlainText(Assert.IsType<MarkdownParagraph>(Assert.Single(second.Definitions)).Inlines));
    }

    // ---- YAML front matter ----------------------------------------------

    [Fact]
    public void Parse_YamlFrontMatter_IsRetainedVerbatim()
    {
        var result = Parser().Parse("---\ntitle: Test\ntags: [a, b]\n---\n# Heading\n");

        var yaml = Assert.IsType<MarkdownYamlFrontMatter>(result.Document.Blocks[0]);
        Assert.Equal("title: Test\ntags: [a, b]", yaml.Yaml);

        var heading = Assert.IsType<MarkdownHeading>(result.Document.Blocks[1]);
        Assert.Equal(1, heading.Level);
    }

    // ---- Emoji -----------------------------------------------------------

    [Fact]
    public void Parse_EmojiShortcodes_AreDecoded()
    {
        var paragraph = ParagraphOf(":smile: and :rocket:");

        var smile = Assert.IsType<MarkdownEmoji>(paragraph.Inlines[0]);
        Assert.Equal(":smile:", smile.Match);
        Assert.Equal("😄", smile.Text);

        var rocket = Assert.IsType<MarkdownEmoji>(paragraph.Inlines[2]);
        Assert.Equal(":rocket:", rocket.Match);
        Assert.Equal("🚀", rocket.Text);
    }

    // ---- HTML and unknown constructs / diagnostics ----------------------

    [Fact]
    public void Parse_HtmlBlock_IsRetainedAsExplicitNode()
    {
        var result = Parser().Parse("<div>\nraw\n</div>");

        var html = Assert.IsType<MarkdownHtmlBlock>(result.Document.Blocks[0]);
        Assert.Contains("raw", html.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void Parse_HtmlInline_IsRetainedAsExplicitNode()
    {
        var paragraph = ParagraphOf("a <span>b</span> c");

        Assert.IsType<MarkdownHtml>(paragraph.Inlines[1]);
        Assert.IsType<MarkdownHtml>(paragraph.Inlines[3]);
    }

    [Fact]
    public void Parse_StrictMode_ReportsHtmlWarnings()
    {
        var result = StrictParser().Parse("<div>\nraw\n</div>\n\npara <span>inline</span>\n");

        Assert.True(result.HasWarnings);
        Assert.Contains(result.Diagnostics, d => d.Severity == MarkdownDiagnosticSeverity.Warning && d.Message.Contains("HTML", StringComparison.OrdinalIgnoreCase));
        Assert.IsType<MarkdownHtmlBlock>(result.Document.Blocks[0]);
    }

    [Fact]
    public void Parse_PermissiveMode_HasNoDiagnosticsForRetainedHtml()
    {
        var result = Parser().Parse("<div>\nraw\n</div>\n\npara <span>inline</span>\n");

        Assert.Empty(result.Diagnostics);
        Assert.IsType<MarkdownHtmlBlock>(result.Document.Blocks[0]);
    }

    [Fact]
    public void Parse_StrictMode_RetainsUnknownInlineMathWithWarning()
    {
        var result = StrictParser().Parse("inline $x^2$ math\n");

        var paragraph = Assert.IsType<MarkdownParagraph>(result.Document.Blocks[0]);
        var unknown = paragraph.Inlines.OfType<MarkdownUnknownInline>().Single();
        Assert.Equal("MathInline", unknown.Kind);
        Assert.Contains(unknown.Raw, "$x^2$", StringComparison.Ordinal);
        Assert.Contains(result.Diagnostics, d => d.NodeKind == nameof(MarkdownUnknownInline));
    }

    [Fact]
    public void Parse_UnknownBlock_IsRetainedVerbatum()
    {
        var result = Parser().Parse(":::note\ncontent\n:::\n");

        var unknown = Assert.IsType<MarkdownUnknownBlock>(result.Document.Blocks[0]);
        Assert.Equal("CustomContainer", unknown.Kind);
        Assert.Contains("content", unknown.Raw, StringComparison.Ordinal);
    }

    [Fact]
    public void Parse_UnmatchedDelimiter_FlattensToLiteralText()
    {
        var result = StrictParser().Parse("an unclosed [link\n");

        var paragraph = Assert.IsType<MarkdownParagraph>(result.Document.Blocks[0]);
        var text = PlainText(paragraph.Inlines);
        Assert.Contains("unclosed", text, StringComparison.Ordinal);
        Assert.Contains("link", text, StringComparison.Ordinal);
        Assert.Contains(result.Diagnostics, d => d.Message.Contains("Unmatched delimiter", StringComparison.Ordinal));
    }

    // ---- Source spans -----------------------------------------------------

    [Fact]
    public void Parse_SourceSpans_ExposeCharacterOffsetsAndLines()
    {
        var result = Parser().Parse("# Title\n\nBody **bold**.\n");

        var heading = Assert.IsType<MarkdownHeading>(result.Document.Blocks[0]);
        Assert.NotNull(heading.Span);
        Assert.Equal(0, heading.Span!.Value.Start);
        Assert.Equal(6, heading.Span.Value.End);
        Assert.Equal(0, heading.Span.Value.Line);

        var paragraph = Assert.IsType<MarkdownParagraph>(result.Document.Blocks[1]);
        var emphasis = Assert.IsType<MarkdownEmphasis>(paragraph.Inlines[1]);
        Assert.NotNull(emphasis.Span);
        Assert.True(emphasis.Span!.Value.Start > 0);
    }

    [Fact]
    public void Parse_DisabledSourceSpans_ReturnNoSpans()
    {
        var options = new MarkdownParseOptions { UseSourceSpans = false };
        var result = new RichMarkdownParser(options).Parse("# Title\n\nBody\n");

        Assert.All(result.Document.Blocks, b => Assert.Null(b.Span));
    }

    // ---- Helpers ----------------------------------------------------------

    private static MarkdownParagraph ParagraphOf(string markdown)
    {
        var result = Parser().Parse(markdown);
        return Assert.IsType<MarkdownParagraph>(Assert.Single(result.Document.Blocks));
    }

    private static string PlainText(IEnumerable<MarkdownInline> inlines)
    {
        var sb = new StringBuilder();
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case MarkdownText t:
                    sb.Append(t.Text);
                    break;
                case MarkdownEntity e:
                    sb.Append(e.Decoded);
                    break;
                case MarkdownEmoji em:
                    sb.Append(em.Text);
                    break;
                case MarkdownCode c:
                    sb.Append(c.Content);
                    break;
                case MarkdownLineBreak br:
                    sb.Append('\n');
                    break;
                case MarkdownEmphasis m:
                    sb.Append(PlainText(m.Children));
                    break;
                case MarkdownLink l:
                    sb.Append(PlainText(l.Children));
                    break;
                case MarkdownImage img:
                    sb.Append(PlainText(img.Children));
                    break;
                case MarkdownFootnoteReference:
                    break;
            }
        }

        return sb.ToString();
    }
}
