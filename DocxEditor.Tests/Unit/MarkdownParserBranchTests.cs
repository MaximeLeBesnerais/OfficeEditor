using DocxEditor.Core.Markdown;
using DocxEditor.Core.Models;
using OfficeEditor.Core.Models;

namespace DocxEditor.Tests.Unit;

public class MarkdownParserBranchTests
{
    [Fact]
    public void Parse_WithNullMarkdown_ShouldThrow()
    {
        var parser = new MarkdownParser();

        Assert.ThrowsAny<Exception>(() => parser.Parse(null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   \n\t\n")]
    public void Parse_WithEmptyOrWhitespaceMarkdown_ShouldReturnNoBlocks(string markdown)
    {
        var parser = new MarkdownParser();

        var blocks = parser.Parse(markdown);

        Assert.Empty(blocks);
    }

    [Fact]
    public void Parse_WithBranchHeavyMarkdown_ShouldConvertSupportedBlockTypesAndStyles()
    {
        var parser = new MarkdownParser();
        var styleMap = new StyleMapping
        {
            StyleMap = new Dictionary<string, string>
            {
                ["heading1"] = "TitleStyle",
                ["heading6"] = "TinyHeading",
                ["paragraph"] = "BodyStyle",
                ["blockquote"] = "QuoteStyle",
                ["code"] = "CodeStyle"
            }
        };
        var markdown = """
        # Title **bold**

        ###### Deep Heading

        Paragraph with **bold**, *italic*, `code`, [link text](https://example.com), and <span>html</span>.

        - Bullet one
        - Bullet two with **bold**

        1. First
        2. Second

        > Quote line one
        >
        > Quote line two

        ```csharp
        var answer = 42;
        ```

        ---

        | Name | Value |
        | ---- | ----- |
        | A    | **B** |
        | C    | `D`   |
        """;

        var blocks = parser.Parse(markdown, styleMap);

        var heading1 = Assert.IsType<HeadingBlock>(blocks[0]);
        Assert.Equal(1, heading1.Level);
        Assert.Equal("Title bold", heading1.Text);
        Assert.Equal("TitleStyle", heading1.Style);

        var heading6 = Assert.IsType<HeadingBlock>(blocks[1]);
        Assert.Equal(6, heading6.Level);
        Assert.Equal("TinyHeading", heading6.Style);

        var paragraph = Assert.IsType<ParagraphBlock>(blocks[2]);
        Assert.Equal("BodyStyle", paragraph.Style);
        Assert.Contains("Paragraph with bold, italic, code, link text, and ", paragraph.Text);
        Assert.NotNull(paragraph.InlineFormats);
        Assert.Contains(paragraph.InlineFormats, f => f.Type == "bold" && f.Text == "bold");
        Assert.Contains(paragraph.InlineFormats, f => f.Type == "italic" && f.Text == "italic");
        Assert.Contains(paragraph.InlineFormats, f => f.Type == "code" && f.Text == "code");
        Assert.Contains(paragraph.InlineFormats, f => f.Type == "text" && f.Text == "link text");

        var unordered = Assert.IsType<ListBlock>(blocks[3]);
        Assert.False(unordered.Ordered);
        Assert.Equal(new[] { "Bullet one", "Bullet two with bold" }, unordered.Items);
        Assert.Equal("BodyStyle", unordered.Style);

        var ordered = Assert.IsType<ListBlock>(blocks[4]);
        Assert.True(ordered.Ordered);
        Assert.Equal(new[] { "First", "Second" }, ordered.Items);

        var quote = Assert.IsType<BlockquoteBlock>(blocks[5]);
        Assert.Equal("QuoteStyle", quote.Style);
        Assert.Contains("Quote line one", quote.Text);
        Assert.Contains("Quote line two", quote.Text);

        var code = Assert.IsType<CodeBlock>(blocks[6]);
        Assert.Equal("csharp", code.Language);
        Assert.Equal("CodeStyle", code.Style);
        Assert.Contains("var answer = 42;", code.Text);

        Assert.IsType<HorizontalRuleBlock>(blocks[7]);

        var table = Assert.IsType<TableBlock>(blocks[8]);
        Assert.Equal(3, table.Rows.Count);
        Assert.Equal(new[] { "Name", "Value" }, table.Rows[0].Cells.Select(c => c.Text));
        Assert.Equal("B", table.Rows[1].Cells[1].Text);
        Assert.Equal("D", table.Rows[2].Cells[1].Text);
    }

    [Fact]
    public void Parse_WithMalformedMarkdown_ShouldPreserveLiteralTextWhereNoSyntaxCloses()
    {
        var parser = new MarkdownParser();

        var blocks = parser.Parse("###Heading without space\n\nUnclosed **bold and [link](");

        Assert.Collection(blocks,
            first => Assert.Equal("###Heading without space", Assert.IsType<ParagraphBlock>(first).Text),
            second =>
            {
                var paragraph = Assert.IsType<ParagraphBlock>(second);
                Assert.Contains("Unclosed", paragraph.Text);
                Assert.Contains("**bold", paragraph.Text);
            });
    }

    [Fact]
    public void Parse_WithIndentedCodeAndUnfencedBackticks_ShouldCoverCodeBranches()
    {
        var parser = new MarkdownParser();

        var blocks = parser.Parse("    indented code\n    second line\n\n```\nno language\n```");

        var indented = Assert.IsType<CodeBlock>(blocks[0]);
        Assert.Null(indented.Language);
        Assert.Contains("indented code", indented.Text);

        var fenced = Assert.IsType<CodeBlock>(blocks[1]);
        Assert.True(fenced.Language is null or "");
        Assert.Contains("no language", fenced.Text);
    }
}
