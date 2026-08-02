using DocxEditor.Core.Markdown;
using DocxEditor.Core.Markdown.Model;

namespace DocxEditor.Tests.Unit.Markdown;

public class MarkdownModelTests
{
    [Fact]
    public void SourceSpan_None_IsEmpty()
    {
        var none = SourceSpan.None;

        Assert.True(none.IsEmpty);
        Assert.Equal(0, none.Length);
        Assert.Equal(-1, none.Start);
    }

    [Fact]
    public void SourceSpan_Length_IsInclusiveCharacterCount()
    {
        var span = new SourceSpan(2, 5, 1);

        Assert.Equal(4, span.Length);
        Assert.False(span.IsEmpty);
    }

    [Fact]
    public void MarkdownDocument_Defaults_ToNoBlocks()
    {
        var document = new MarkdownDocument();

        Assert.Empty(document.Blocks);
        Assert.Null(document.Span);
    }

    [Fact]
    public void ParseResult_Defaults_ReportNoErrors()
    {
        var result = new MarkdownParseResult(new MarkdownDocument(), []);

        Assert.False(result.HasErrors);
        Assert.False(result.HasWarnings);
        Assert.Empty(result.Diagnostics);
    }

    [Fact]
    public void ParseResult_HasErrors_ReflectsErrorSeverity()
    {
        var result = new MarkdownParseResult(
            new MarkdownDocument(),
            [new MarkdownDiagnostic(MarkdownDiagnosticSeverity.Error, "boom")]);

        Assert.True(result.HasErrors);
        Assert.False(result.HasWarnings);
    }

    [Fact]
    public void Nodes_AreImmutableAndRecursive()
    {
        var inner = new MarkdownEmphasis
        {
            Kind = EmphasisKind.Italic,
            Children = [new MarkdownText { Text = "nested" }]
        };
        var outer = new MarkdownEmphasis
        {
            Kind = EmphasisKind.Bold,
            Children = [inner]
        };

        Assert.Single(outer.Children);
        Assert.IsType<MarkdownEmphasis>(outer.Children[0]);
        Assert.Equal("nested", ((MarkdownEmphasis)outer.Children[0]).Children.OfType<MarkdownText>().Single().Text);
    }
}
