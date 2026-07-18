using PptxEditor.Core.Models;
using PptxEditor.Core.Serialization;

namespace DocxEditor.Tests.Unit;

public class PptxJsonInstructionParserTests
{
    private readonly PptxJsonInstructionParser _parser = new();

    [Fact]
    public void Parse_ReplaceText_ReturnsInstruction()
    {
        var set = _parser.Parse("""
            { "operations": [ { "type": "replaceText", "slide": 2, "elementId": 5, "text": "Hello" } ] }
            """);

        var op = Assert.IsType<PptxReplaceTextInstruction>(Assert.Single(set.Operations));
        Assert.Equal("replaceText", op.Type);
        Assert.Equal(2, op.Slide);
        Assert.Equal(5u, op.ElementId);
        Assert.Equal("Hello", op.Text);
    }

    [Fact]
    public void Parse_ReplaceImage_WithFit_ReturnsInstruction()
    {
        var set = _parser.Parse("""
            { "operations": [ { "type": "replaceImage", "slide": 1, "elementId": 3, "image": "aGVsbG8=", "fit": "fill" } ] }
            """);

        var op = Assert.IsType<PptxReplaceImageInstruction>(Assert.Single(set.Operations));
        Assert.Equal(1, op.Slide);
        Assert.Equal(3u, op.ElementId);
        Assert.Equal("aGVsbG8=", op.Image);
        Assert.Equal("fill", op.Fit);
    }

    [Fact]
    public void Parse_ReplaceImage_WithoutFit_FitIsNull()
    {
        var set = _parser.Parse("""
            { "operations": [ { "type": "replaceImage", "slide": 1, "elementId": 3, "image": "aGVsbG8=" } ] }
            """);

        var op = Assert.IsType<PptxReplaceImageInstruction>(Assert.Single(set.Operations));
        Assert.Null(op.Fit);
    }

    [Fact]
    public void Parse_ReplaceTable_ReturnsRows()
    {
        var set = _parser.Parse("""
            { "operations": [ { "type": "replaceTable", "slide": 1, "elementId": 4, "rows": [["a","b"],["c","d"]] } ] }
            """);

        var op = Assert.IsType<PptxReplaceTableInstruction>(Assert.Single(set.Operations));
        Assert.Equal(2, op.Rows.Count);
        Assert.Equal(new List<string> { "a", "b" }, op.Rows[0]);
        Assert.Equal(new List<string> { "c", "d" }, op.Rows[1]);
    }

    [Fact]
    public void Parse_MoveSlide_ReturnsInstruction()
    {
        var set = _parser.Parse("""
            { "operations": [ { "type": "moveSlide", "from": 3, "to": 1 } ] }
            """);

        var op = Assert.IsType<PptxMoveSlideInstruction>(Assert.Single(set.Operations));
        Assert.Equal(3, op.From);
        Assert.Equal(1, op.To);
    }

    [Fact]
    public void Parse_DuplicateSlide_WithAndWithoutPosition()
    {
        var set = _parser.Parse("""
            { "operations": [
                { "type": "duplicateSlide", "slide": 2 },
                { "type": "duplicateSlide", "slide": 1, "position": 5 }
            ] }
            """);

        Assert.Equal(2, set.Operations.Count);
        var first = Assert.IsType<PptxDuplicateSlideInstruction>(set.Operations[0]);
        Assert.Equal(2, first.Slide);
        Assert.Null(first.Position);
        var second = Assert.IsType<PptxDuplicateSlideInstruction>(set.Operations[1]);
        Assert.Equal(1, second.Slide);
        Assert.Equal(5, second.Position);
    }

    [Fact]
    public void Parse_DeleteSlide_ReturnsInstruction()
    {
        var set = _parser.Parse("""
            { "operations": [ { "type": "deleteSlide", "slide": 4 } ] }
            """);

        var op = Assert.IsType<PptxDeleteSlideInstruction>(Assert.Single(set.Operations));
        Assert.Equal(4, op.Slide);
    }

    [Fact]
    public void Parse_PropertyNames_AreCaseInsensitive()
    {
        var set = _parser.Parse("""
            { "Operations": [ { "Type": "ReplaceText", "Slide": 1, "ElementId": 2, "Text": "x" } ] }
            """);

        var op = Assert.IsType<PptxReplaceTextInstruction>(Assert.Single(set.Operations));
        Assert.Equal(1, op.Slide);
        Assert.Equal(2u, op.ElementId);
        Assert.Equal("x", op.Text);
    }

    [Fact]
    public void Parse_EmptyOperations_ReturnsEmptySet()
    {
        var set = _parser.Parse("""{ "operations": [] }""");
        Assert.Empty(set.Operations);
    }

    [Theory]
    [InlineData("not json")]
    [InlineData("""{ "operations": {} }""")]
    [InlineData("""[1,2,3]""")]
    [InlineData("""{ "operations": null }""")]
    public void Parse_InvalidEnvelope_ThrowsArgumentException(string json)
    {
        Assert.Throws<ArgumentException>(() => _parser.Parse(json));
    }

    [Fact]
    public void Parse_UnknownType_ThrowsWithOpIndex()
    {
        var ex = Assert.Throws<ArgumentException>(() => _parser.Parse("""
            { "operations": [
                { "type": "deleteSlide", "slide": 1 },
                { "type": "explode", "slide": 1 }
            ] }
            """));

        Assert.Contains("operations[1]", ex.Message);
        Assert.Contains("explode", ex.Message);
    }

    [Fact]
    public void Parse_MissingRequiredField_ThrowsWithOpIndex()
    {
        var ex = Assert.Throws<ArgumentException>(() => _parser.Parse("""
            { "operations": [
                { "type": "replaceText", "slide": 1, "elementId": 2, "text": "ok" },
                { "type": "replaceText", "slide": 1, "elementId": 2 }
            ] }
            """));

        Assert.Contains("operations[1]", ex.Message);
        Assert.Contains("'text'", ex.Message);
    }

    [Theory]
    [InlineData("""{ "type": "replaceText", "slide": "one", "elementId": 2, "text": "x" }""", "slide")]
    [InlineData("""{ "type": "replaceText", "slide": 0, "elementId": 2, "text": "x" }""", "slide")]
    [InlineData("""{ "type": "replaceText", "slide": -1, "elementId": 2, "text": "x" }""", "slide")]
    [InlineData("""{ "type": "replaceText", "slide": 1.5, "elementId": 2, "text": "x" }""", "slide")]
    [InlineData("""{ "type": "replaceText", "slide": 1, "elementId": 0, "text": "x" }""", "elementId")]
    [InlineData("""{ "type": "replaceText", "slide": 1, "elementId": 2, "text": 42 }""", "text")]
    public void Parse_BadFieldType_ThrowsWithOpIndex(string operation, string field)
    {
        var ex = Assert.Throws<ArgumentException>(() => _parser.Parse(
            $$"""{ "operations": [{{operation}}] }"""));

        Assert.Contains("operations[0]", ex.Message);
        Assert.Contains($"'{field}'", ex.Message);
    }

    [Fact]
    public void Parse_BadFitValue_ThrowsWithOpIndex()
    {
        var ex = Assert.Throws<ArgumentException>(() => _parser.Parse("""
            { "operations": [ { "type": "replaceImage", "slide": 1, "elementId": 2, "image": "aA==", "fit": "squash" } ] }
            """));

        Assert.Contains("operations[0]", ex.Message);
        Assert.Contains("'fit'", ex.Message);
    }

    [Theory]
    [InlineData("""{ "type": "replaceTable", "slide": 1, "elementId": 2, "rows": [] }""")]
    [InlineData("""{ "type": "replaceTable", "slide": 1, "elementId": 2, "rows": [[]] }""")]
    [InlineData("""{ "type": "replaceTable", "slide": 1, "elementId": 2, "rows": [["a"], ["b","c"]] }""")]
    [InlineData("""{ "type": "replaceTable", "slide": 1, "elementId": 2, "rows": [["a"], [7]] }""")]
    [InlineData("""{ "type": "replaceTable", "slide": 1, "elementId": 2, "rows": "nope" }""")]
    public void Parse_BadRows_ThrowsWithOpIndex(string operation)
    {
        var ex = Assert.Throws<ArgumentException>(() => _parser.Parse(
            $$"""{ "operations": [{{operation}}] }"""));

        Assert.Contains("operations[0]", ex.Message);
        Assert.Contains("rows", ex.Message);
    }

    [Fact]
    public void Parse_DuplicateSlide_BadPosition_ThrowsWithOpIndex()
    {
        var ex = Assert.Throws<ArgumentException>(() => _parser.Parse("""
            { "operations": [ { "type": "duplicateSlide", "slide": 1, "position": 0 } ] }
            """));

        Assert.Contains("operations[0]", ex.Message);
        Assert.Contains("'position'", ex.Message);
    }

    [Fact]
    public void Parse_NonObjectOperation_ThrowsWithOpIndex()
    {
        var ex = Assert.Throws<ArgumentException>(() => _parser.Parse("""
            { "operations": [42] }
            """));

        Assert.Contains("operations[0]", ex.Message);
    }
}
