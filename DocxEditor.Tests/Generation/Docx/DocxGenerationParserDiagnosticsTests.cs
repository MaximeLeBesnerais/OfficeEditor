using DocxEditor.Core.Generation.Schema;

namespace DocxEditor.Tests.Generation.Docx;

/// <summary>
/// Strict parser diagnostics: unknown properties, value kinds, required fields, enum/color
/// validation, typo hints and JSON paths. The parser must report every contract violation in
/// one pass (loud, actionable, path-prefixed) instead of failing fast.
/// </summary>
public class DocxGenerationParserDiagnosticsTests
{
    [Fact]
    public void ValidMinimalDocument_ParsesWithoutErrorsOrWarnings()
    {
        var result = DocxTestHarness.Validate(DocxTestHarness.MinimalJson);

        Assert.True(result.IsValid);
        Assert.Empty(result.Errors);
        Assert.Empty(result.Warnings);
        Assert.NotNull(result.Document);
        Assert.Equal("1.0", result.Document!.Version);
        Assert.Single(result.Document.Sections);
    }

    [Fact]
    public void UnknownRootProperty_IsRejectedWithPath()
    {
        var result = DocxTestHarness.Validate("""
            { "version": "1.0", "sectionz": [] }
            """);

        Assert.False(result.IsValid);
        var issue = Assert.Single(result.Errors, e => e.Path == "$.sectionz");
        Assert.Contains("unknown property 'sectionz'", issue.Message);
        Assert.Contains("Did you mean 'sections'?", issue.Suggestion);
    }

    [Fact]
    public void UnknownSectionProperty_IsRejectedWithIndexedPath()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "sections": [ { "blockz": [] } ]
            }
            """);

        var issue = Assert.Single(result.Errors, e => e.Path == "$.sections[0].blockz");
        Assert.Contains("unknown property 'blockz'", issue.Message);
    }

    [Fact]
    public void UnknownParagraphProperty_IsRejectedWithNestedPath()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "sections": [
                { "blocks": [ { "type": "paragraph", "text": "hi", "spacingg": 12 } ] }
              ]
            }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.sections[0].blocks[0].spacingg", issue.Path);
        Assert.Contains("unknown property 'spacingg'", issue.Message);
        Assert.Contains("Did you mean 'spacing'?", issue.Suggestion);
    }

    [Theory]
    [InlineData("pagesize", "page size is declared per section")]
    [InlineData("margin", "use 'margins'")]
    [InlineData("column", "use 'columns'")]
    [InlineData("break", "use 'breakType'")]
    [InlineData("templatepath", "use 'template'")]
    [InlineData("listtype", "use 'kind'")]
    [InlineData("alignement", "use 'alignment'")]
    [InlineData("aligment", "use 'alignment'")]
    public void DocxSpecificTypos_ProduceHint(string typo, string hint)
    {
        var result = DocxTestHarness.Validate($$"""
            {
              "version": "1.0",
              "sections": [
                { "blocks": [ { "type": "paragraph", "text": "hi", "{{typo}}": "x" } ] }
              ]
            }
            """);

        Assert.False(result.IsValid);
        var issue = Assert.Single(result.Errors);
        Assert.Contains(hint, issue.Message);
    }

    [Theory]
    [InlineData("version", "123")]
    [InlineData("version", "null")]
    public void WrongValueKind_ReportsTypeError(string property, string value)
    {
        var result = DocxTestHarness.Validate($$"""
            {
              "{{property}}": {{value}},
              "sections": [ { "blocks": [ { "type": "paragraph", "text": "hi" } ] } ]
            }
            """);

        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, e => e.Path == "$.version" && e.Message.Contains("must be a string"));
    }

    [Fact]
    public void SectionsNotAnArray_IsRejected()
    {
        var result = DocxTestHarness.Validate("""
            { "version": "1.0", "sections": {} }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.sections", issue.Path);
        Assert.Contains("must be an array of section objects", issue.Message);
    }

    [Fact]
    public void UnsupportedVersion_IsRejected()
    {
        var result = DocxTestHarness.Validate("""
            { "version": "9.9", "sections": [ { "blocks": [ { "type": "paragraph", "text": "hi" } ] } ] }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.version", issue.Path);
        Assert.Contains("unsupported version '9.9'", issue.Message);
    }

    [Fact]
    public void MissingSections_IsRejected()
    {
        var result = DocxTestHarness.Validate("""{ "version": "1.0" }""");

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$", issue.Path);
        Assert.Contains("'sections' is required", issue.Message);
    }

    [Fact]
    public void EmptySectionsArray_IsRejected()
    {
        var result = DocxTestHarness.Validate("""{ "version": "1.0", "sections": [] }""");

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.sections", issue.Path);
        Assert.Contains("at least one section is required", issue.Message);
    }

    [Fact]
    public void MissingBlocks_IsRejected()
    {
        var result = DocxTestHarness.Validate("""
            { "version": "1.0", "sections": [ {} ] }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.sections[0]", issue.Path);
        Assert.Contains("'blocks' is required", issue.Message);
    }

    [Fact]
    public void EmptyBlocksArray_IsRejected()
    {
        var result = DocxTestHarness.Validate("""
            { "version": "1.0", "sections": [ { "blocks": [] } ] }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.sections[0].blocks", issue.Path);
        Assert.Contains("at least one flow block", issue.Message);
    }

    [Fact]
    public void BlockWithoutType_IsRejected()
    {
        var result = DocxTestHarness.Validate("""
            { "version": "1.0", "sections": [ { "blocks": [ { "text": "hi" } ] } ] }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.sections[0].blocks[0]", issue.Path);
        Assert.Contains("'type' is required", issue.Message);
    }

    [Fact]
    public void UnknownBlockType_ReportsValidValuesAndSuggestion()
    {
        var result = DocxTestHarness.Validate("""
            { "version": "1.0", "sections": [ { "blocks": [ { "type": "paragraf", "text": "hi" } ] } ] }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.sections[0].blocks[0]", issue.Path);
        Assert.Contains("unknown flow block type 'paragraf'", issue.Message);
        Assert.Contains("paragraph", issue.Message);
        Assert.Contains("Did you mean 'paragraph'?", issue.Suggestion);
    }

    [Fact]
    public void UnknownPositionedType_ReportsValidValues()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "sections": [ {
                "blocks": [ { "type": "paragraph", "text": "hi" } ],
                "positioned": [ { "type": "box", "x": 0, "y": 0, "width": 10, "height": 10 } ]
              } ]
            }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.sections[0].positioned[0]", issue.Path);
        Assert.Contains("unknown positioned element type 'box'", issue.Message);
        Assert.Contains("textBox|image|rect|line|callout", issue.Message);
    }

    [Fact]
    public void InvalidEnumValue_ReportsValidValues()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "sections": [ { "blocks": [ { "type": "image", "src": "data:image/png;base64,AAAA", "fit": "contian" } ] } ]
            }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.sections[0].blocks[0].fit", issue.Path);
        Assert.Contains("'contian' is not a valid image fit mode", issue.Message);
        Assert.Contains("fill, contain, crop, stretch", issue.Message);
        Assert.Contains("Did you mean 'contain'?", issue.Suggestion);
    }

    [Fact]
    public void RequiredImageSource_IsEnforced()
    {
        var result = DocxTestHarness.Validate("""
            { "version": "1.0", "sections": [ { "blocks": [ { "type": "image" } ] } ] }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.sections[0].blocks[0]", issue.Path);
        Assert.Contains("'src' is required", issue.Message);
    }

    [Fact]
    public void UnknownTypographyToken_IsRejectedWithSuggestion()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "design": { "palette": { "ink": "#1F2937" }, "typography": { "body": { "size": 11 } } },
              "sections": [ { "blocks": [ { "type": "paragraph", "text": "hi", "token": "bodyz" } ] } ]
            }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.sections[0].blocks[0].token", issue.Path);
        Assert.Contains("unknown typography token 'bodyz'", issue.Message);
        Assert.Contains("Did you mean 'body'?", issue.Suggestion);
    }

    [Fact]
    public void UnknownSpacingToken_IsRejected()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "design": { "palette": { "ink": "#1F2937" }, "spacing": { "normal": 8 } },
              "sections": [ { "blocks": [ { "type": "paragraph", "text": "hi", "spacing": { "before": "bogus" } } ] } ]
            }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.sections[0].blocks[0].spacing.before", issue.Path);
        Assert.Contains("unknown spacing token 'bogus'", issue.Message);
    }

    [Theory]
    [InlineData("#GGHHII", "'#GGHHII' is not a valid hex color")]
    [InlineData("notacolor", "unknown color 'notacolor'")]
    public void InvalidColor_IsRejected(string color, string expected)
    {
        var result = DocxTestHarness.Validate($$"""
            {
              "version": "1.0",
              "design": { "palette": { "ink": "#1F2937" } },
              "sections": [ { "blocks": [ { "type": "paragraph", "runs": [ { "text": "hi", "color": "{{color}}" } ] } ] } ]
            }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.sections[0].blocks[0].runs[0].color", issue.Path);
        Assert.Contains(expected, issue.Message);
    }

    [Fact]
    public void PaletteEntry_NotHex_IsRejected()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "design": { "palette": { "ink": "blue" } },
              "sections": [ { "blocks": [ { "type": "paragraph", "text": "hi" } ] } ]
            }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.design.palette.ink", issue.Path);
        Assert.Contains("palette colors must be #RRGGBB hex literals", issue.Message);
    }

    [Fact]
    public void OffPaletteRawHex_ProducesWarningNotError()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "design": { "palette": { "ink": "#1F2937" } },
              "sections": [ { "blocks": [ { "type": "paragraph", "runs": [ { "text": "hi", "color": "#123ABC" } ] } ] } ]
            }
            """);

        Assert.True(result.IsValid);
        var warning = Assert.Single(result.Warnings);
        Assert.Equal("$.sections[0].blocks[0].runs[0].color", warning.Path);
        Assert.Contains("off-palette", warning.Message);
    }

    [Fact]
    public void HeadinglevelOutOfRange_IsRejected()
    {
        var result = DocxTestHarness.Validate("""
            { "version": "1.0", "sections": [ { "blocks": [ { "type": "heading", "level": 7, "text": "x" } ] } ] }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.sections[0].blocks[0].level", issue.Path);
        Assert.Contains("must be between 1 and 6", issue.Message);
    }

    [Fact]
    public void RaggedTableRows_AreRejected()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "sections": [ { "blocks": [
                { "type": "table", "rows": [
                  { "cells": [ { "text": "a" }, { "text": "b" } ] },
                  { "cells": [ { "text": "c" } ] }
                ] }
              ] } ]
            }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.sections[0].blocks[0].rows[1]", issue.Path);
        Assert.Contains("row has 1 cells but the table has 2 columns", issue.Message);
    }

    [Fact]
    public void WidthsMismatchColumnCount_IsRejected()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "sections": [ { "blocks": [
                { "type": "table", "widths": [ 100 ], "rows": [
                  { "cells": [ { "text": "a" }, { "text": "b" } ] }
                ] }
              ] } ]
            }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.sections[0].blocks[0].widths", issue.Path);
        Assert.Contains("got 1 column widths but the table has 2 columns", issue.Message);
    }

    [Fact]
    public void FirstSectionBreakType_Warns()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "sections": [
                { "pageSetup": { "breakType": "nextPage" }, "blocks": [ { "type": "paragraph", "text": "a" } ] }
              ]
            }
            """);

        Assert.True(result.IsValid);
        var warning = Assert.Single(result.Warnings);
        Assert.Equal("$.sections[0].pageSetup.breakType", warning.Path);
        Assert.Contains("ignored", warning.Message);
    }

    [Fact]
    public void LineWithHeight_IsRejected()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "sections": [ {
                "blocks": [ { "type": "paragraph", "text": "hi" } ],
                "positioned": [ { "type": "line", "x": 0, "y": 0, "width": 100, "height": 50 } ]
              } ]
            }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.sections[0].positioned[0].height", issue.Path);
        Assert.Contains("a line has no height", issue.Message);
    }

    [Fact]
    public void MarginsExceedingPage_AreRejected()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "sections": [ {
                "pageSetup": { "size": "a4", "margins": { "top": 450, "right": 10, "bottom": 450, "left": 10 } },
                "blocks": [ { "type": "paragraph", "text": "hi" } ]
              } ]
            }
            """);

        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, e => e.Path == "$.sections[0].pageSetup" && e.Message.Contains("margins exceed the page height"));
    }

    [Fact]
    public void CropRemovingWholeImage_IsRejected()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "sections": [ { "blocks": [
                { "type": "image", "src": "data:image/png;base64,AAAA", "crop": { "left": 0.6, "right": 0.6 } }
              ] } ]
            }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.sections[0].blocks[0].crop", issue.Path);
        Assert.Contains("left + right crop", issue.Message);
    }

    [Fact]
    public void MultipleErrors_AreCollectedInOnePass()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "2.0",
              "sections": [
                { "blocks": [ { "type": "wat", "text": "hi" }, { "type": "paragraph", "text": 5 } ] }
              ]
            }
            """);

        Assert.False(result.IsValid);
        Assert.True(result.Errors.Count >= 3);
        Assert.Contains(result.Errors, e => e.Path == "$.version");
        Assert.Contains(result.Errors, e => e.Path == "$.sections[0].blocks[0]");
        Assert.Contains(result.Errors, e => e.Path == "$.sections[0].blocks[1].text");
    }

    [Fact]
    public void MalformedJson_ReportsSingleRootError()
    {
        var result = DocxTestHarness.Validate("{ not json ");

        Assert.False(result.IsValid);
        var issue = Assert.Single(result.Errors);
        Assert.Equal("$", issue.Path);
        Assert.Contains("malformed JSON", issue.Message);
    }

    [Fact]
    public void NonObjectRoot_IsRejected()
    {
        var result = DocxTestHarness.Validate("[1, 2, 3]");

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$", issue.Path);
        Assert.Contains("root must be a JSON object", issue.Message);
    }

    [Fact]
    public void Parse_ThrowsValidationExceptionWithAllIssues()
    {
        var json = """
            {
              "version": "2.0",
              "sections": [ { "blocks": [ { "type": "paragraph" } ] } ]
            }
            """;

        var exception = Assert.Throws<DocxGenerationValidationException>(() => DocxTestHarness.Parse(json));
        Assert.True(exception.Issues.Count >= 2);
        Assert.Contains(exception.Issues, i => i.Path == "$.version");
        Assert.Contains(exception.Issues, i => i.Path == "$.sections[0].blocks[0]");
        Assert.Contains("Invalid DOCX generation document", exception.Message);
        Assert.Contains("$.version", exception.Message);
    }

    [Fact]
    public void ParserInstance_IsStatelessAndReusable()
    {
        var parser = new DocxGenerationDocumentParser();

        Assert.True(parser.Validate(DocxTestHarness.MinimalJson).IsValid);
        Assert.False(parser.Validate("{ not json").IsValid);
        Assert.True(parser.Validate(DocxTestHarness.MinimalJson).IsValid);
        Assert.NotNull(parser.Parse(DocxTestHarness.MinimalJson));
    }

    [Fact]
    public void Issue_ToString_IncludesPathAndSuggestion()
    {
        var issue = new DocxGenerationIssue("$.sectionz", "boom", "Did you mean 'sections'?", DocxGenerationIssueSeverity.Error);
        Assert.Equal("$.sectionz: boom Did you mean 'sections'?", issue.ToString());

        var bare = new DocxGenerationIssue("$.x", "boom", null, DocxGenerationIssueSeverity.Warning);
        Assert.Equal("$.x: boom", bare.ToString());
    }

    [Fact]
    public void EmptyTemplatePath_IsRejected()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "template": "",
              "sections": [ { "blocks": [ { "type": "paragraph", "text": "hi" } ] } ]
            }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.template", issue.Path);
        Assert.Contains("must not be empty", issue.Message);
    }

    [Fact]
    public void TextAndRunsTogether_AreRejected()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "sections": [ { "blocks": [
                { "type": "paragraph", "text": "hi", "runs": [ { "text": "x" } ] }
              ] } ]
            }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.sections[0].blocks[0]", issue.Path);
        Assert.Contains("exactly one of 'text' (string) or 'runs' (array)", issue.Message);
    }

    [Fact]
    public void BulletListWithStart_Warns()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "sections": [ { "blocks": [
                { "type": "list", "kind": "bullet", "start": 3, "items": [ "a" ] }
              ] } ]
            }
            """);

        Assert.True(result.IsValid);
        var warning = Assert.Single(result.Warnings);
        Assert.Equal("$.sections[0].blocks[0].start", warning.Path);
        Assert.Contains("ignored for bullet lists", warning.Message);
    }
}
