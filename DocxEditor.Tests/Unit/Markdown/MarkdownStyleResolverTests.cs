using DocxEditor.Core.Markdown;
using DocxEditor.Core.Markdown.Model;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeEditor.Core.Models;

namespace DocxEditor.Tests.Unit.Markdown;

public class MarkdownStyleResolverTests
{
    private static Style ParagraphStyle(string id, string name) =>
        new(new StyleName { Val = name }) { Type = StyleValues.Paragraph, StyleId = id };

    private static Style CharacterStyle(string id, string name) =>
        new(new StyleName { Val = name }) { Type = StyleValues.Character, StyleId = id };

    private static Style TableStyle(string id, string name) =>
        new(new StyleName { Val = name }) { Type = StyleValues.Table, StyleId = id };

    // ---- resolution against user-defined styles -------------------------

    [Fact]
    public void Resolve_ExactStyleId_ReturnsMatchingStyle()
    {
        var styles = new Styles(
            ParagraphStyle("CustomPara", "Custom Para"),
            CharacterStyle("CustomChar", "Custom Char"),
            TableStyle("CustomTable", "Custom Table"));

        var resolver = new MarkdownStyleResolver(styles);

        var para = resolver.Resolve("CustomPara", MarkdownStyleKind.Paragraph);
        Assert.True(para.Resolved);
        Assert.Equal("CustomPara", para.StyleId);
        Assert.True(para.HasStyle);
        Assert.Empty(para.Diagnostics);

        var charStyle = resolver.Resolve("CustomChar", MarkdownStyleKind.Character);
        Assert.True(charStyle.Resolved);
        Assert.Equal("CustomChar", charStyle.StyleId);

        var table = resolver.Resolve("CustomTable", MarkdownStyleKind.Table);
        Assert.True(table.Resolved);
        Assert.Equal("CustomTable", table.StyleId);
    }

    [Fact]
    public void Resolve_VisibleNameWithSpaces_ResolvesToStyleId()
    {
        var styles = new Styles(
            ParagraphStyle("CustomPara", "Custom Para"),
            CharacterStyle("CustomChar", "Custom Char"));

        var resolver = new MarkdownStyleResolver(styles);

        Assert.Equal("CustomPara", resolver.Resolve("Custom Para", MarkdownStyleKind.Paragraph).StyleId);
        Assert.Equal("CustomChar", resolver.Resolve("Custom Char", MarkdownStyleKind.Character).StyleId);
    }

    [Fact]
    public void Resolve_UniqueCaseInsensitiveName_ResolvesToStyleId()
    {
        var styles = new Styles(ParagraphStyle("CustomPara", "Custom Para"));

        var resolver = new MarkdownStyleResolver(styles);

        Assert.Equal("CustomPara", resolver.Resolve("custom para", MarkdownStyleKind.Paragraph).StyleId);
    }

    [Fact]
    public void Resolve_StyleIdTakesPrecedenceOverStyleName()
    {
        // "Note" is both a StyleId and the visible name of another style.
        var styles = new Styles(
            ParagraphStyle("Note", "Note Paragraph"),
            ParagraphStyle("Other", "Note"));

        var resolver = new MarkdownStyleResolver(styles);

        var resolution = resolver.Resolve("Note", MarkdownStyleKind.Paragraph);
        Assert.True(resolution.Resolved);
        Assert.Equal("Note", resolution.StyleId);
    }

    [Fact]
    public void Resolve_ExactNameTakesPrecedenceOverCaseInsensitiveMatch()
    {
        var styles = new Styles(
            ParagraphStyle("Exact", "Body"),
            ParagraphStyle("CaseInsensitive", "body"));

        var resolver = new MarkdownStyleResolver(styles);

        // "Body" matches the first style's name exactly; "body" the second's exactly.
        Assert.Equal("Exact", resolver.Resolve("Body", MarkdownStyleKind.Paragraph).StyleId);
        Assert.Equal("CaseInsensitive", resolver.Resolve("body", MarkdownStyleKind.Paragraph).StyleId);
    }

    // ---- ambiguity ------------------------------------------------------

    [Fact]
    public void Resolve_AmbiguousCaseInsensitiveName_StrictRejectsWithError()
    {
        var styles = new Styles(
            ParagraphStyle("A", "Body"),
            ParagraphStyle("B", "body"));

        var resolver = new MarkdownStyleResolver(styles, MarkdownStyleResolverOptions.StrictMode);

        // "BODY" matches neither name exactly, but both case-insensitively.
        var resolution = resolver.Resolve("BODY", MarkdownStyleKind.Paragraph, element: "paragraph", path: "blocks[0]");

        Assert.False(resolution.Resolved);
        Assert.Null(resolution.StyleId);
        Assert.Null(resolution.FallbackStyle);
        var diagnostic = Assert.Single(resolution.Diagnostics);
        Assert.Equal(MarkdownDiagnosticSeverity.Error, diagnostic.Severity);
        Assert.Contains("ambiguous", diagnostic.Message);
        Assert.Equal("paragraph", diagnostic.Element);
        Assert.Equal("blocks[0]", diagnostic.Path);
    }

    [Fact]
    public void Resolve_AmbiguousCaseInsensitiveName_PermissiveFallsBackWithError()
    {
        var styles = new Styles(
            ParagraphStyle("A", "Body"),
            ParagraphStyle("B", "body"));

        var resolver = new MarkdownStyleResolver(styles);

        var resolution = resolver.Resolve("BODY", MarkdownStyleKind.Paragraph, element: "paragraph");

        Assert.False(resolution.Resolved);
        Assert.NotNull(resolution.StyleId);
        Assert.NotNull(resolution.FallbackStyle);
        Assert.Equal(StyleValues.Paragraph, resolution.FallbackStyle.Type?.Value);
        var diagnostic = Assert.Single(resolution.Diagnostics);
        Assert.Equal(MarkdownDiagnosticSeverity.Warning, diagnostic.Severity);
    }

    [Fact]
    public void Resolve_DuplicateExactNames_Rejected()
    {
        var styles = new Styles(
            ParagraphStyle("A", "Duplicate"),
            ParagraphStyle("B", "Duplicate"));

        var resolver = new MarkdownStyleResolver(styles, MarkdownStyleResolverOptions.StrictMode);

        var resolution = resolver.Resolve("Duplicate", MarkdownStyleKind.Paragraph);
        Assert.False(resolution.Resolved);
        Assert.Null(resolution.StyleId);
        Assert.Null(resolution.FallbackStyle);
        Assert.Equal(MarkdownDiagnosticSeverity.Error, Assert.Single(resolution.Diagnostics).Severity);
    }

    // ---- wrong kind -----------------------------------------------------

    [Fact]
    public void Resolve_WrongKind_StrictRejectsWithError()
    {
        var styles = new Styles(CharacterStyle("InlineCode", "Inline Code"));

        var resolver = new MarkdownStyleResolver(styles, MarkdownStyleResolverOptions.StrictMode);

        var resolution = resolver.Resolve("InlineCode", MarkdownStyleKind.Paragraph, element: "paragraph");

        Assert.False(resolution.Resolved);
        Assert.Null(resolution.StyleId);
        Assert.Equal(MarkdownDiagnosticSeverity.Error, Assert.Single(resolution.Diagnostics).Severity);
    }

    [Fact]
    public void Resolve_WrongKind_PermissiveGeneratesFallbackOfExpectedKind()
    {
        // The template defines "Inline Code" as a character style, but a paragraph
        // context resolves it: permissive mode must not emit the character style as a
        // paragraph style, so it generates a paragraph fallback instead.
        var styles = new Styles(CharacterStyle("InlineCode", "Inline Code"));

        var resolver = new MarkdownStyleResolver(styles);

        var resolution = resolver.Resolve("InlineCode", MarkdownStyleKind.Paragraph, element: "paragraph");

        Assert.False(resolution.Resolved);
        Assert.NotNull(resolution.FallbackStyle);
        Assert.Equal(StyleValues.Paragraph, resolution.FallbackStyle.Type?.Value);
        Assert.Equal(MarkdownDiagnosticSeverity.Warning, Assert.Single(resolution.Diagnostics).Severity);
    }

    // ---- fallback creation types ----------------------------------------

    [Fact]
    public void Resolve_UnresolvedInlineCode_GeneratesCharacterFallback()
    {
        var resolver = new MarkdownStyleResolver(styles: null);

        var resolution = resolver.ResolveElement("codeInline", "CodeChar", path: "blocks[0].inlines[2]");

        Assert.False(resolution.Resolved);
        Assert.NotNull(resolution.StyleId);
        Assert.Equal(MarkdownStyleKind.Character, resolution.Kind);
        Assert.Equal(StyleValues.Character, resolution.FallbackStyle?.Type?.Value);
        Assert.Equal("CodeChar", resolution.FallbackStyle?.StyleName?.Val?.Value);
        Assert.Equal(MarkdownDiagnosticSeverity.Warning, Assert.Single(resolution.Diagnostics).Severity);
        Assert.Equal("codeInline", Assert.Single(resolution.Diagnostics).Element);
        Assert.Equal("blocks[0].inlines[2]", Assert.Single(resolution.Diagnostics).Path);
    }

    [Fact]
    public void Resolve_UnresolvedCodeBlock_GeneratesParagraphFallback()
    {
        var resolver = new MarkdownStyleResolver(styles: null);

        var resolution = resolver.ResolveElement("codeBlock", "Code");

        Assert.Equal(MarkdownStyleKind.Paragraph, resolution.Kind);
        Assert.Equal(StyleValues.Paragraph, resolution.FallbackStyle?.Type?.Value);
        Assert.Equal("Code", resolution.FallbackStyle?.StyleName?.Val?.Value);
    }

    [Fact]
    public void Resolve_UnresolvedHyperlink_GeneratesCharacterFallback()
    {
        var resolver = new MarkdownStyleResolver(styles: null);

        var resolution = resolver.ResolveElement("hyperlink", "Hyperlink");

        Assert.Equal(MarkdownStyleKind.Character, resolution.Kind);
        Assert.Equal(StyleValues.Character, resolution.FallbackStyle?.Type?.Value);
    }

    [Fact]
    public void Resolve_UnresolvedTable_GeneratesTableFallback()
    {
        var resolver = new MarkdownStyleResolver(styles: null);

        var resolution = resolver.ResolveElement("table", "TableGrid");

        Assert.Equal(MarkdownStyleKind.Table, resolution.Kind);
        Assert.Equal(StyleValues.Table, resolution.FallbackStyle?.Type?.Value);
    }

    [Fact]
    public void Resolve_FallbackIdsAreCollisionFreeAcrossCalls()
    {
        var styles = new Styles(ParagraphStyle("Code", "Code"));

        var resolver = new MarkdownStyleResolver(styles);

        // Existing "Code" is a paragraph style; an inline-code request must not reuse it.
        var block = resolver.ResolveElement("codeBlock", "Code");
        Assert.Equal("Code", block.StyleId);

        var inline = resolver.ResolveElement("codeInline", "Code");
        Assert.Equal(StyleValues.Character, inline.FallbackStyle?.Type?.Value);
        Assert.StartsWith("Code", inline.FallbackStyle!.StyleId!.Value, StringComparison.Ordinal);
        Assert.NotEqual("Code", inline.FallbackStyle!.StyleId!.Value);
        Assert.Equal(inline.FallbackStyle.StyleId.Value, inline.StyleId);
    }

    // ---- fallback formatting -------------------------------------------

    [Fact]
    public void Fallback_Headings_HaveDistinctSizesAndOutlineLevels()
    {
        var resolver = new MarkdownStyleResolver(styles: null);

        var headings = Enumerable.Range(1, 6)
            .Select(level => resolver.ResolveElement($"heading{level}", $"Heading{level}"))
            .ToList();

        foreach (var (resolution, index) in headings.Select((r, i) => (r, i)))
        {
            var style = resolution.FallbackStyle!;
            Assert.Equal(StyleValues.Paragraph, style.Type?.Value);
            Assert.NotNull(style.StyleRunProperties!.Bold);
            Assert.NotNull(style.StyleRunProperties.FontSize);
            Assert.NotNull(style.StyleParagraphProperties!.KeepNext);
            var outline = style.StyleParagraphProperties.OutlineLevel!;
            Assert.Equal(index, outline.Val!.Value);
        }

        var sizes = headings
            .Select(h => h.FallbackStyle!.StyleRunProperties!.FontSize!.Val!.Value)
            .Distinct()
            .ToList();
        Assert.Equal(6, sizes.Count);

        var before = headings
            .Select(h => h.FallbackStyle!.StyleParagraphProperties!.SpacingBetweenLines!.Before!.Value)
            .Distinct()
            .ToList();
        Assert.Equal(6, before.Count);
    }

    [Fact]
    public void Fallback_Normal_HasReadableBodyDefaults()
    {
        var resolver = new MarkdownStyleResolver(styles: null);

        var style = resolver.ResolveElement("paragraph", "Normal").FallbackStyle!;

        Assert.Equal("22", style.StyleRunProperties!.FontSize!.Val!.Value);
        Assert.Equal("Calibri", style.StyleRunProperties.RunFonts!.Ascii!.Value);
        var spacing = style.StyleParagraphProperties!.SpacingBetweenLines!;
        Assert.Equal("160", spacing.After!.Value);
        Assert.Equal("276", spacing.Line!.Value);
        Assert.Equal(LineSpacingRuleValues.Auto, spacing.LineRule!.Value);
    }

    [Fact]
    public void Fallback_Quote_HasIndentItalicAndLeftBorder()
    {
        var resolver = new MarkdownStyleResolver(styles: null);

        var style = resolver.ResolveElement("blockquote", "Quote").FallbackStyle!;

        Assert.Equal("720", style.StyleParagraphProperties!.Indentation!.Left!.Value);
        Assert.NotNull(style.StyleRunProperties!.Italic);
        var leftBorder = style.StyleParagraphProperties.ParagraphBorders!.LeftBorder!;
        Assert.Equal(BorderValues.Single, leftBorder.Val!.Value);
    }

    [Fact]
    public void Fallback_CodeBlock_HasMonospaceShadingAndTightSpacing()
    {
        var resolver = new MarkdownStyleResolver(styles: null);

        var style = resolver.ResolveElement("codeBlock", "Code").FallbackStyle!;

        Assert.Equal("Consolas", style.StyleRunProperties!.RunFonts!.Ascii!.Value);
        Assert.NotNull(style.StyleParagraphProperties!.Shading);
        var spacing = style.StyleParagraphProperties.SpacingBetweenLines!;
        Assert.Equal("0", spacing.After!.Value);
        Assert.Equal("240", spacing.Line!.Value);
    }

    [Fact]
    public void Fallback_CodeChar_HasMonospaceAndShading()
    {
        var resolver = new MarkdownStyleResolver(styles: null);

        var style = resolver.ResolveElement("codeInline", "CodeChar").FallbackStyle!;

        Assert.Equal(StyleValues.Character, style.Type?.Value);
        Assert.Equal("Consolas", style.StyleRunProperties!.RunFonts!.Ascii!.Value);
        Assert.NotNull(style.StyleRunProperties.Shading);
    }

    [Fact]
    public void Fallback_Hyperlink_HasBlueUnderline()
    {
        var resolver = new MarkdownStyleResolver(styles: null);

        var style = resolver.ResolveElement("hyperlink", "Hyperlink").FallbackStyle!;

        Assert.Equal(StyleValues.Character, style.Type?.Value);
        Assert.Equal("0563C1", style.StyleRunProperties!.Color!.Val!.Value);
        Assert.NotNull(style.StyleRunProperties.Underline);
    }

    [Fact]
    public void Fallback_DefinitionFootnoteListAndHeader_HaveDefaults()
    {
        var resolver = new MarkdownStyleResolver(styles: null);

        var term = resolver.ResolveElement("definitionTerm", "DefinitionTerm").FallbackStyle!;
        Assert.NotNull(term.StyleRunProperties!.Bold);

        var description = resolver.ResolveElement("definitionDescription", "DefinitionDescription").FallbackStyle!;
        Assert.Equal("720", description.StyleParagraphProperties!.Indentation!.Left!.Value);

        var footnote = resolver.ResolveElement("footnoteText", "FootnoteText").FallbackStyle!;
        Assert.Equal("18", footnote.StyleRunProperties!.FontSize!.Val!.Value);
        Assert.Equal("0", footnote.StyleParagraphProperties!.SpacingBetweenLines!.After!.Value);

        var list = resolver.ResolveElement("list", "ListParagraph").FallbackStyle!;
        Assert.Equal("80", list.StyleParagraphProperties!.SpacingBetweenLines!.After!.Value);

        var header = resolver.ResolveElement("tableHeader", "TableHeader").FallbackStyle!;
        Assert.NotNull(header.StyleRunProperties!.Bold);
        Assert.Equal(JustificationValues.Center, header.StyleParagraphProperties!.Justification!.Val!.Value);
    }

    [Fact]
    public void Fallback_TableGrid_IsTableStyleWithBorders()
    {
        var resolver = new MarkdownStyleResolver(styles: null);

        var style = resolver.ResolveElement("table", "TableGrid").FallbackStyle!;

        Assert.Equal(StyleValues.Table, style.Type?.Value);
        var borders = style.StyleTableProperties!.GetFirstChild<TableBorders>()!;
        Assert.NotNull(borders.TopBorder);
        Assert.NotNull(borders.InsideHorizontalBorder);
        Assert.NotNull(borders.InsideVerticalBorder);
    }

    [Fact]
    public void Fallback_DefaultMappingElements_ProduceValidOoxml()
    {
        var resolver = new MarkdownStyleResolver(styles: null);

        var fallbacks = StyleMapping.Default.StyleMap
            .Select(kv => resolver.ResolveElement(kv.Key, kv.Value))
            .Where(r => r.FallbackStyle is not null)
            .Select(r => r.FallbackStyle!)
            .GroupBy(s => s.StyleId!.Value)
            .Select(g => g.First())
            .ToList();

        Assert.NotEmpty(fallbacks);

        var errors = new OpenXmlValidator().Validate(new Styles(fallbacks)).ToList();
        Assert.True(errors.Count == 0, string.Join("\n", errors.Select(e => e.Description)));
    }

    [Fact]
    public void Fallback_WellKnownReferenceNames_ApplyWithoutElementKey()
    {
        var resolver = new MarkdownStyleResolver(styles: null);

        var h1 = resolver.Resolve("Heading1", MarkdownStyleKind.Paragraph).FallbackStyle!;
        Assert.Equal("48", h1.StyleRunProperties!.FontSize!.Val!.Value);

        var quote = resolver.Resolve("Quote", MarkdownStyleKind.Paragraph).FallbackStyle!;
        Assert.NotNull(quote.StyleRunProperties!.Italic);

        var code = resolver.Resolve("CodeChar", MarkdownStyleKind.Character).FallbackStyle!;
        Assert.NotNull(code.StyleRunProperties!.Shading);

        var table = resolver.Resolve("TableGrid", MarkdownStyleKind.Table).FallbackStyle!;
        Assert.Equal(StyleValues.Table, table.Type?.Value);
    }

    // ---- strict / permissive unresolved behavior ------------------------

    [Fact]
    public void Resolve_Unresolved_StrictRejectsWithError()
    {
        var resolver = new MarkdownStyleResolver(styles: null, MarkdownStyleResolverOptions.StrictMode);

        var resolution = resolver.Resolve("Missing", MarkdownStyleKind.Paragraph, element: "paragraph", path: "blocks[1]");

        Assert.False(resolution.Resolved);
        Assert.False(resolution.HasStyle);
        Assert.Null(resolution.FallbackStyle);
        var diagnostic = Assert.Single(resolution.Diagnostics);
        Assert.Equal(MarkdownDiagnosticSeverity.Error, diagnostic.Severity);
        Assert.Contains("not defined", diagnostic.Message);
        Assert.Equal("blocks[1]", diagnostic.Path);
    }

    [Fact]
    public void Resolve_Unresolved_PermissiveGeneratesFallbackWithWarning()
    {
        var resolver = new MarkdownStyleResolver(styles: null);

        var resolution = resolver.Resolve("Missing", MarkdownStyleKind.Paragraph, element: "paragraph");

        Assert.False(resolution.Resolved);
        Assert.True(resolution.HasStyle);
        Assert.NotNull(resolution.FallbackStyle);
        Assert.Equal(MarkdownDiagnosticSeverity.Warning, Assert.Single(resolution.Diagnostics).Severity);
    }

    [Fact]
    public void Resolve_NullReference_ReturnsNoStyle()
    {
        var resolver = new MarkdownStyleResolver(styles: null);

        var resolution = resolver.Resolve(null, MarkdownStyleKind.Paragraph, element: "paragraph");

        Assert.False(resolution.Resolved);
        Assert.False(resolution.HasStyle);
        Assert.Null(resolution.FallbackStyle);
        Assert.Empty(resolution.Diagnostics);
    }

    // ---- existing style preservation ------------------------------------

    [Fact]
    public void Resolve_NeverMutatesTemplateStyles()
    {
        var styles = new Styles(
            ParagraphStyle("CustomPara", "Custom Para"),
            CharacterStyle("CustomChar", "Custom Char"),
            TableStyle("CustomTable", "Custom Table"));

        var originalXml = styles.OuterXml;
        var originalCount = styles.Elements<Style>().Count();

        var resolver = new MarkdownStyleResolver(styles);

        resolver.Resolve("CustomPara", MarkdownStyleKind.Paragraph);
        resolver.Resolve("Missing", MarkdownStyleKind.Paragraph);
        resolver.ResolveElement("codeInline", "CodeChar");
        resolver.Resolve("CustomChar", MarkdownStyleKind.Character);

        // Nothing was added, removed or rewritten on the template's styles.
        Assert.Equal(originalXml, styles.OuterXml);
        Assert.Equal(originalCount, styles.Elements<Style>().Count());
    }

    [Fact]
    public void Resolve_ResolvedStyleReturnsNoFallback()
    {
        var styles = new Styles(ParagraphStyle("Normal", "Normal"));

        var resolver = new MarkdownStyleResolver(styles);

        var resolution = resolver.Resolve("Normal", MarkdownStyleKind.Paragraph);
        Assert.True(resolution.Resolved);
        Assert.Null(resolution.FallbackStyle);
        Assert.Empty(resolution.Diagnostics);
    }

    // ---- semantic key mapping -------------------------------------------

    [Fact]
    public void MarkdownStyleKinds_MapsSemanticKeysToKinds()
    {
        Assert.Equal(MarkdownStyleKind.Character, MarkdownStyleKinds.ForElement("codeInline"));
        Assert.Equal(MarkdownStyleKind.Character, MarkdownStyleKinds.ForElement("hyperlink"));
        Assert.Equal(MarkdownStyleKind.Table, MarkdownStyleKinds.ForElement("table"));
        Assert.Equal(MarkdownStyleKind.Paragraph, MarkdownStyleKinds.ForElement("heading1"));
        Assert.Equal(MarkdownStyleKind.Paragraph, MarkdownStyleKinds.ForElement("paragraph"));
        Assert.Equal(MarkdownStyleKind.Paragraph, MarkdownStyleKinds.ForElement("blockquote"));
        Assert.Equal(MarkdownStyleKind.Paragraph, MarkdownStyleKinds.ForElement("codeBlock"));
        Assert.Equal(MarkdownStyleKind.Paragraph, MarkdownStyleKinds.ForElement("tableHeader"));
        Assert.Equal(MarkdownStyleKind.Paragraph, MarkdownStyleKinds.ForElement("list"));
        Assert.Equal(MarkdownStyleKind.Paragraph, MarkdownStyleKinds.ForElement("definitionTerm"));
        Assert.Equal(MarkdownStyleKind.Paragraph, MarkdownStyleKinds.ForElement("definitionDescription"));
        Assert.Equal(MarkdownStyleKind.Paragraph, MarkdownStyleKinds.ForElement("footnoteText"));
        Assert.Equal(MarkdownStyleKind.Paragraph, MarkdownStyleKinds.ForElement("thematicBreak"));
        Assert.Equal(MarkdownStyleKind.Paragraph, MarkdownStyleKinds.ForElement("imageCaption"));
    }
}
