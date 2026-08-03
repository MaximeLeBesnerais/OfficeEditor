using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Generation;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

namespace DocxEditor.Tests.Generation;

/// <summary>
/// Emitted OOXML coverage for the flow emitters: keep-next/keep-lines/widow control, density
/// spacing, semantic role styles, theme table treatment (light borders, padding, header fill,
/// banding, row-split prevention, repeating headers) and the body-size/table-width guardrails.
/// </summary>
public class DocxFlowEmitTests
{
    [Fact]
    public void Heading_EmitsKeepNextKeepLinesOutlineAndThemeSpacing()
    {
        var json = """
            {"version":"1.0","sections":[{"blocks":[{"type":"heading","level":1,"text":"Overview"}]}]}
            """;

        using var doc = Open(Generate(json).Content);
        var paragraph = BodyParagraphs(doc).Single();
        var properties = paragraph.ParagraphProperties!;

        Assert.Equal("Heading1", properties.ParagraphStyleId!.Val!.Value);
        Assert.Equal(0, properties.OutlineLevel!.Val!.Value);
        Assert.NotNull(properties.KeepNext);
        Assert.NotNull(properties.KeepLines);
        Assert.Null(properties.WidowControl);

        // Theme heading1 spacing: before 22pt (440 twips), after 6pt (120 twips).
        Assert.Equal("440", properties.SpacingBetweenLines!.Before!.Value);
        Assert.Equal("120", properties.SpacingBetweenLines.After!.Value);

        // Theme heading1 run defaults: Georgia 20pt navy.
        var run = paragraph.Elements<Wp.Run>().Single();
        Assert.Equal("Georgia", run.RunProperties!.RunFonts!.Ascii!.Value);
        Assert.Equal("40", run.RunProperties.FontSize!.Val!.Value);
        Assert.Equal("1F3A5F", run.RunProperties.Color!.Val!.Value);
    }

    [Fact]
    public void BodyParagraph_EmitsWidowControlAndDensitySpacing()
    {
        var json = """
            {"version":"1.0","sections":[{"blocks":[{"type":"paragraph","text":"Body copy"}]}]}
            """;

        var generated = Generate(json);
        using var doc = Open(generated.Content);
        var paragraph = BodyParagraphs(doc).Single();
        var properties = paragraph.ParagraphProperties!;

        Assert.NotNull(properties.WidowControl);
        // Theme body spacing: after 8pt (160 twips), line 1.3 (312).
        Assert.Equal("160", properties.SpacingBetweenLines!.After!.Value);
        Assert.Equal("312", properties.SpacingBetweenLines.Line!.Value);
        Assert.Equal(LineSpacingRuleValues.Auto, properties.SpacingBetweenLines.LineRule!.Value);

        // The document defaults carry the editorial body font (Arial) — runs inherit it.
        var defaults = doc.MainDocumentPart!.StyleDefinitionsPart!.Styles!.DocDefaults!
            .RunPropertiesDefault!.RunPropertiesBaseStyle!;
        Assert.Equal("Arial", defaults.RunFonts!.Ascii!.Value);
    }

    [Fact]
    public void Role_Subtitle_AppliesThemeFormattingAndStyle()
    {
        var json = """
            {"version":"1.0","sections":[{"blocks":[{"type":"paragraph","role":"subtitle","text":"Under the title"}]}]}
            """;

        using var doc = Open(Generate(json).Content);
        var paragraph = BodyParagraphs(doc).Single();
        var properties = paragraph.ParagraphProperties!;

        Assert.Equal("Subtitle", properties.ParagraphStyleId!.Val!.Value);
        // Subtitle theme: Arial 13pt muted, after 14pt (280 twips).
        Assert.Equal("280", properties.SpacingBetweenLines!.After!.Value);
        var run = paragraph.Elements<Wp.Run>().Single();
        Assert.Equal("Arial", run.RunProperties!.RunFonts!.Ascii!.Value);
        Assert.Equal("26", run.RunProperties.FontSize!.Val!.Value);
        Assert.Equal("5A6B7B", run.RunProperties.Color!.Val!.Value);
    }

    [Fact]
    public void Role_Eyebrow_EmitsAllCapsAndAccent()
    {
        var json = """
            {"version":"1.0","sections":[{"blocks":[{"type":"paragraph","role":"eyebrow","text":"QUARTERLY"}]}]}
            """;

        using var doc = Open(Generate(json).Content);
        var paragraph = BodyParagraphs(doc).Single();
        var run = paragraph.Elements<Wp.Run>().Single();

        Assert.NotNull(run.RunProperties!.Caps);
        Assert.Equal("E4674A", run.RunProperties.Color!.Val!.Value);
    }

    [Fact]
    public void Role_Metric_AppliesLargeDisplayNumber()
    {
        var json = """
            {"version":"1.0","sections":[{"blocks":[{"type":"paragraph","role":"metric","text":"+34%"}]}]}
            """;

        using var doc = Open(Generate(json).Content);
        var paragraph = BodyParagraphs(doc).Single();
        var run = paragraph.Elements<Wp.Run>().Single();

        Assert.Equal("Metric", paragraph.ParagraphProperties!.ParagraphStyleId!.Val!.Value);
        Assert.Equal("Georgia", run.RunProperties!.RunFonts!.Ascii!.Value);
        Assert.Equal("48", run.RunProperties.FontSize!.Val!.Value); // 24pt
        Assert.Equal("1F3A5F", run.RunProperties.Color!.Val!.Value);
    }

    [Fact]
    public void ExplicitToken_OverridesRoleDefaults()
    {
        var json = """
            {"version":"1.0","design":{"typography":{"custom":{"font":"body","size":14,"color":"teal"}}},
             "sections":[{"blocks":[{"type":"paragraph","role":"body","token":"custom","text":"Overridden"}]}]}
            """;

        using var doc = Open(Generate(json).Content);
        var paragraph = BodyParagraphs(doc).Single();
        var run = paragraph.Elements<Wp.Run>().Single();

        // Token wins over the body role's 11pt ink: 14pt teal.
        Assert.Equal("28", run.RunProperties!.FontSize!.Val!.Value);
        Assert.Equal("1F7A6E", run.RunProperties.Color!.Val!.Value);
    }

    [Fact]
    public void Table_EmitsThemeBordersPaddingHeaderBandingAndCantSplit()
    {
        var json = """
            {"version":"1.0","sections":[{"blocks":[{"type":"table","widths":[100,100],"rows":[
              {"header":true,"cells":[{"text":"Name"},{"text":"Value"}]},
              {"cells":[{"text":"A"},{"text":"1"}]},
              {"cells":[{"text":"B"},{"text":"2"}]}
            ]}]}]}
            """;

        using var doc = Open(Generate(json).Content);
        var table = BodyTable(doc);
        var tableProperties = table.TableProperties!;

        // Theme-aware light borders.
        var borders = tableProperties.TableBorders!;
        Assert.Equal("D6DEE6", borders.TopBorder!.Color!.Value);
        Assert.Equal("D6DEE6", borders.InsideVerticalBorder!.Color!.Value);

        // Cell padding.
        Assert.NotNull(tableProperties.TableCellMarginDefault);

        var rows = table.Elements<Wp.TableRow>().ToList();
        Assert.Equal(3, rows.Count);

        // Header row: repeat + no split + theme fill + header style.
        var headerRowProperties = rows[0].TableRowProperties!;
        Assert.NotNull(headerRowProperties.GetFirstChild<TableHeader>());
        Assert.NotNull(headerRowProperties.GetFirstChild<CantSplit>());
        var headerCell = rows[0].Elements<Wp.TableCell>().First();
        Assert.Equal("1F3A5F", headerCell.TableCellProperties!.Shading!.Fill!.Value);
        Assert.Equal("TableHeader", HeaderCellParagraph(headerCell).ParagraphProperties!.ParagraphStyleId!.Val!.Value);

        // Body rows: no split, first body row banded pale, header/body style.
        Assert.NotNull(rows[1].TableRowProperties!.GetFirstChild<CantSplit>());
        var bandedCell = rows[1].Elements<Wp.TableCell>().First();
        Assert.Equal("EEF2F6", bandedCell.TableCellProperties!.Shading!.Fill!.Value);
        Assert.Equal("TableBody", bandedCell.Elements<Wp.Paragraph>().Single().ParagraphProperties!.ParagraphStyleId!.Val!.Value);

        // Second body row: no banding (alternate).
        var secondBodyCell = rows[2].Elements<Wp.TableCell>().First();
        Assert.Null(secondBodyCell.TableCellProperties!.Shading);
    }

    [Fact]
    public void ExplicitCellFill_OverridesThemeBand()
    {
        var json = """
            {"version":"1.0","sections":[{"blocks":[{"type":"table","rows":[
              {"cells":[{"text":"A","fill":"white"}]},
              {"cells":[{"text":"B","fill":"#123456"}]}
            ]}]}]}
            """;

        using var doc = Open(Generate(json).Content);
        var table = BodyTable(doc);
        var rows = table.Elements<Wp.TableRow>().ToList();

        Assert.Equal("FFFFFF", rows[0].Elements<Wp.TableCell>().Single().TableCellProperties!.Shading!.Fill!.Value);
        Assert.Equal("123456", rows[1].Elements<Wp.TableCell>().Single().TableCellProperties!.Shading!.Fill!.Value);
    }

    [Fact]
    public void Table_HeaderRow_EmitsCantSplitBeforeTableHeaderAndValidates()
    {
        // CT_TrPr requires cantSplit to precede tblHeader; emitting them the other way
        // around produces schema-invalid OOXML that Word opens "with repair".
        var json = """
            {"version":"1.0","sections":[{"blocks":[{"type":"table","widths":[100,100],"rows":[
              {"header":true,"cells":[{"text":"Name"},{"text":"Value"}]},
              {"cells":[{"text":"A"},{"text":"1"}]}
            ]}]}]}
            """;

        using var doc = Open(Generate(json).Content);
        var headerRow = BodyTable(doc).Elements<Wp.TableRow>().First();
        var children = headerRow.TableRowProperties!.ChildElements.ToList();
        var cantSplitIndex = children.FindIndex(c => c is CantSplit);
        var headerIndex = children.FindIndex(c => c is TableHeader);

        Assert.True(cantSplitIndex >= 0, "header row must carry cantSplit");
        Assert.True(headerIndex >= 0, "header row must carry tblHeader");
        Assert.True(cantSplitIndex < headerIndex,
            $"expected cantSplit (index {cantSplitIndex}) before tblHeader (index {headerIndex})");

        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void Callout_StyleUsesBareHexForColorAndFill()
    {
        var json = """
            {"version":"1.0","sections":[{"blocks":[{"type":"paragraph","role":"callout","text":"Callout text"}]}]}
            """;

        using var doc = Open(Generate(json).Content);
        var callout = doc.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<Wp.Style>().Single(s => s.StyleId?.Value == "Callout");
        var paragraphProperties = callout.StyleParagraphProperties!;

        // w:color / w:fill reject a leading '#': the editorial primary (#1F3A5F) and pale (#EEF2F6)
        // must be emitted as bare hex.
        var color = paragraphProperties.ParagraphBorders!.LeftBorder!.Color!.Value!;
        Assert.Equal("1F3A5F", color);
        Assert.False(color.StartsWith('#'));

        var fill = paragraphProperties.Shading!.Fill!.Value!;
        Assert.Equal("EEF2F6", fill);
        Assert.False(fill.StartsWith('#'));
    }

    [Fact]
    public void DesignFontOverrides_ApplyToSemanticRolesAndHeadings()
    {
        var json = """
            {"version":"1.0","design":{"fonts":{"display":"Times New Roman","body":"Verdana"}},
             "sections":[{"blocks":[
               {"type":"heading","level":1,"text":"Overview"},
               {"type":"paragraph","role":"subtitle","text":"Under the title"}
             ]}]}
            """;

        using var doc = Open(Generate(json).Content);
        var paragraphs = BodyParagraphs(doc);

        // The heading resolves its "display" slot to the document font, not the theme's Georgia.
        var headingRun = paragraphs[0].Elements<Wp.Run>().Single();
        Assert.Equal("Times New Roman", headingRun.RunProperties!.RunFonts!.Ascii!.Value);

        // The subtitle resolves its "body" slot to the document font, not the theme's Arial.
        var subtitleRun = paragraphs[1].Elements<Wp.Run>().Single();
        Assert.Equal("Verdana", subtitleRun.RunProperties!.RunFonts!.Ascii!.Value);
    }

    [Fact]
    public void UndersizedBodyRun_WarnsWithoutPaginationPrediction()
    {
        var json = """
            {"version":"1.0","sections":[{"blocks":[{"type":"paragraph","runs":[{"text":"tiny","size":7}]}]}]}
            """;

        var generated = Generate(json);

        var warning = Assert.Single(generated.Result.Warnings);
        Assert.Contains("minimum body size", warning.Message);
        Assert.Contains("7", warning.Message);
    }

    [Fact]
    public void OversizedTable_Warns()
    {
        var json = """
            {"version":"1.0","sections":[{"blocks":[{"type":"table","widths":[400,400],"rows":[
              {"cells":[{"text":"A"},{"text":"B"}]}
            ]}]}]}
            """;

        var generated = Generate(json);

        var warning = Assert.Single(generated.Result.Warnings);
        Assert.Contains("exceeds the recommended maximum", warning.Message);
    }

    [Fact]
    public void FitTable_DoesNotWarn()
    {
        var json = """
            {"version":"1.0","sections":[{"blocks":[{"type":"table","widths":[100,100],"rows":[
              {"cells":[{"text":"A"},{"text":"B"}]}
            ]}]}]}
            """;

        var generated = Generate(json);

        Assert.DoesNotContain(generated.Result.Warnings, w => w.Message.Contains("table width"));
    }

    [Fact]
    public void FullEditorialDocument_IsSchemaValidWithNoWarnings()
    {
        var json = """
            {"version":"1.0","sections":[{"blocks":[
              {"type":"paragraph","role":"eyebrow","text":"NORTHWIND / OPERATIONS"},
              {"type":"heading","level":1,"text":"Operations Brief"},
              {"type":"paragraph","role":"subtitle","text":"Quarterly review of the supply chain"},
              {"type":"paragraph","text":"Body copy with enough text to read comfortably."},
              {"type":"paragraph","role":"muted","text":"Supporting detail rendered in muted ink."},
              {"type":"paragraph","role":"label","text":"FIELD LABEL"},
              {"type":"paragraph","role":"metric","text":"+34%"},
              {"type":"paragraph","role":"metricLabel","text":"Revenue growth"},
              {"type":"table","widths":[120,120],"rows":[
                {"header":true,"cells":[{"text":"Workstream","role":"tableHeader"},{"text":"Signal","role":"tableHeader"}]},
                {"cells":[{"text":"Inventory","role":"tableBody"},{"text":"Stable","role":"tableBody"}]}
              ]},
              {"type":"callout","tone":"tip","text":"Tip: keep body text at least 9pt."},
              {"type":"paragraph","role":"footer","text":"Confidential"}
            ]}]}
            """;

        var generated = Generate(json);
        Assert.Empty(generated.Result.Warnings);

        using var doc = Open(generated.Content);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    // ---- helpers ----

    private static GeneratedDocx Generate(string json) => new DocxGenerator().GenerateToBytes(json);

    private static WordprocessingDocument Open(byte[] bytes)
    {
        var stream = new MemoryStream(bytes);
        return WordprocessingDocument.Open(stream, false);
    }

    private static Wp.Paragraph HeaderCellParagraph(Wp.TableCell cell) =>
        cell.Elements<Wp.Paragraph>().Single();

    private static List<Wp.Paragraph> BodyParagraphs(WordprocessingDocument doc) =>
        doc.MainDocumentPart!.Document!.Body!.Descendants<Wp.Paragraph>().ToList();

    private static Wp.Table BodyTable(WordprocessingDocument doc) =>
        doc.MainDocumentPart!.Document!.Body!.Descendants<Wp.Table>().Single();
}
