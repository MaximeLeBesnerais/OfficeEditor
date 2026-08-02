using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Generation;
using OfficeEditor.Core.Exceptions;

namespace DocxEditor.Tests.Generation.Docx;

/// <summary>
/// Flow-tier structure tests: sections, page setup and section breaks, per-section
/// headers/footers, list numbering (fresh instances + collision-free abstract definitions),
/// tables, flow callouts, page breaks, group flattening, run formatting and template append
/// with style preservation.
/// </summary>
public class DocxGenerationFlowStructureTests
{
    [Fact]
    public void MultiSection_EmitsBreakTypesAndPageGeometry()
    {
        var json = """
            {
              "version": "1.0",
              "design": { "palette": { "ink": "#1F2937" } },
              "sections": [
                { "blocks": [ { "type": "paragraph", "text": "s0" } ] },
                {
                  "pageSetup": {
                    "size": { "width": 400, "height": 600 },
                    "orientation": "landscape",
                    "margins": { "top": 30, "right": 40, "bottom": 30, "left": 40 },
                    "breakType": "continuous",
                    "columns": { "count": 2, "spacing": 12, "separator": true }
                  },
                  "blocks": [ { "type": "paragraph", "text": "s1" } ]
                },
                {
                  "pageSetup": { "size": "a5", "breakType": "oddPage" },
                  "blocks": [ { "type": "paragraph", "text": "s2" } ]
                }
              ]
            }
            """;

        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(json);
        var mainPart = DocxTestHarness.OpenMainPart(path);
        var body = mainPart.Document!.Body!;

        var intermediate = body.Elements<Paragraph>()
            .Where(p => p.ParagraphProperties?.SectionProperties is not null)
            .Select(p => p.ParagraphProperties!.SectionProperties!)
            .ToList();
        Assert.Equal(2, intermediate.Count);

        // Section 0 → section 1: continuous break; section 0 geometry is default A4 portrait.
        var s0 = intermediate[0];
        Assert.Equal(SectionMarkValues.Continuous, s0.Elements<SectionType>().Single().Val!.Value);
        Assert.Equal("11906", s0.Elements<DocumentFormat.OpenXml.Wordprocessing.PageSize>().Single().Width!.Value.ToString());
        Assert.Equal("16838", s0.Elements<DocumentFormat.OpenXml.Wordprocessing.PageSize>().Single().Height!.Value.ToString());
        Assert.Equal(PageOrientationValues.Portrait, s0.Elements<DocumentFormat.OpenXml.Wordprocessing.PageSize>().Single().Orient!.Value);

        // Section 1 → section 2: oddPage break; section 1 is custom 400×600 landscape (swapped).
        var s1 = intermediate[1];
        Assert.Equal(SectionMarkValues.OddPage, s1.Elements<SectionType>().Single().Val!.Value);
        var s1PageSize = s1.Elements<DocumentFormat.OpenXml.Wordprocessing.PageSize>().Single();
        Assert.Equal("12000", s1PageSize.Width!.Value.ToString());
        Assert.Equal("8000", s1PageSize.Height!.Value.ToString());
        Assert.Equal(PageOrientationValues.Landscape, s1PageSize.Orient!.Value);
        var s1Margins = s1.Elements<PageMargin>().Single();
        Assert.Equal("600", s1Margins.Top!.Value.ToString());
        Assert.Equal("800", s1Margins.Left!.Value.ToString());
        var s1Columns = s1.Elements<Columns>().Single();
        Assert.Equal(2, s1Columns.ColumnCount!.Value);
        Assert.Equal("240", s1Columns.Space!.Value!);
        Assert.True(s1Columns.Separator!.Value);

        // Final section (section 2): body-level sectPr is A5 portrait.
        var final = Assert.Single(body.Elements<SectionProperties>());
        var finalPageSize = final.Elements<DocumentFormat.OpenXml.Wordprocessing.PageSize>().Single();
        Assert.Equal("8390", finalPageSize.Width!.Value.ToString());
        Assert.Equal("11906", finalPageSize.Height!.Value.ToString());
        Assert.Equal(PageOrientationValues.Portrait, finalPageSize.Orient!.Value);
    }

    [Fact]
    public void PerSectionHeadersFooters_AreIsolated()
    {
        var json = """
            {
              "version": "1.0",
              "design": { "palette": { "ink": "#1F2937" } },
              "sections": [
                { "blocks": [ { "type": "paragraph", "text": "s0" } ] },
                {
                  "header": [ { "type": "paragraph", "text": "S1 HEADER" } ],
                  "footer": [ { "type": "paragraph", "text": "S1 FOOTER" } ],
                  "blocks": [ { "type": "paragraph", "text": "s1" } ]
                },
                {
                  "header": [ { "type": "paragraph", "text": "S2 HEADER" } ],
                  "blocks": [ { "type": "paragraph", "text": "s2" } ]
                }
              ]
            }
            """;

        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(json);
        var mainPart = DocxTestHarness.OpenMainPart(path);

        Assert.Equal(3, mainPart.HeaderParts.Count());
        Assert.Equal(3, mainPart.FooterParts.Count());

        var headerTexts = mainPart.HeaderParts
            .Select(h => string.Join(" ", h.Header!.Descendants<Text>().Select(t => t.Text)))
            .ToList();
        Assert.Contains(headerTexts, t => t.Contains("S1 HEADER"));
        Assert.Contains(headerTexts, t => t.Contains("S2 HEADER"));

        var footerTexts = mainPart.FooterParts
            .Select(f => string.Join(" ", f.Footer!.Descendants<Text>().Select(t => t.Text)))
            .ToList();
        Assert.Contains(footerTexts, t => t.Contains("S1 FOOTER"));

        // Every sectPr carries the default/even/first reference trio for header and footer.
        var body = mainPart.Document!.Body!;
        var allSectionProperties = body.Elements<SectionProperties>()
            .Concat(body.Elements<Paragraph>()
                .Where(p => p.ParagraphProperties?.SectionProperties is not null)
                .Select(p => p.ParagraphProperties!.SectionProperties!))
            .ToList();
        Assert.Equal(3, allSectionProperties.Count);
        foreach (var sectPr in allSectionProperties)
        {
            Assert.Equal(3, sectPr.Elements<HeaderReference>().Count());
            Assert.Equal(3, sectPr.Elements<FooterReference>().Count());
        }
    }

    [Fact]
    public void Lists_AllocateFreshInstancesAndDistinctAbstracts()
    {
        var json = """
            {
              "version": "1.0",
              "design": { "palette": { "ink": "#1F2937" } },
              "sections": [ { "blocks": [
                { "type": "list", "kind": "ordered", "start": 2, "items": [ "a", "b" ] },
                { "type": "list", "kind": "ordered", "items": [ "c" ] },
                { "type": "list", "kind": "bullet", "items": [ "x" ] },
                { "type": "list", "kind": "bullet", "items": [ "y" ] }
              ] } ]
            }
            """;

        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(json);
        var mainPart = DocxTestHarness.OpenMainPart(path);
        var numbering = mainPart.NumberingDefinitionsPart!.Numbering!;

        var abstracts = numbering.Elements<AbstractNum>().ToList();
        Assert.Equal(3, abstracts.Count);
        Assert.Contains(abstracts, a => a.Elements<Level>().Single().StartNumberingValue!.Val!.Value == 2);

        var instances = numbering.Elements<NumberingInstance>().ToList();
        Assert.Equal(4, instances.Count);
        Assert.Equal(4, instances.Select(n => n.NumberID!.Value).Distinct().Count());

        // Each list got its own instance, so the four lists use four different numIds.
        var numIds = mainPart.Document!.Descendants<NumberingProperties>()
            .Select(n => n.NumberingId!.Val!.Value)
            .ToList();
        Assert.Equal(5, numIds.Count); // 2 + 1 + 1 + 1 items
        Assert.Equal(4, numIds.Distinct().Count());

        // Ordered lists reference the start-2 decimal abstract; bullets the bullet abstract.
        var orderedAbstract = abstracts.Single(a => a.Elements<Level>().Single().StartNumberingValue!.Val!.Value == 2);
        Assert.Equal(NumberFormatValues.Decimal, orderedAbstract.Elements<Level>().Single().NumberingFormat!.Val!.Value);
        var bulletAbstract = abstracts.Single(a => a.Elements<Level>().Single().NumberingFormat!.Val!.Value == NumberFormatValues.Bullet);
        var orderedAbstractId = orderedAbstract.AbstractNumberId!.Value;
        var bulletAbstractId = bulletAbstract.AbstractNumberId!.Value;
        var orderedNum = instances.First(n => n.NumberID!.Value == numIds[0]);
        var bulletNum = instances.First(n => n.NumberID!.Value == numIds[3]);
        Assert.Equal(orderedAbstractId, orderedNum.Elements<AbstractNumId>().Single().Val!.Value);
        Assert.Equal(bulletAbstractId, bulletNum.Elements<AbstractNumId>().Single().Val!.Value);
    }

    [Fact]
    public void Table_EmitsPropertiesGridHeaderFillsAndBlankCell()
    {
        var json = """
            {
              "version": "1.0",
              "design": { "palette": { "ink": "#1F2937", "navy": "#1F4E79" } },
              "sections": [ { "blocks": [
                { "type": "table", "alignment": "center", "widths": [ 216, 108 ], "rows": [
                  { "header": true, "cells": [ { "text": "H1", "fill": "navy", "alignment": "center" }, { "text": "H2" } ] },
                  { "cells": [ { "text": "a" }, {} ] }
                ] }
              ] } ]
            }
            """;

        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(json);
        var mainPart = DocxTestHarness.OpenMainPart(path);
        var table = Assert.Single(mainPart.Document!.Body!.Descendants<Table>());

        var tableProperties = table.TableProperties!;
        Assert.Equal(TableRowAlignmentValues.Center, tableProperties.TableJustification!.Val!.Value);
        Assert.Equal("6480", tableProperties.TableWidth!.Width!.Value!); // (216 + 108) pt → twips
        Assert.Equal(6, tableProperties.TableBorders!.ChildElements.Count);

        var grid = Assert.Single(table.Elements<TableGrid>());
        Assert.Equal(2, grid.Elements<GridColumn>().Count());
        Assert.Equal("4320", grid.Elements<GridColumn>().ElementAt(0).Width!.Value!);
        Assert.Equal("2160", grid.Elements<GridColumn>().ElementAt(1).Width!.Value!);

        var rows = table.Elements<TableRow>().ToList();
        Assert.Equal(2, rows.Count);
        Assert.True(rows[0].TableRowProperties?.Elements<TableHeader>().Any() == true);
        Assert.False(rows[1].TableRowProperties?.Elements<TableHeader>().Any() == true);

        var headerCell = rows[0].Elements<TableCell>().First();
        Assert.Equal("1F4E79", headerCell.TableCellProperties!.Shading!.Fill!.Value);
        Assert.Equal(JustificationValues.Center, headerCell.Descendants<Justification>().Single().Val!.Value);

        var blankCell = rows[1].Elements<TableCell>().ElementAt(1);
        Assert.Single(blankCell.Elements<Paragraph>());
        Assert.Empty(blankCell.Descendants<Text>());
    }

    [Fact]
    public void Callout_EmitsSingleCellShadedTableWithToneAccent()
    {
        var json = """
            {
              "version": "1.0",
              "design": { "palette": { "ink": "#1F2937" } },
              "sections": [ { "blocks": [
                { "type": "callout", "tone": "error", "text": "Something broke" }
              ] } ]
            }
            """;

        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(json);
        var mainPart = DocxTestHarness.OpenMainPart(path);
        var table = Assert.Single(mainPart.Document!.Body!.Descendants<Table>());
        var cell = Assert.Single(table.Elements<TableRow>()).Elements<TableCell>().Single();

        Assert.Equal("FDECEA", cell.TableCellProperties!.Shading!.Fill!.Value);
        var borders = table.TableProperties!.TableBorders!;
        Assert.Equal("24", borders.LeftBorder!.Size!.Value.ToString());
        Assert.Equal("C62828", borders.LeftBorder.Color!.Value);
    }

    [Fact]
    public void PageBreak_EmitsPageBreakParagraph()
    {
        var json = """
            {
              "version": "1.0",
              "design": { "palette": { "ink": "#1F2937" } },
              "sections": [ { "blocks": [
                { "type": "paragraph", "text": "before" },
                { "type": "pageBreak" },
                { "type": "paragraph", "text": "after" }
              ] } ]
            }
            """;

        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(json);
        var mainPart = DocxTestHarness.OpenMainPart(path);

        var breakElement = mainPart.Document!.Descendants<Break>().Single();
        Assert.Equal(BreakValues.Page, breakElement.Type!.Value);
    }

    [Fact]
    public void Group_IsFlattenedIntoParentFlow()
    {
        var json = """
            {
              "version": "1.0",
              "design": { "palette": { "ink": "#1F2937" } },
              "sections": [ { "blocks": [
                { "type": "paragraph", "text": "before" },
                { "type": "group", "blocks": [
                  { "type": "heading", "level": 3, "text": "Inside group" },
                  { "type": "paragraph", "text": "inside" }
                ] },
                { "type": "paragraph", "text": "after" }
              ] } ]
            }
            """;

        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(json);
        var mainPart = DocxTestHarness.OpenMainPart(path);

        var textOrder = mainPart.Document!.Body!.Descendants<Text>().Select(t => t.Text).ToList();
        Assert.Equal(new[] { "before", "Inside group", "inside", "after" }, textOrder);
    }

    [Fact]
    public void Headings_EmitOutlineLevelAndBaselineStyleIds()
    {
        var json = """
            {
              "version": "1.0",
              "design": { "palette": { "ink": "#1F2937" } },
              "sections": [ { "blocks": [
                { "type": "heading", "level": 1, "text": "One" },
                { "type": "heading", "level": 2, "text": "Two" },
                { "type": "heading", "level": 3, "text": "Three" }
              ] } ]
            }
            """;

        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(json);
        var mainPart = DocxTestHarness.OpenMainPart(path);

        var outlineLevels = mainPart.Document!.Descendants<OutlineLevel>().Select(o => o.Val!.Value).ToList();
        Assert.Equal(new[] { 0, 1, 2 }, outlineLevels);

        var styleIds = mainPart.Document.Descendants<ParagraphStyleId>().Select(p => p.Val!.Value).ToList();
        Assert.Contains("Heading1", styleIds);
        Assert.Contains("Heading2", styleIds);
        Assert.Contains("Heading3", styleIds);

        var styles = mainPart.StyleDefinitionsPart!.Styles!;
        Assert.NotNull(styles.Elements<Style>().FirstOrDefault(s => s.StyleId == "Heading1"));
        Assert.NotNull(styles.Elements<Style>().FirstOrDefault(s => s.StyleId == "Heading2"));
        Assert.NotNull(styles.Elements<Style>().FirstOrDefault(s => s.StyleId == "Heading3"));
        Assert.NotNull(styles.Elements<Style>().FirstOrDefault(s => s.StyleId == "Normal"));
    }

    [Fact]
    public void Runs_ApplyDirectCharacterFormatting()
    {
        var json = """
            {
              "version": "1.0",
              "design": {
                "palette": { "ink": "#1F2937" },
                "fonts": { "body": "Aptos", "display": "Aptos Display" },
                "typography": { "body": { "font": "body", "size": 10.5 } }
              },
              "sections": [ { "blocks": [
                { "type": "paragraph", "token": "body", "runs": [
                  { "text": "B", "bold": true },
                  { "text": "I", "italic": true },
                  { "text": "U", "underline": true },
                  { "text": "C", "color": "ink" },
                  { "text": "S", "size": 18, "font": "display" }
                ] }
              ] } ]
            }
            """;

        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(json);
        var mainPart = DocxTestHarness.OpenMainPart(path);

        var runs = mainPart.Document!.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>().ToList();
        Assert.Equal(5, runs.Count);

        Assert.NotNull(runs[0].RunProperties!.Bold);
        Assert.NotNull(runs[1].RunProperties!.Italic);
        Assert.Equal(UnderlineValues.Single, runs[2].RunProperties!.Underline!.Val!.Value);
        Assert.Equal("1F2937", runs[3].RunProperties!.Color!.Val!.Value);
        Assert.Equal("36", runs[4].RunProperties!.FontSize!.Val!.Value);      // 18pt → 36 half-points
        Assert.Equal("Aptos Display", runs[4].RunProperties!.RunFonts!.Ascii!.Value);

        // Token-driven defaults: body token supplies size 10.5 (21 half-points).
        Assert.Equal("21", runs[0].RunProperties!.FontSize!.Val!.Value);
    }

    [Fact]
    public void UnknownStyleReference_WarnsAndFallsBackToBody()
    {
        var json = """
            {
              "version": "1.0",
              "design": { "palette": { "ink": "#1F2937" } },
              "sections": [ { "blocks": [
                { "type": "paragraph", "style": "NoSuchStyle", "text": "x" }
              ] } ]
            }
            """;

        var result = DocxTestHarness.GenerateToBytes(json);

        Assert.Contains(result.Result.Warnings, w =>
            w.Message.Contains("UnknownStyleReference") && w.Message.Contains("NoSuchStyle"));
        Assert.Contains(result.Result.Warnings, w => w.Path == "$.sections[0].blocks[0]");
    }

    [Fact]
    public void TemplateAppend_PreservesTemplateContentAndStyles()
    {
        using var temp = new TempDirectory();
        var templatePath = BuildTemplate(temp.Path);

        var templateBytes = File.ReadAllBytes(templatePath);
        var templateStyleXml = ReadTemplateStyleXml(templatePath);

        var json = $$"""
            {
              "version": "1.0",
              "template": "{{templatePath.Replace("\\", "\\\\")}}",
              "sections": [ { "blocks": [
                { "type": "paragraph", "style": "MyCustom", "text": "GENERATED CONTENT" },
                { "type": "list", "kind": "bullet", "items": [ "a", "b" ] }
              ] } ]
            }
            """;

        using var temp2 = new TempDirectory();
        var output = temp2.File("out.docx");
        new DocxGenerator().Generate(json, output);

        // The template file on disk is never mutated.
        Assert.Equal(templateBytes, File.ReadAllBytes(templatePath));

        var mainPart = DocxTestHarness.OpenMainPart(output);
        var body = mainPart.Document!.Body!;

        // Template body content survives before the generated content.
        var textOrder = body.Descendants<Text>().Select(t => t.Text).ToList();
        Assert.Equal(new[] { "TEMPLATE BODY", "GENERATED CONTENT", "a", "b" }, textOrder);

        // The template's paragraph style is referenced, not redefined.
        var templateParagraph = body.Elements<Paragraph>().Single(p => p.InnerText.Contains("TEMPLATE BODY"));
        var generatedParagraph = body.Elements<Paragraph>().Single(p => p.InnerText.Contains("GENERATED CONTENT"));
        Assert.Equal("MyCustom", templateParagraph.ParagraphProperties!.ParagraphStyleId!.Val!.Value);
        Assert.Equal("MyCustom", generatedParagraph.ParagraphProperties!.ParagraphStyleId!.Val!.Value);

        // The template's own body-level sectPr was converted into an intermediate break.
        Assert.Single(body.Elements<Paragraph>(), p => p.ParagraphProperties?.SectionProperties is not null);
        Assert.Single(body.Elements<SectionProperties>());

        // Existing style definition is byte-identical (never mutated) and still present.
        var generatedStyleXml = mainPart.StyleDefinitionsPart!.Styles!.Elements<Style>()
            .Single(s => s.StyleId == "MyCustom")
            .OuterXml;
        Assert.Equal(templateStyleXml, generatedStyleXml);

        // Template numbering (num 1 / abstract 0) is preserved and new definitions are collision-free.
        var numbering = mainPart.NumberingDefinitionsPart!.Numbering!;
        Assert.Contains(numbering.Elements<NumberingInstance>(), n => n.NumberID!.Value == 1);
        Assert.Contains(numbering.Elements<AbstractNum>(), a => a.AbstractNumberId!.Value == 0);
        var newInstances = numbering.Elements<NumberingInstance>().Select(n => n.NumberID!.Value).ToList();
        Assert.True(newInstances.Max() >= 2);
        var newAbstracts = numbering.Elements<AbstractNum>().Select(a => a.AbstractNumberId!.Value).ToList();
        Assert.True(newAbstracts.Max() >= 1);

        // The generated section geometry is present (default A4 portrait).
        var final = Assert.Single(body.Elements<SectionProperties>());
        Assert.Equal("11906", final.Elements<DocumentFormat.OpenXml.Wordprocessing.PageSize>().Single().Width!.Value.ToString());
    }

    [Fact]
    public void MissingTemplateFile_FailsWithDomainError()
    {
        using var temp = new TempDirectory();
        var missing = temp.File("missing-template.docx");
        var json = $$"""
            {
              "version": "1.0",
              "template": "{{missing}}",
              "sections": [ { "blocks": [ { "type": "paragraph", "text": "x" } ] } ]
            }
            """;

        var exception = Assert.Throws<OfficeEditorException>(() =>
            DocxTestHarness.GenerateToTempFile(json));

        Assert.Contains("template file", exception.Message);
    }

    [Fact]
    public void MissingTemplateDirectory_ReportsDomainErrorNotRawIOException()
    {
        // Regression guard: a template whose parent directory does not exist currently leaks a
        // raw DirectoryNotFoundException instead of the domain error (File.ReadAllBytes throws
        // DirectoryNotFoundException, which the emitter's catch only maps for FileNotFoundException).
        var json = """
            {
              "version": "1.0",
              "template": "/definitely/not/here/template.docx",
              "sections": [ { "blocks": [ { "type": "paragraph", "text": "x" } ] } ]
            }
            """;

        var exception = Assert.Throws<OfficeEditorException>(() =>
            DocxTestHarness.GenerateToTempFile(json));

        Assert.Contains("template file", exception.Message);
    }

    private static string ReadTemplateStyleXml(string templatePath)
    {
        using var document = WordprocessingDocument.Open(templatePath, false);
        return document.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Elements<Style>()
            .Single(s => s.StyleId == "MyCustom")
            .OuterXml;
    }

    /// <summary>Builds a small template DOCX with a custom paragraph style, a body paragraph
    /// using it, a numbering part and a body-level sectPr.</summary>
    private static string BuildTemplate(string directory)
    {
        var path = Path.Combine(directory, "template.docx");
        using (var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();

            var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
            var customStyle = new Style(
                new StyleName { Val = "My Custom" },
                new BasedOn { Val = "Normal" },
                new StyleParagraphProperties())
            {
                Type = StyleValues.Paragraph,
                StyleId = "MyCustom"
            };
            stylesPart.Styles = new Styles(customStyle);

            var numberingPart = main.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = new Numbering(
                new AbstractNum(
                    new Level(
                        new NumberingFormat { Val = NumberFormatValues.Decimal },
                        new LevelText { Val = "%1." })
                    { LevelIndex = 0 })
                { AbstractNumberId = 0 },
                new NumberingInstance(new AbstractNumId { Val = 0 }) { NumberID = 1 });

            var body = new Body(
                new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId { Val = "MyCustom" }),
                    new DocumentFormat.OpenXml.Wordprocessing.Run(new Text("TEMPLATE BODY"))),
                new SectionProperties(
                    new PageSize { Width = 12240, Height = 15840 },
                    new PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440 }));

            main.Document = new Document(body);
        }
        return path;
    }
}
