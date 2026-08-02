using DocxEditor.Core.Generation.Expand;
using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Tests.Generation.Docx;

/// <summary>
/// The pure expansion stage: lowers semantic report archetypes into concrete flow blocks,
/// recurses into nested flow, keeps archetype-free documents untouched, and surfaces budget
/// warnings with JSON paths. After expansion no <see cref="ReportArchetype"/> remains.
/// </summary>
public class DocxGenerationArchetypeExpanderTests
{
    private static DocxExpansionResult Expand(string json) =>
        DocxGenerationExpander.ExpandWithIssues(DocxTestHarness.Parse(json));

    private static string Section(string blocksJson) =>
        $$"""
        {
          "version": "1.0",
          "sections": [ { "blocks": [ {{blocksJson}} ] } ]
        }
        """;

    [Fact]
    public void DocumentWithoutArchetypes_PassesThroughIdentically()
    {
        var document = DocxTestHarness.Parse(DocxTestHarness.MinimalJson);

        var expansion = DocxGenerationExpander.ExpandWithIssues(document);

        Assert.Same(document, expansion.Document);
        Assert.Empty(expansion.Warnings);
    }

    [Fact]
    public void Cover_ExpandsToEditorialTitleBlockAndOptionalParts()
    {
        var result = Expand(Section("""
            {
              "type": "cover",
              "eyebrow": "Q3 2026",
              "title": "State of the Product",
              "subtitle": "A quarterly review.",
              "metadata": "Prepared by the team",
              "kpis": [
                { "value": "12.4k", "label": "Active workspaces", "tone": "positive" },
                { "value": "98%", "label": "On-time delivery" }
              ],
              "pageBreak": true
            }
            """));

        var blocks = result.Document.Sections[0].Blocks;
        Assert.Empty(result.Warnings);
        Assert.Equal(6, blocks.Count);

        var eyebrow = Assert.IsType<ParagraphBlock>(blocks[0]);
        Assert.Equal(TextRole.Eyebrow, eyebrow.Content.Role);
        var title = Assert.IsType<ParagraphBlock>(blocks[1]);
        Assert.Equal(TextRole.Title, title.Content.Role);
        var subtitle = Assert.IsType<ParagraphBlock>(blocks[2]);
        Assert.Equal(TextRole.Subtitle, subtitle.Content.Role);
        var metadata = Assert.IsType<ParagraphBlock>(blocks[3]);
        Assert.Equal(TextRole.Muted, metadata.Content.Role);

        var kpiTable = Assert.IsType<TableBlock>(blocks[4]);
        Assert.Equal(2, kpiTable.Rows.Count);
        Assert.Equal(2, kpiTable.Rows[0].Cells.Count);
        Assert.IsType<PageBreakBlock>(blocks[5]);
    }

    [Fact]
    public void Cover_WithoutKpisOrPageBreak_OmitsTableAndBreak()
    {
        var result = Expand(Section("""
            { "type": "cover", "title": "Just a title" }
            """));

        var blocks = result.Document.Sections[0].Blocks;
        var title = Assert.IsType<ParagraphBlock>(Assert.Single(blocks));
        Assert.Equal(TextRole.Title, title.Content.Role);
        Assert.Empty(result.Warnings);
    }

    [Fact]
    public void Cover_ToneTintsKpiValuesAgainstThemeTokens()
    {
        var result = Expand(Section("""
            {
              "type": "cover",
              "title": "T",
              "kpis": [
                { "value": "98%", "label": "On time", "tone": "positive" },
                { "value": "3", "label": "Sev-1", "tone": "negative" },
                { "value": "27", "label": "Shipments" }
              ]
            }
            """));

        var table = Assert.IsType<TableBlock>(result.Document.Sections[0].Blocks[1]);
        var valueCell = table.Rows[0].Cells[0];
        var run = valueCell.Content!.Runs!.Single();
        Assert.Equal("teal", run.Color);
        Assert.Equal(TextRole.Metric, valueCell.Content.Role);
        Assert.Equal("coral", table.Rows[0].Cells[1].Content!.Runs!.Single().Color);
        // Neutral tone leaves the plain value untouched (no tint, no runs).
        Assert.Equal("27", table.Rows[0].Cells[2].Content!.Text);
        Assert.Null(table.Rows[0].Cells[2].Content!.Runs);
        Assert.Equal("pale", valueCell.Fill);
    }

    [Fact]
    public void KpiRow_ExpandsToSingleBandTable()
    {
        var result = Expand(Section("""
            { "type": "kpiRow", "items": [
              { "value": "31%", "label": "Growth" },
              { "value": "27", "label": "Shipments" },
              { "value": "3", "label": "Open issues", "tone": "negative" }
            ] }
            """));

        var table = Assert.IsType<TableBlock>(Assert.Single(result.Document.Sections[0].Blocks));
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(3, table.Rows[0].Cells.Count);
        Assert.Equal(TextRole.Metric, table.Rows[0].Cells[0].Content!.Role);
        Assert.Equal(TextRole.MetricLabel, table.Rows[1].Cells[0].Content!.Role);
        Assert.Empty(result.Warnings);
    }

    [Fact]
    public void SemanticSection_ExpandsHeadingIntroAndRecursivelyLowersNestedArchetypes()
    {
        var result = Expand(Section("""
            {
              "type": "section",
              "title": "Executive summary",
              "intro": "Adoption keeps compounding.",
              "blocks": [
                { "type": "paragraph", "text": "Body copy." },
                { "type": "kpiRow", "items": [
                  { "value": "1", "label": "One" },
                  { "value": "2", "label": "Two" }
                ] }
              ]
            }
            """));

        var blocks = result.Document.Sections[0].Blocks;
        Assert.Equal(4, blocks.Count);

        var heading = Assert.IsType<HeadingBlock>(blocks[0]);
        Assert.Equal(1, heading.Level);
        Assert.Equal(TextRole.Heading1, heading.Content.Role);

        var intro = Assert.IsType<ParagraphBlock>(blocks[1]);
        Assert.Equal(TextRole.Body, intro.Content.Role);

        var paragraph = Assert.IsType<ParagraphBlock>(blocks[2]);
        Assert.Equal("Body copy.", paragraph.Content.Text);

        Assert.IsType<TableBlock>(blocks[3]);
        Assert.Empty(result.Warnings);
    }

    [Fact]
    public void ComparisonTable_ExpandsToHeaderAndBodyRows()
    {
        var result = Expand(Section("""
            {
              "type": "comparisonTable",
              "emphasisFirstColumn": true,
              "columns": ["Criterion", "OfficeEditor", "Legacy"],
              "rows": [
                { "cells": ["Time to first deck", "Under a minute", "About a day"] },
                { "cells": ["Fidelity", "95% parity", "Manual"] }
              ]
            }
            """));

        var table = Assert.IsType<TableBlock>(Assert.Single(result.Document.Sections[0].Blocks));
        Assert.Equal(3, table.Rows.Count);

        Assert.True(table.Rows[0].IsHeader);
        Assert.Equal(TextRole.TableHeader, table.Rows[0].Cells[0].Content!.Role);

        var firstBodyRow = table.Rows[1];
        Assert.False(firstBodyRow.IsHeader);
        // Emphasized first column uses the label role; the rest are table body.
        Assert.Equal(TextRole.Label, firstBodyRow.Cells[0].Content!.Role);
        Assert.Equal(TextRole.TableBody, firstBodyRow.Cells[1].Content!.Role);
        Assert.Empty(result.Warnings);
    }

    [Fact]
    public void Roadmap_ExpandsToPhaseWindowActionEvidenceTable()
    {
        var result = Expand(Section("""
            {
              "type": "roadmap",
              "phases": [
                { "window": "Aug", "action": "Ship hardening.", "evidence": "Zero Sev-1", "tone": "positive" },
                { "window": "Sep", "action": "Open preview." }
              ]
            }
            """));

        var table = Assert.IsType<TableBlock>(Assert.Single(result.Document.Sections[0].Blocks));
        Assert.Equal(3, table.Rows.Count);

        var header = table.Rows[0];
        Assert.True(header.IsHeader);
        Assert.Equal("Phase", header.Cells[0].Content!.Text);
        Assert.Equal("Window", header.Cells[1].Content!.Text);
        Assert.Equal("Action", header.Cells[2].Content!.Text);
        Assert.Equal("Evidence", header.Cells[3].Content!.Text);

        var first = table.Rows[1];
        Assert.Equal("1", first.Cells[0].Content!.Runs!.Single().Text);
        Assert.Equal("teal", first.Cells[0].Content!.Runs!.Single().Color);
        Assert.Equal("Aug", first.Cells[1].Content!.Text);
        Assert.Equal("Ship hardening.", first.Cells[2].Content!.Text);
        Assert.Equal("Zero Sev-1", first.Cells[3].Content!.Text);

        // Missing evidence yields a blank cell.
        var second = table.Rows[2];
        Assert.Null(second.Cells[3].Content);
        Assert.Empty(result.Warnings);
    }

    [Fact]
    public void Expansion_LeavesNoArchetypeBlocksAnywhere()
    {
        var json = DocxTestHarness.ReadEditorialReport();

        var result = Expand(json);

        AssertArchetypeFree(result.Document.Sections[0].Blocks);
        Assert.Empty(result.Warnings);
    }

    [Fact]
    public void KpiRow_OverBudget_WarnsWithPath()
    {
        var result = Expand(Section("""
            { "type": "kpiRow", "items": [
              { "value": "1", "label": "One" }
            ] }
            """));

        var warning = Assert.Single(result.Warnings);
        Assert.Equal("$.sections[0].blocks[0].items", warning.Path);
        Assert.Contains("recommended budget is 2–4", warning.Message);
        // The band still renders (advisory, never a hard failure).
        Assert.IsType<TableBlock>(Assert.Single(result.Document.Sections[0].Blocks));
    }

    [Fact]
    public void Cover_WithTooManyKpis_Warns()
    {
        var result = Expand(Section("""
            {
              "type": "cover",
              "title": "T",
              "kpis": [
                { "value": "1", "label": "A" },
                { "value": "2", "label": "B" },
                { "value": "3", "label": "C" },
                { "value": "4", "label": "D" },
                { "value": "5", "label": "E" }
              ]
            }
            """));

        var warning = Assert.Single(result.Warnings);
        Assert.Equal("$.sections[0].blocks[0].kpis", warning.Path);
        Assert.Contains("has 5 kpi items", warning.Message);
        Assert.Equal("pale", Assert.IsType<TableBlock>(result.Document.Sections[0].Blocks[1]).Rows[0].Cells[0].Fill);
    }

    [Fact]
    public void ComparisonTable_ColumnOverBudget_Warns()
    {
        var result = Expand(Section("""
            {
              "type": "comparisonTable",
              "columns": ["a"],
              "rows": [ { "cells": ["x"] } ]
            }
            """));

        var warning = Assert.Single(result.Warnings);
        Assert.Equal("$.sections[0].blocks[0].columns", warning.Path);
        Assert.Contains("recommended budget is 2–6", warning.Message);
    }

    [Fact]
    public void Roadmap_OverBudget_Warns()
    {
        var phases = string.Join(", ", Enumerable.Range(1, 8).Select(i =>
            $"{{\"window\": \"W{i}\", \"action\": \"Act {i}\"}}"));
        var result = Expand(Section($$"""
            { "type": "roadmap", "phases": [ {{phases}} ] }
            """));

        var warning = Assert.Single(result.Warnings);
        Assert.Equal("$.sections[0].blocks[0].phases", warning.Path);
        Assert.Contains("has 8 phases", warning.Message);
        Assert.IsType<TableBlock>(Assert.Single(result.Document.Sections[0].Blocks));
    }

    [Fact]
    public void HeaderAndFooterArchetypes_AreAlsoExpanded()
    {
        var document = DocxTestHarness.Parse("""
            {
              "version": "1.0",
              "sections": [ {
                "header": [ { "type": "kpiRow", "items": [
                  { "value": "1", "label": "A" },
                  { "value": "2", "label": "B" }
                ] } ],
                "blocks": [ { "type": "paragraph", "text": "hi" } ]
              } ]
            }
            """);

        var result = DocxGenerationExpander.ExpandWithIssues(document);

        var header = result.Document.Sections[0].Header;
        Assert.IsType<TableBlock>(Assert.Single(header));
        Assert.Empty(result.Warnings);
    }

    private static void AssertArchetypeFree(IReadOnlyList<FlowBlock> blocks)
    {
        foreach (var block in blocks)
        {
            Assert.IsNotType<CoverBlock>(block);
            Assert.IsNotType<KpiRowBlock>(block);
            Assert.IsNotType<SemanticSectionBlock>(block);
            Assert.IsNotType<ComparisonTableBlock>(block);
            Assert.IsNotType<RoadmapBlock>(block);
            if (block is FlowContainerBlock group)
            {
                AssertArchetypeFree(group.Blocks);
            }
        }
    }
}
