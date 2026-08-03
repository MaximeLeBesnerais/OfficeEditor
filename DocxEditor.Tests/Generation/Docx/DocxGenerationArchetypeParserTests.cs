using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Tests.Generation.Docx;

/// <summary>
/// Strict parser diagnostics for the semantic report archetypes: unknown properties, required
/// fields, value kinds, enum validation, typo hints and JSON paths. Counts outside a component's
/// budget are advisory (the expander warns), so structural validity is what the parser enforces.
/// </summary>
public class DocxGenerationArchetypeParserTests
{
    [Fact]
    public void Cover_ParsesWithAllOptionalFields()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "sections": [ { "blocks": [
                {
                  "type": "cover",
                  "eyebrow": "Q3 2026",
                  "title": "State of the Product",
                  "subtitle": "A quarterly review.",
                  "metadata": "Prepared by Platform Group",
                  "kpis": [
                    { "value": "12.4k", "label": "Active workspaces", "tone": "positive" },
                    { "value": "98%", "label": "On-time delivery" }
                  ],
                  "pageBreak": true
                }
              ] } ]
            }
            """);

        Assert.True(result.IsValid);
        Assert.Empty(result.Errors);
        Assert.Empty(result.Warnings);
        var cover = Assert.IsType<CoverBlock>(result.Document!.Sections[0].Blocks[0]);
        Assert.Equal("Q3 2026", cover.Eyebrow!.Text);
        Assert.Equal("State of the Product", cover.Title.Text);
        Assert.Equal("A quarterly review.", cover.Subtitle!.Text);
        Assert.Equal("Prepared by Platform Group", cover.Metadata!.Text);
        Assert.True(cover.PageBreak);
        Assert.Equal(2, cover.Kpis!.Count);
        Assert.Equal(ReportTone.Positive, cover.Kpis[0].Tone);
        Assert.Equal(ReportTone.Neutral, cover.Kpis[1].Tone);
    }

    [Fact]
    public void KpiRow_ParsesItemsAndDefaultTone()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "sections": [ { "blocks": [
                { "type": "kpiRow", "items": [
                  { "value": "31%", "label": "Growth", "tone": "positive" },
                  { "value": "27", "label": "Shipments" },
                  { "value": "3", "label": "Open issues", "tone": "negative" }
                ] }
              ] } ]
            }
            """);

        Assert.True(result.IsValid);
        var row = Assert.IsType<KpiRowBlock>(result.Document!.Sections[0].Blocks[0]);
        Assert.Equal(3, row.Items.Count);
        Assert.Equal(ReportTone.Neutral, row.Items[1].Tone);
        Assert.Equal(ReportTone.Negative, row.Items[2].Tone);
    }

    [Fact]
    public void SemanticSection_ParsesTitleIntroAndNestedBlocks()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "sections": [ { "blocks": [
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
              ] } ]
            }
            """);

        Assert.True(result.IsValid);
        var section = Assert.IsType<SemanticSectionBlock>(result.Document!.Sections[0].Blocks[0]);
        Assert.Equal("Executive summary", section.Title.Text);
        Assert.Equal("Adoption keeps compounding.", section.Intro!.Text);
        Assert.Equal(2, section.Blocks.Count);
        Assert.IsType<ParagraphBlock>(section.Blocks[0]);
        Assert.IsType<KpiRowBlock>(section.Blocks[1]);
    }

    [Fact]
    public void ComparisonTable_ParsesColumnsRowsAndEmphasis()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "sections": [ { "blocks": [
                {
                  "type": "comparisonTable",
                  "emphasisFirstColumn": true,
                  "columns": ["Criterion", "OfficeEditor", "Legacy"],
                  "rows": [
                    { "cells": ["Time to first deck", "Under a minute", "About a day"] },
                    { "cells": [{ "text": "Fidelity" }, "95% parity", "Manual"] }
                  ]
                }
              ] } ]
            }
            """);

        Assert.True(result.IsValid);
        var table = Assert.IsType<ComparisonTableBlock>(result.Document!.Sections[0].Blocks[0]);
        Assert.True(table.EmphasisFirstColumn);
        Assert.Equal(3, table.Columns.Count);
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(3, table.Rows[1].Cells.Count);
        Assert.Equal("Fidelity", table.Rows[1].Cells[0].Text);
    }

    [Fact]
    public void Roadmap_ParsesPhasesWithTone()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "sections": [ { "blocks": [
                { "type": "roadmap", "phases": [
                  { "window": "Aug", "action": "Ship hardening.", "evidence": "Zero Sev-1", "tone": "positive" },
                  { "window": "Sep", "action": "Open preview." }
                ] }
              ] } ]
            }
            """);

        Assert.True(result.IsValid);
        var roadmap = Assert.IsType<RoadmapBlock>(result.Document!.Sections[0].Blocks[0]);
        Assert.Equal(2, roadmap.Phases.Count);
        Assert.Equal(ReportTone.Positive, roadmap.Phases[0].Tone);
        Assert.Equal("Zero Sev-1", roadmap.Phases[0].Evidence!.Text);
        Assert.Null(roadmap.Phases[1].Evidence);
    }

    [Fact]
    public void Cover_WithoutTitle_IsRejected()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "sections": [ { "blocks": [
                { "type": "cover", "subtitle": "missing title" }
              ] } ]
            }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.sections[0].blocks[0]", issue.Path);
        Assert.Contains("'title' is required", issue.Message);
    }

    [Fact]
    public void KpiItem_MissingValue_IsRejected()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "sections": [ { "blocks": [
                { "type": "kpiRow", "items": [ { "label": "No value" } ] }
              ] } ]
            }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.sections[0].blocks[0].items[0]", issue.Path);
        Assert.Contains("'value' is required", issue.Message);
    }

    [Fact]
    public void RoadmapPhase_MissingAction_IsRejected()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "sections": [ { "blocks": [
                { "type": "roadmap", "phases": [ { "window": "Aug" } ] }
              ] } ]
            }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.sections[0].blocks[0].phases[0]", issue.Path);
        Assert.Contains("'action' is required", issue.Message);
    }

    [Fact]
    public void InvalidTone_IsRejectedWithValidValues()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "sections": [ { "blocks": [
                { "type": "kpiRow", "items": [
                  { "value": "1", "label": "One", "tone": "postive" }
                ] }
              ] } ]
            }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.sections[0].blocks[0].items[0].tone", issue.Path);
        Assert.Contains("'postive' is not a valid report tone", issue.Message);
        Assert.Contains("positive, neutral, negative", issue.Message);
        Assert.Contains("Did you mean 'positive'?", issue.Suggestion);
    }

    [Fact]
    public void UnknownCoverProperty_IsRejectedWithPath()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "sections": [ { "blocks": [
                { "type": "cover", "title": "T", "kpi": [] }
              ] } ]
            }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.sections[0].blocks[0].kpi", issue.Path);
        Assert.Contains("unknown property 'kpi'", issue.Message);
        Assert.Contains("Did you mean 'kpis'?", issue.Suggestion);
    }

    [Fact]
    public void ComparisonTableTypo_ProducesHint()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "sections": [ { "blocks": [
                { "type": "comparsionTable", "columns": ["a", "b"], "rows": [ { "cells": ["x", "y"] } ] }
              ] } ]
            }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.sections[0].blocks[0]", issue.Path);
        Assert.Contains("unknown flow block type 'comparsionTable'", issue.Message);
        Assert.Contains("Did you mean 'comparisonTable'?", issue.Suggestion);
    }

    [Fact]
    public void UnknownArchetypeType_ReportsValidValues()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "sections": [ { "blocks": [
                { "type": "coverr", "title": "T" }
              ] } ]
            }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Contains("unknown flow block type 'coverr'", issue.Message);
        Assert.Contains("cover|kpiRow|section|comparisonTable|roadmap", issue.Message);
        Assert.Contains("Did you mean 'cover'?", issue.Suggestion);
    }

    [Fact]
    public void RaggedComparisonTableRows_AreRejected()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "sections": [ { "blocks": [
                {
                  "type": "comparisonTable",
                  "columns": ["a", "b"],
                  "rows": [
                    { "cells": ["x", "y"] },
                    { "cells": ["only one"] }
                  ]
                }
              ] } ]
            }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.sections[0].blocks[0].rows[1]", issue.Path);
        Assert.Contains("row has 1 cells but the comparison table has 2 columns", issue.Message);
    }

    [Fact]
    public void EmptyKpiRowItems_AreRejected()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "sections": [ { "blocks": [
                { "type": "kpiRow", "items": [] }
              ] } ]
            }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.sections[0].blocks[0].items", issue.Path);
        Assert.Contains("must contain at least one KPI item", issue.Message);
    }

    [Fact]
    public void SemanticSection_WithoutBlocks_IsRejected()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "sections": [ { "blocks": [
                { "type": "section", "title": "Only a title" }
              ] } ]
            }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.sections[0].blocks[0]", issue.Path);
        Assert.Contains("'blocks' is required", issue.Message);
    }

    [Fact]
    public void EmptyRoadmapPhases_AreRejected()
    {
        var result = DocxTestHarness.Validate("""
            {
              "version": "1.0",
              "sections": [ { "blocks": [
                { "type": "roadmap", "phases": [] }
              ] } ]
            }
            """);

        var issue = Assert.Single(result.Errors);
        Assert.Equal("$.sections[0].blocks[0].phases", issue.Path);
        Assert.Contains("must contain at least one phase", issue.Message);
    }
}
