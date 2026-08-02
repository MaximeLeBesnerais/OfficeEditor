using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Generation.Expand;
using DocxEditor.Core.Generation.Model;
using WordCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using WordRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using WordTable = DocumentFormat.OpenXml.Wordprocessing.Table;

namespace DocxEditor.Tests.Generation.Docx;

/// <summary>
/// End-to-end emission of the report archetypes: the report fixture lowers to Word-native
/// flow content (editorial title blocks, pale KPI bands, comparison and roadmap tables, page
/// breaks) with the theme's restrained borders. The generated package reopens cleanly and its
/// warnings surface from the expansion stage through <see cref="Core.Generation.DocxGenerator"/>.
/// </summary>
public class DocxGenerationArchetypeEmitTests
{
    [Fact]
    public void Report_FixtureParsesWithNoErrorsOrWarnings()
    {
        var result = DocxTestHarness.Validate(DocxTestHarness.ReadEditorialReport());

        Assert.True(result.IsValid);
        Assert.Empty(result.Errors);
        Assert.Empty(result.Warnings);

        var document = result.Document!;
        Assert.Equal("1.0", document.Version);
        Assert.Equal("Northwind Labs — Quarterly Editorial Report", document.Metadata!.Title);
        Assert.IsType<CoverBlock>(document.Sections[0].Blocks[0]);
        Assert.Contains(document.Sections[0].Blocks, b => b is SemanticSectionBlock);
    }

    [Fact]
    public void Report_GeneratesToBytesWithoutWarnings()
    {
        var bytes = DocxTestHarness.GenerateToBytes(DocxTestHarness.ReadEditorialReport());

        Assert.NotNull(bytes.Content);
        Assert.True(bytes.Content.Length > 4_000, $"expected a substantial package, got {bytes.Content.Length} bytes");
        Assert.Equal((byte)'P', bytes.Content[0]);
        Assert.Equal((byte)'K', bytes.Content[1]);
        Assert.Empty(bytes.Result.Outputs);
        Assert.Empty(bytes.Result.Warnings);
    }

    [Fact]
    public void Report_ReopensWithExpandedContent()
    {
        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(DocxTestHarness.ReadEditorialReport());
        var mainPart = DocxTestHarness.OpenMainPart(path);
        var allText = string.Join(" ", mainPart.Document!.Descendants<Text>().Select(t => t.Text));

        // Cover surfaces.
        Assert.Contains("Q3 2026 · Editorial Edition", allText);
        Assert.Contains("State of the Product", allText);
        Assert.Contains("Prepared by the Platform Group", allText);

        // Semantic sections became level-1 headings.
        Assert.Contains("Executive summary", allText);
        Assert.Contains("Adoption comparison", allText);
        Assert.Contains("Delivery roadmap", allText);

        // Page break from the cover is a real w:br of type page.
        Assert.Contains(mainPart.Document.Descendants<Break>(), b => b.Type?.Value == BreakValues.Page);

        // Five tables in body order: cover KPIs, summary KPI row, decision callout,
        // comparison, roadmap.
        var tables = mainPart.Document.Body!.Descendants<WordTable>().ToList();
        Assert.Equal(5, tables.Count);
    }

    [Fact]
    public void Report_KpiBandUsesPaleSurfaceAndMetricRoles()
    {
        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(DocxTestHarness.ReadEditorialReport());
        var mainPart = DocxTestHarness.OpenMainPart(path);

        var tables = mainPart.Document!.Body!.Descendants<WordTable>().ToList();
        var kpiBand = tables[0];
        Assert.Equal(2, kpiBand.Elements<WordRow>().Count());
        Assert.Equal(3, kpiBand.Elements<WordRow>().First().Elements<WordCell>().Count());

        var valueCell = kpiBand.Descendants<WordCell>().First();
        var shading = valueCell.TableCellProperties?.Shading;
        Assert.NotNull(shading);
        Assert.Equal("EEF2F6", shading!.Fill?.Value); // editorial 'pale'

        // The positive KPI tints teal (#1F7A6E), the negative one coral (#E4674A).
        var colors = kpiBand.Descendants<Color>().Select(c => c.Val!.Value).ToList();
        Assert.Contains("1F7A6E", colors);
        Assert.Contains("E4674A", colors);
    }

    [Fact]
    public void Report_ComparisonTableHasRepeatingHeaderRow()
    {
        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(DocxTestHarness.ReadEditorialReport());
        var mainPart = DocxTestHarness.OpenMainPart(path);

        var comparison = mainPart.Document!.Body!.Descendants<WordTable>().ElementAt(3);
        Assert.Equal(4, comparison.Elements<WordRow>().Count()); // header + 3 rows
        var headerRow = comparison.Elements<WordRow>().First();
        Assert.True(headerRow.TableRowProperties?.Elements<TableHeader>().Any() == true);

        var headerCells = headerRow.Elements<WordCell>().Select(c => c.InnerText).ToList();
        Assert.Equal(new[] { "Criterion", "OfficeEditor", "Legacy suite" }, headerCells);
    }

    [Fact]
    public void Report_RoadmapTableListsPhasesInOrder()
    {
        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(DocxTestHarness.ReadEditorialReport());
        var mainPart = DocxTestHarness.OpenMainPart(path);

        var roadmap = mainPart.Document!.Body!.Descendants<WordTable>().ElementAt(4);
        Assert.Equal(5, roadmap.Elements<WordRow>().Count()); // header + 4 phases

        var phaseNumbers = roadmap.Descendants<WordCell>()
            .Select(c => c.InnerText)
            .Where(t => t is "1" or "2" or "3" or "4")
            .ToList();
        Assert.Equal(new[] { "1", "2", "3", "4" }, phaseNumbers);

        var allText = string.Join(" ", roadmap.Descendants<Text>().Select(t => t.Text));
        Assert.Contains("Ship the reliability hardening workstream", allText);
        Assert.Contains("Zero open Sev-1 issues", allText);
    }

    [Fact]
    public void Report_DocumentTitleAndRolesAreWritten()
    {
        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(DocxTestHarness.ReadEditorialReport());
        var mainPart = DocxTestHarness.OpenMainPart(path);
        var properties = mainPart.OpenXmlPackage.PackageProperties;

        Assert.Equal("Northwind Labs — Quarterly Editorial Report", properties.Title);
        Assert.Equal("OfficeEditor", properties.Creator);
        Assert.Equal("en-US", properties.Language);
    }

    [Fact]
    public void ExpansionWarningsFlowIntoGenerationResult()
    {
        var json = """
            {
              "version": "1.0",
              "sections": [ { "blocks": [
                { "type": "kpiRow", "items": [ { "value": "1", "label": "Only one" } ] }
              ] } ]
            }
            """;

        var bytes = DocxTestHarness.GenerateToBytes(json);

        var warning = Assert.Single(bytes.Result.Warnings);
        Assert.Equal("$.sections[0].blocks[0].items", warning.Path);
        Assert.Contains("recommended budget is 2–4", warning.Message);
    }

    [Fact]
    public void ExistingV10Document_PassesThroughExpansionUnchanged()
    {
        var document = DocxTestHarness.Parse(DocxTestHarness.ReadBuiltInSample());

        // The sample contains no archetypes: expansion is an identity and adds no warnings.
        var expansion = DocxGenerationExpander.ExpandWithIssues(document);
        Assert.Same(document, expansion.Document);
        Assert.Empty(expansion.Warnings);

        // Generation still succeeds with exactly the pre-existing emitter warnings (the
        // sample's DPI-less PNG), proving the expansion stage changed nothing for v1.0 docs.
        var generated = DocxTestHarness.GenerateToBytes(DocxTestHarness.ReadBuiltInSample());
        Assert.True(generated.Content.Length > 4_000);
        Assert.Contains(generated.Result.Warnings, w => w.Message.Contains("IntrinsicDpiUnavailable"));
        Assert.DoesNotContain(generated.Result.Warnings, w => w.Message.Contains("budget"));
    }
}
