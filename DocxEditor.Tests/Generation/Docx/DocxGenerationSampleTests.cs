using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Generation;
using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Tests.Generation.Docx;

/// <summary>
/// Built-in sample (<c>examples/Docx/generation/comprehensive.json</c>): parse to a model,
/// generate a package, reopen it and verify the required OOXML structure (sections, page
/// geometry, headers/footers, numbering, tables, callouts, images, positioned anchors,
/// metadata) and package relationships.
/// </summary>
public class DocxGenerationSampleTests
{
    [Fact]
    public void Sample_ParsesWithNoErrorsOrWarnings()
    {
        var result = DocxTestHarness.Validate(DocxTestHarness.ReadBuiltInSample());

        Assert.True(result.IsValid);
        Assert.Empty(result.Errors);
        Assert.Empty(result.Warnings);

        var document = result.Document!;
        Assert.Equal("1.0", document.Version);
        Assert.NotNull(document.Metadata);
        Assert.Equal("Northwind Operations Brief", document.Metadata!.Title);
        Assert.Equal("OfficeEditor", document.Metadata.Author);
        Assert.Equal("en-US", document.Metadata.Language);
        Assert.NotNull(document.Design);
        Assert.Equal(8, document.Design!.Palette.Count);
        Assert.Equal(4, document.Design.Typography.Count);
        Assert.Equal(2, document.Sections.Count);
    }

    [Fact]
    public void Sample_SectionsHaveExpectedShape()
    {
        var document = DocxTestHarness.Parse(DocxTestHarness.ReadBuiltInSample());
        var first = document.Sections[0];

        Assert.Equal(11, first.Blocks.Count);
        Assert.Equal(5, first.Positioned.Count);
        Assert.Single(first.Header);
        Assert.Single(first.Footer);

        Assert.IsType<HeadingBlock>(first.Blocks[0]);
        Assert.IsType<ParagraphBlock>(first.Blocks[1]);
        Assert.IsType<ListBlock>(first.Blocks[3]);
        Assert.IsType<TableBlock>(first.Blocks[4]);
        Assert.IsType<CalloutBlock>(first.Blocks[5]);
        Assert.IsType<ImageElement>(first.Blocks[6]);
        Assert.IsType<FlowContainerBlock>(first.Blocks[7]);
        Assert.IsType<PageBreakBlock>(first.Blocks[8]);

        var list = Assert.IsType<ListBlock>(first.Blocks[3]);
        Assert.Equal(ListKind.Ordered, list.Kind);
        Assert.Equal(3, list.StartIndex);

        var table = Assert.IsType<TableBlock>(first.Blocks[4]);
        Assert.Equal(4, table.Rows.Count);
        Assert.True(table.Rows[0].IsHeader);

        var image = Assert.IsType<ImageElement>(first.Blocks[6]);
        Assert.StartsWith("data:image/png;base64,", image.Source);
        Assert.Equal(ImageFitMode.Contain, image.Fit);

        var second = document.Sections[1];
        Assert.Equal(SectionBreakType.NextPage, second.PageSetup!.BreakType);
        Assert.Equal(PageOrientation.Landscape, second.PageSetup!.Orientation);
        Assert.Equal(2, second.PageSetup.Columns!.Count);
        Assert.True(second.PageSetup.Columns.SeparatorLine);
    }

    [Fact]
    public void Sample_GeneratesToBytesAsZippedPackage()
    {
        var bytes = DocxTestHarness.GenerateToBytes(DocxTestHarness.ReadBuiltInSample());

        Assert.NotNull(bytes.Content);
        Assert.True(bytes.Content.Length > 4_000, $"expected a substantial package, got {bytes.Content.Length} bytes");
        Assert.Equal((byte)'P', bytes.Content[0]);
        Assert.Equal((byte)'K', bytes.Content[1]);
        Assert.Empty(bytes.Result.Outputs); // byte mode writes no file artifact
        // The sample's PNG carries no pHYs chunk, so the loader warns about the 96-DPI fallback.
        Assert.Contains(bytes.Result.Warnings, w => w.Message.Contains("IntrinsicDpiUnavailable"));
    }

    [Fact]
    public void Sample_ReopensWithTwoSectionProperties()
    {
        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(DocxTestHarness.ReadBuiltInSample(), new DocxGeneratorOptions());
        var mainPart = DocxTestHarness.OpenMainPart(path);
        var body = mainPart.Document!.Body!;

        // One intermediate section break attached to a paragraph's pPr, one body-level sectPr.
        var intermediate = body.Elements<Paragraph>()
            .Where(p => p.ParagraphProperties?.SectionProperties is not null)
            .ToList();
        Assert.Single(intermediate);

        var bodyLevel = body.Elements<SectionProperties>().ToList();
        Assert.Single(bodyLevel);

        // Section 1 (custom 612×792, landscape): pgSz is swapped, margins 45pt → 900 twips.
        var finalProperties = bodyLevel[0];
        var finalPageSize = finalProperties.Elements<DocumentFormat.OpenXml.Wordprocessing.PageSize>().Single();
        Assert.Equal("15840", finalPageSize.Width!.Value.ToString());
        Assert.Equal("12240", finalPageSize.Height!.Value.ToString());
        Assert.Equal(PageOrientationValues.Landscape, finalPageSize.Orient!.Value);
        Assert.Equal("900", finalProperties.Elements<PageMargin>().Single().Top!.Value.ToString());

        // The incoming break of section 1 is recorded as NextPage on section 0's sectPr.
        var firstProperties = intermediate[0].ParagraphProperties!.SectionProperties!;
        Assert.Equal(SectionMarkValues.NextPage, firstProperties.Elements<SectionType>().Single().Val!.Value);

        // Section 0 (letter, portrait): margins 54pt → 1080 twips.
        var firstPageSize = firstProperties.Elements<DocumentFormat.OpenXml.Wordprocessing.PageSize>().Single();
        Assert.Equal("12240", firstPageSize.Width!.Value.ToString());
        Assert.Equal("15840", firstPageSize.Height!.Value.ToString());
        Assert.Equal(PageOrientationValues.Portrait, firstPageSize.Orient!.Value);
        Assert.Equal("1080", firstProperties.Elements<PageMargin>().Single().Left!.Value.ToString());
    }

    [Fact]
    public void Sample_HeadersFootersAreEmittedPerSection()
    {
        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(DocxTestHarness.ReadBuiltInSample(), new DocxGeneratorOptions());
        var mainPart = DocxTestHarness.OpenMainPart(path);

        // Every section gets a header and footer part (even empty ones).
        Assert.Equal(2, mainPart.HeaderParts.Count());
        Assert.Equal(2, mainPart.FooterParts.Count());

        var headerText = string.Join(
            " ",
            mainPart.HeaderParts.SelectMany(h => h.Header!.Descendants<Text>()).Select(t => t.Text));
        Assert.Contains("NORTHWIND / OPERATIONS", headerText);

        var footerText = string.Join(
            " ",
            mainPart.FooterParts.SelectMany(f => f.Footer!.Descendants<Text>()).Select(t => t.Text));
        Assert.Contains("Confidential", footerText);
    }

    [Fact]
    public void Sample_NumberingHasOrderedAndBulletAbstracts()
    {
        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(DocxTestHarness.ReadBuiltInSample(), new DocxGeneratorOptions());
        var mainPart = DocxTestHarness.OpenMainPart(path);

        var numberingPart = mainPart.NumberingDefinitionsPart;
        Assert.NotNull(numberingPart);
        var abstracts = numberingPart!.Numbering!.Elements<AbstractNum>().ToList();
        Assert.Equal(2, abstracts.Count);
        Assert.Equal(2, numberingPart.Numbering.Elements<NumberingInstance>().Count());

        var ordered = abstracts[0];
        Assert.Equal(0, ordered.AbstractNumberId!.Value);
        var level = Assert.Single(ordered.Elements<Level>());
        Assert.Equal("3", level.StartNumberingValue!.Val!.Value.ToString());
        Assert.Equal(NumberFormatValues.Decimal, level.NumberingFormat!.Val!.Value);
        Assert.Equal("%1.", level.LevelText!.Val!.Value);

        var bullet = abstracts[1];
        Assert.Equal(1, bullet.AbstractNumberId!.Value);
        var bulletLevel = Assert.Single(bullet.Elements<Level>());
        Assert.Equal(NumberFormatValues.Bullet, bulletLevel.NumberingFormat!.Val!.Value);
        Assert.Equal("•", bulletLevel.LevelText!.Val!.Value);
    }

    [Fact]
    public void Sample_FlowContentIsPresent()
    {
        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(DocxTestHarness.ReadBuiltInSample(), new DocxGeneratorOptions());
        var mainPart = DocxTestHarness.OpenMainPart(path);
        var allText = string.Join(" ", mainPart.Document!.Descendants<Text>().Select(t => t.Text));

        Assert.Contains("Operations Brief", allText);
        Assert.Contains("Priority actions", allText);
        Assert.Contains("Two-column appendix", allText);
        Assert.Contains("Decision needed: approve the alternate carrier", allText);

        // Page break: at least one w:br of type page.
        Assert.Contains(mainPart.Document.Descendants<Break>(), b => b.Type?.Value == BreakValues.Page);

        // Table: 4 rows; the first carries a repeating header marker.
        var tables = mainPart.Document.Body!.Descendants<Table>().ToList();
        Assert.NotEmpty(tables);
        var contentTable = tables.First(t => t.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCell>().Count() >= 12);
        Assert.Equal(4, contentTable.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>().Count());
        var headerRow = contentTable.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>().First();
        Assert.True(headerRow.TableRowProperties?.Elements<TableHeader>().Any() == true);

        // Callout: a single-cell table with a thick accent left border (24 eighths of a point).
        var calloutTable = tables.First(t =>
            t.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>().Count() == 1 &&
            t.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>().Single().Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>().Count() == 1);
        var calloutCell = calloutTable.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>().Single().Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>().Single();
        var shading = calloutCell.TableCellProperties?.Shading;
        Assert.NotNull(shading);
        Assert.Equal("FFF8E1", shading!.Fill?.Value); // warning tone fill
        var leftBorder = calloutTable.TableProperties?.TableBorders?.LeftBorder;
        Assert.NotNull(leftBorder);
        Assert.Equal("24", leftBorder!.Size?.Value.ToString());
        Assert.Equal("F9A825", leftBorder.Color!.Value); // warning tone accent

        // Group flattening: the group's heading lands as a sibling paragraph in the body.
        Assert.Contains("Grouped follow-up", allText);
    }

    [Fact]
    public void Sample_ImagesDeduplicateToOnePart()
    {
        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(DocxTestHarness.ReadBuiltInSample(), new DocxGeneratorOptions());
        var mainPart = DocxTestHarness.OpenMainPart(path);

        var imageParts = mainPart.ImageParts.ToList();
        Assert.Single(imageParts);
        Assert.Equal("image/png", imageParts[0].ContentType);
    }

    [Fact]
    public void Sample_PositionedAnchorsArePresent()
    {
        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(DocxTestHarness.ReadBuiltInSample(), new DocxGeneratorOptions());
        var mainPart = DocxTestHarness.OpenMainPart(path);

        var anchors = mainPart.Document!.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor>().ToList();
        Assert.Equal(5, anchors.Count);

        // Z-order paints low to high: the behind-document band is first, the image last.
        Assert.True(anchors[0].BehindDoc!.Value);
        Assert.False(anchors[^1].BehindDoc!.Value);
    }

    [Fact]
    public void Sample_DrawingIdsAreUniqueAcrossAllParts()
    {
        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(DocxTestHarness.ReadBuiltInSample(), new DocxGeneratorOptions());
        var mainPart = DocxTestHarness.OpenMainPart(path);

        var ids = new List<uint>();
        void Collect(OpenXmlElement root)
        {
            foreach (var element in root.Descendants())
            {
                if (element.LocalName is not ("docPr" or "cNvPr"))
                {
                    continue;
                }
                if (uint.TryParse(element.GetAttribute("id", string.Empty).Value, out var id))
                {
                    ids.Add(id);
                }
            }
        }

        Collect(mainPart.Document!);
        foreach (var header in mainPart.HeaderParts)
        {
            Collect(header.Header!);
        }
        foreach (var footer in mainPart.FooterParts)
        {
            Collect(footer.Footer!);
        }

        Assert.NotEmpty(ids);
        Assert.Equal(ids.Count, ids.Distinct().Count());
        Assert.Equal(ids.Count, ids.Count(id => id > 0));
    }

    [Fact]
    public void Sample_InlineImageCarriesRelationship()
    {
        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(DocxTestHarness.ReadBuiltInSample(), new DocxGeneratorOptions());
        var mainPart = DocxTestHarness.OpenMainPart(path);

        var drawings = mainPart.Document!.Descendants<Drawing>().ToList();
        Assert.NotEmpty(drawings);
        var inline = Assert.Single(mainPart.Document!.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline>());

        var extent = inline.Extent!;
        // Flow image is 36×36 pt → 457200 EMU per side (contain on a square 16×16 PNG).
        Assert.Equal("457200", extent.Cx!.Value.ToString());
        Assert.Equal("457200", extent.Cy!.Value.ToString());

        var blip = inline.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().Single();
        var embedId = blip.Embed!.Value;
        Assert.False(string.IsNullOrEmpty(embedId));
        Assert.NotNull(embedId);
        Assert.NotNull(mainPart.GetPartById(embedId!));
        Assert.True(Regex.IsMatch(embedId, "^image[0-9a-f]{12}$"), $"unexpected relationship id '{embedId}'");
    }

    [Fact]
    public void Sample_MetadataIsWrittenToCoreProperties()
    {
        using var temp = new TempDirectory();
        var path = DocxTestHarness.GenerateToTempFile(DocxTestHarness.ReadBuiltInSample(), new DocxGeneratorOptions());
        var mainPart = DocxTestHarness.OpenMainPart(path);
        var properties = mainPart.OpenXmlPackage.PackageProperties;

        Assert.Equal("Northwind Operations Brief", properties.Title);
        Assert.Equal("OfficeEditor", properties.Creator);
        Assert.Equal("Declarative DOCX generation example", properties.Subject);
        Assert.Equal("DOCX, JSON, OfficeEditor", properties.Keywords);
        Assert.Equal("en-US", properties.Language);
        Assert.Contains("self-contained example", properties.Description);
    }
}
