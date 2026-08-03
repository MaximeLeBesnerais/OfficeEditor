using DocxEditor.Core.Generation;
using DocxEditor.Core.Generation.Model;
using DocxEditor.Core.Generation.Schema;

namespace DocxEditor.Tests.Generation.Docx;

/// <summary>
/// Shared model validation: the semantic <see cref="DocxGenerationModelValidator"/> runs on
/// materialized models, not just JSON. It is internal to the core, so these tests exercise it
/// through the public <see cref="DocxGenerator"/> model overload, which validates then throws
/// <see cref="DocxGenerationValidationException"/> on any violation.
/// </summary>
public class DocxGenerationModelValidationTests
{
    private static DocxGenerationDocument ValidModel() => new()
    {
        Version = DocxGenerationDocument.SupportedVersion,
        Design = new DesignTokens
        {
            Palette = new Dictionary<string, string> { ["ink"] = "#1F2937", ["navy"] = "#1F4E79" }
        },
        Sections =
        [
            new Section
            {
                Blocks = [new ParagraphBlock { Content = new TextModel { Text = "Hello" } }]
            }
        ]
    };

    private static DocxGenerationValidationException ValidateModelThrows(DocxGenerationDocument document)
    {
        using var temp = new TempDirectory();
        var output = temp.File("out.docx");
        return Assert.Throws<DocxGenerationValidationException>(() =>
            new DocxGenerator().Generate(document, output));
    }

    private static void AssertNotCreated(DocxGenerationValidationException _, string output) =>
        Assert.False(File.Exists(output));

    [Fact]
    public void ValidModel_GeneratesSuccessfully()
    {
        using var temp = new TempDirectory();
        var output = temp.File("out.docx");
        var model = ValidModel();

        var result = new DocxGenerator().Generate(model, output);

        Assert.True(File.Exists(output));
        Assert.Same(model, result.Document);
    }

    [Fact]
    public void MissingVersion_IsRejected()
    {
        var model = ValidModel() with { Version = null! };
        var exception = ValidateModelThrows(model);

        Assert.Contains(exception.Issues, i => i.Path == "$.version" && i.Message.Contains("required"));
    }

    [Fact]
    public void UnsupportedVersion_IsRejected()
    {
        var model = ValidModel() with { Version = "3.0" };
        var exception = ValidateModelThrows(model);

        Assert.Contains(exception.Issues, i => i.Path == "$.version" && i.Message.Contains("unsupported version '3.0'"));
    }

    [Fact]
    public void NullSections_IsRejected()
    {
        var model = ValidModel() with { Sections = null! };
        var exception = ValidateModelThrows(model);

        Assert.Contains(exception.Issues, i => i.Path == "$.sections" && i.Message.Contains("at least one section"));
    }

    [Fact]
    public void EmptySections_IsRejected()
    {
        var model = ValidModel() with { Sections = [] };
        var exception = ValidateModelThrows(model);

        Assert.Contains(exception.Issues, i => i.Path == "$.sections" && i.Message.Contains("at least one section"));
    }

    [Fact]
    public void NullSectionElement_IsRejected()
    {
        var model = ValidModel() with
        {
            Sections = new List<Section> { null!, new Section { Blocks = [new ParagraphBlock { Content = new TextModel { Text = "x" } }] } }
        };
        var exception = ValidateModelThrows(model);

        Assert.Contains(exception.Issues, i => i.Path == "$.sections[0]" && i.Message.Contains("must not be null"));
    }

    [Fact]
    public void NullBlocks_IsRejected()
    {
        var model = ValidModel() with
        {
            Sections = [new Section { Blocks = null! }]
        };
        var exception = ValidateModelThrows(model);

        Assert.Contains(exception.Issues, i => i.Path == "$.sections[0].blocks" && i.Message.Contains("must not be null"));
    }

    [Fact]
    public void NullParagraphContent_IsRejected()
    {
        var model = ValidModel() with
        {
            Sections = [new Section { Blocks = [new ParagraphBlock { Content = null! }] }]
        };
        var exception = ValidateModelThrows(model);

        Assert.Contains(exception.Issues, i => i.Path == "$.sections[0].blocks[0]" && i.Message.Contains("text content is required"));
    }

    [Fact]
    public void HeadingLevelOutOfRange_IsRejected()
    {
        var model = ValidModel() with
        {
            Sections = [new Section { Blocks = [new HeadingBlock { Level = 7, Content = new TextModel { Text = "x" } }] }]
        };
        var exception = ValidateModelThrows(model);

        Assert.Contains(exception.Issues, i => i.Path == "$.sections[0].blocks[0].level" && i.Message.Contains("between 1 and 6"));
    }

    [Fact]
    public void RaggedTable_IsRejected()
    {
        var model = ValidModel() with
        {
            Sections = [new Section
            {
                Blocks =
                [
                    new TableBlock
                    {
                        Rows =
                        [
                            new TableRow { Cells = [new TableCell { Content = new TextModel { Text = "a" } }, new TableCell { Content = new TextModel { Text = "b" } }] },
                            new TableRow { Cells = [new TableCell { Content = new TextModel { Text = "c" } }] }
                        ]
                    }
                ]
            }]
        };
        var exception = ValidateModelThrows(model);

        Assert.Contains(exception.Issues, i => i.Path == "$.sections[0].blocks[0].rows[1]" && i.Message.Contains("row has 1 cells but the table has 2 columns"));
    }

    [Fact]
    public void EmptyTableRows_IsRejected()
    {
        var model = ValidModel() with
        {
            Sections = [new Section { Blocks = [new TableBlock { Rows = [] }] }]
        };
        var exception = ValidateModelThrows(model);

        Assert.Contains(exception.Issues, i => i.Path == "$.sections[0].blocks[0].rows" && i.Message.Contains("at least one row"));
    }

    [Fact]
    public void NullTableCell_IsRejected()
    {
        var model = ValidModel() with
        {
            Sections = [new Section
            {
                Blocks =
                [
                    new TableBlock { Rows = [new TableRow { Cells = new List<TableCell> { new TableCell { Content = new TextModel { Text = "a" } }, null! } }] }
                ]
            }]
        };
        var exception = ValidateModelThrows(model);

        Assert.Contains(exception.Issues, i => i.Path == "$.sections[0].blocks[0].rows[0].cells[1]" && i.Message.Contains("must not be null"));
    }

    [Fact]
    public void NullPaletteValue_IsRejected()
    {
        var model = ValidModel() with
        {
            Design = new DesignTokens
            {
                Palette = new Dictionary<string, string> { ["ink"] = null! }
            }
        };
        var exception = ValidateModelThrows(model);

        Assert.Contains(exception.Issues, i => i.Path == "$.design.palette.ink" && i.Message.Contains("#RRGGBB"));
    }

    [Fact]
    public void InvalidEnumValue_IsRejected()
    {
        var model = ValidModel() with
        {
            Sections = [new Section
            {
                Blocks =
                [
                    new ImageElement { Source = "data:image/png;base64,AAAA", Fit = (ImageFitMode)99 }
                ]
            }]
        };
        var exception = ValidateModelThrows(model);

        Assert.Contains(exception.Issues, i => i.Path == "$.sections[0].blocks[0].fit" && i.Message.Contains("not a valid image fit mode"));
    }

    [Fact]
    public void NullImageSource_IsRejected()
    {
        var model = ValidModel() with
        {
            Sections = [new Section { Blocks = [new ImageElement { Source = null! }] }]
        };
        var exception = ValidateModelThrows(model);

        Assert.Contains(exception.Issues, i => i.Path == "$.sections[0].blocks[0].src" && i.Message.Contains("required"));
    }

    [Fact]
    public void NegativeMargins_AreRejected()
    {
        var model = ValidModel() with
        {
            Sections = [new Section
            {
                PageSetup = new PageSetup { Margins = new Margins { TopPt = -1 } },
                Blocks = [new ParagraphBlock { Content = new TextModel { Text = "x" } }]
            }]
        };
        var exception = ValidateModelThrows(model);

        Assert.Contains(exception.Issues, i => i.Path == "$.sections[0].pageSetup.margins.top" && i.Message.Contains("≥ 0"));
    }

    [Fact]
    public void PositionedLineWithHeight_IsRejected()
    {
        var model = ValidModel() with
        {
            Sections = [new Section
            {
                Blocks = [new ParagraphBlock { Content = new TextModel { Text = "x" } }],
                Positioned =
                [
                    new PositionedLine
                    {
                        Position = new PositionSpec { X = 0, Y = 0, WidthPt = 100, HeightPt = 50 }
                    }
                ]
            }]
        };
        var exception = ValidateModelThrows(model);

        Assert.Contains(exception.Issues, i => i.Path == "$.sections[0].positioned[0].height" && i.Message.Contains("a line has no height"));
    }

    [Fact]
    public void PositionedBoxWithoutWidth_IsRejected()
    {
        var model = ValidModel() with
        {
            Sections = [new Section
            {
                Blocks = [new ParagraphBlock { Content = new TextModel { Text = "x" } }],
                Positioned =
                [
                    new PositionedRectangle
                    {
                        Position = new PositionSpec { X = 0, Y = 0, HeightPt = 10 }
                    }
                ]
            }]
        };
        var exception = ValidateModelThrows(model);

        Assert.Contains(exception.Issues, i => i.Path == "$.sections[0].positioned[0].width" && i.Message.Contains("required"));
    }

    [Fact]
    public void NegativeCropFraction_IsRejected()
    {
        var model = ValidModel() with
        {
            Sections = [new Section
            {
                Blocks =
                [
                    new ImageElement
                    {
                        Source = "data:image/png;base64,AAAA",
                        Crop = new ImageCrop { Left = -0.1 }
                    }
                ]
            }]
        };
        var exception = ValidateModelThrows(model);

        Assert.Contains(exception.Issues, i => i.Path == "$.sections[0].blocks[0].crop.left" && i.Message.Contains("≥ 0"));
    }

    [Fact]
    public void InvalidSpacingLineMultiple_IsRejected()
    {
        var model = ValidModel() with
        {
            Sections = [new Section
            {
                Blocks =
                [
                    new ParagraphBlock
                    {
                        Content = new TextModel
                        {
                            Text = "x",
                            Spacing = new ParagraphSpacing { LineMultiple = -2 }
                        }
                    }
                ]
            }]
        };
        var exception = ValidateModelThrows(model);

        Assert.Contains(exception.Issues, i => i.Path == "$.sections[0].blocks[0].spacing.line" && i.Message.Contains("> 0"));
    }

    [Fact]
    public void FirstSectionBreakType_WarnsButGenerates()
    {
        var model = ValidModel() with
        {
            Sections = [new Section
            {
                PageSetup = new PageSetup { BreakType = SectionBreakType.NextPage },
                Blocks = [new ParagraphBlock { Content = new TextModel { Text = "x" } }]
            }]
        };

        using var temp = new TempDirectory();
        var output = temp.File("out.docx");
        var result = new DocxGenerator().Generate(model, output);

        Assert.True(File.Exists(output));
        Assert.Contains(result.Warnings, w => w.Path == "$.sections[0].pageSetup.breakType" && w.Message.Contains("ignored"));
    }

    [Fact]
    public void OffPaletteColor_WarnsButGenerates()
    {
        var model = ValidModel() with
        {
            Sections = [new Section
            {
                Blocks =
                [
                    new ParagraphBlock
                    {
                        Content = new TextModel
                        {
                            Runs = [new Run { Text = "x", Color = "#123ABC" }]
                        }
                    }
                ]
            }]
        };

        using var temp = new TempDirectory();
        var output = temp.File("out.docx");
        var result = new DocxGenerator().Generate(model, output);

        Assert.True(File.Exists(output));
        Assert.Contains(result.Warnings, w => w.Path == "$.sections[0].blocks[0].runs[0].color" && w.Message.Contains("off-palette"));
    }

    [Fact]
    public void ValidationException_ReportsEveryIssueInMessage()
    {
        var model = ValidModel() with
        {
            Version = "2.0",
            Sections = [new Section { Blocks = [new HeadingBlock { Level = 9, Content = null! }] }]
        };
        var exception = ValidateModelThrows(model);

        Assert.True(exception.Issues.Count >= 2);
        Assert.Contains("$.version", exception.Message);
        Assert.Contains("$.sections[0].blocks[0]", exception.Message);
        Assert.Contains("Invalid DOCX generation document", exception.Message);
        Assert.Contains("errors):", exception.Message);
    }

    [Fact]
    public void SingleIssue_UsesSingularHeader()
    {
        var model = ValidModel() with { Version = "5.0" };
        var exception = ValidateModelThrows(model);

        Assert.Contains("(1 error)", exception.Message);
    }

    // ------------------------------------------------------------------ programmatic archetypes

    [Fact]
    public void EmptyKpiRowItems_AreRejectedProgrammatically()
    {
        var model = ValidModel() with
        {
            Sections = [new Section { Blocks = [new KpiRowBlock { Items = [] }] }]
        };
        var exception = ValidateModelThrows(model);

        Assert.Contains(exception.Issues, i =>
            i.Path == "$.sections[0].blocks[0].items" && i.Message.Contains("at least one KPI item"));
    }

    [Fact]
    public void EmptyCoverKpis_AreRejectedProgrammatically()
    {
        var model = ValidModel() with
        {
            Sections = [new Section
            {
                Blocks =
                [
                    new CoverBlock { Title = new TextModel { Text = "T" }, Kpis = [] }
                ]
            }]
        };
        var exception = ValidateModelThrows(model);

        Assert.Contains(exception.Issues, i =>
            i.Path == "$.sections[0].blocks[0].kpis" && i.Message.Contains("at least one KPI item"));
    }

    [Fact]
    public void EmptyComparisonColumns_AreRejectedProgrammatically()
    {
        var model = ValidModel() with
        {
            Sections = [new Section
            {
                Blocks =
                [
                    new ComparisonTableBlock
                    {
                        Columns = [],
                        Rows = [new ComparisonTableRow { Cells = [new TextModel { Text = "x" }] }]
                    }
                ]
            }]
        };
        var exception = ValidateModelThrows(model);

        Assert.Contains(exception.Issues, i =>
            i.Path == "$.sections[0].blocks[0].columns" && i.Message.Contains("at least one column"));
    }

    [Fact]
    public void EmptyComparisonRows_AreRejectedProgrammatically()
    {
        var model = ValidModel() with
        {
            Sections = [new Section
            {
                Blocks =
                [
                    new ComparisonTableBlock { Columns = [new TextModel { Text = "c" }], Rows = [] }
                ]
            }]
        };
        var exception = ValidateModelThrows(model);

        Assert.Contains(exception.Issues, i =>
            i.Path == "$.sections[0].blocks[0].rows" && i.Message.Contains("at least one row"));
    }

    [Fact]
    public void EmptyRoadmapPhases_AreRejectedProgrammatically()
    {
        var model = ValidModel() with
        {
            Sections = [new Section { Blocks = [new RoadmapBlock { Phases = [] }] }]
        };
        var exception = ValidateModelThrows(model);

        Assert.Contains(exception.Issues, i =>
            i.Path == "$.sections[0].blocks[0].phases" && i.Message.Contains("at least one phase"));
    }

    [Fact]
    public void EmptySemanticSectionBlocks_AreRejectedProgrammatically()
    {
        var model = ValidModel() with
        {
            Sections = [new Section
            {
                Blocks =
                [
                    new SemanticSectionBlock { Title = new TextModel { Text = "T" }, Blocks = [] }
                ]
            }]
        };
        var exception = ValidateModelThrows(model);

        Assert.Contains(exception.Issues, i =>
            i.Path == "$.sections[0].blocks[0].blocks" && i.Message.Contains("at least one flow block"));
    }

    [Fact]
    public void ArchetypeDocument_ExpandsAndRevalidatesToAValidPackage()
    {
        var model = ValidModel() with
        {
            Sections = [new Section
            {
                Blocks =
                [
                    new CoverBlock
                    {
                        Title = new TextModel { Text = "State of the Product" },
                        Kpis =
                        [
                            new KpiItem { Value = new TextModel { Text = "12.4k" }, Label = new TextModel { Text = "Active" } },
                            new KpiItem { Value = new TextModel { Text = "98%" }, Label = new TextModel { Text = "On time" } }
                        ]
                    },
                    new SemanticSectionBlock
                    {
                        Title = new TextModel { Text = "Outlook" },
                        Blocks =
                        [
                            new ParagraphBlock { Content = new TextModel { Text = "Body copy." } },
                            new RoadmapBlock
                            {
                                Phases =
                                [
                                    new RoadmapPhase { Window = new TextModel { Text = "Aug" }, Action = new TextModel { Text = "Ship" } }
                                ]
                            }
                        ]
                    }
                ]
            }]
        };

        // The lowered model is revalidated inside the generator before emission and must pass.
        var generated = new DocxGenerator().GenerateToBytes(model);

        Assert.True(generated.Content.Length > 1_000);
        Assert.Equal((byte)'P', generated.Content[0]);
        // Generation reports the expanded model, with every archetype lowered.
        var blocks = generated.Result.Document.Sections[0].Blocks;
        Assert.DoesNotContain(blocks, b => b is ReportArchetype);
        Assert.Contains(blocks, b => b is TableBlock);
    }
}
