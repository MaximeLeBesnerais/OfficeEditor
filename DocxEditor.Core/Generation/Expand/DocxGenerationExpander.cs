using DocxEditor.Core.Generation.Model;
using DocxEditor.Core.Generation.Schema;

namespace DocxEditor.Core.Generation.Expand;

/// <summary>
/// Outcome of expanding a document's report archetypes: the fully-lowered document (no
/// <see cref="ReportArchetype"/> blocks remain) and any non-fatal budget warnings.
/// </summary>
public sealed record DocxExpansionResult
{
    /// <summary>The expanded document. Identical to the input when it contained no archetypes.</summary>
    public required DocxGenerationDocument Document { get; init; }

    /// <summary>Path-qualified warnings (e.g. item counts outside the component budget).</summary>
    public required IReadOnlyList<DocxGenerationIssue> Warnings { get; init; }
}

/// <summary>
/// Pure expansion stage for the DOCX generation vocabulary. It lowers the semantic report
/// archetypes (<c>cover</c>, <c>kpiRow</c>, <c>section</c>, <c>comparisonTable</c>, <c>roadmap</c>)
/// into the vocabulary's concrete flow blocks (paragraphs, headings, tables, page breaks) using
/// the theme's semantic roles and editorial surfaces. It is deterministic and stateless: the
/// same input always yields the same output and the same warnings, the input model is never
/// mutated, and a document without archetypes passes through untouched.
///
/// Expansion is authoring sugar only — it introduces no positioned composition and no second
/// layout engine. <see cref="DocxGenerator"/> invokes it exactly once, after parse/validation
/// and before design resolution/emission.
/// </summary>
public static class DocxGenerationExpander
{
    /// <summary>Expands a document in place-identity terms: returns the same instance when no
    /// archetype is present, otherwise a lowered copy. Budget warnings are discarded; use
    /// <see cref="ExpandWithIssues"/> to surface them.</summary>
    public static DocxGenerationDocument Expand(DocxGenerationDocument document) =>
        ExpandWithIssues(document).Document;

    /// <summary>Expands a document and returns the lowered model together with its warnings.</summary>
    public static DocxExpansionResult ExpandWithIssues(DocxGenerationDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);
        var session = new ExpansionSession();
        var sections = session.ExpandSections(document.Sections);
        if (sections is null)
        {
            return new DocxExpansionResult { Document = document, Warnings = [] };
        }
        return new DocxExpansionResult
        {
            Document = document with { Sections = sections },
            Warnings = session.Warnings
        };
    }

    private sealed class ExpansionSession
    {
        private readonly List<DocxGenerationIssue> _warnings = [];

        public IReadOnlyList<DocxGenerationIssue> Warnings => _warnings;

        public IReadOnlyList<Section>? ExpandSections(IReadOnlyList<Section> sections)
        {
            List<Section>? result = null;
            for (var i = 0; i < sections.Count; i++)
            {
                var section = sections[i];
                var expanded = ExpandSection(section, $"$.sections[{i}]");
                if (ReferenceEquals(expanded, section))
                {
                    result?.Add(section);
                    continue;
                }
                result ??= sections.Take(i).ToList();
                result.Add(expanded);
            }
            return result;
        }

        private Section ExpandSection(Section section, string path)
        {
            var header = ExpandBlockList(section.Header, $"{path}.header");
            var footer = ExpandBlockList(section.Footer, $"{path}.footer");
            var blocks = ExpandBlockList(section.Blocks, $"{path}.blocks");
            if (header is null && footer is null && blocks is null)
            {
                return section;
            }
            return section with
            {
                Header = header ?? section.Header,
                Footer = footer ?? section.Footer,
                Blocks = blocks ?? section.Blocks
            };
        }

        /// <summary>
        /// Slices archetypes out of a block list, splicing their lowered blocks into place.
        /// Returns null when nothing changed so callers can keep the original reference.
        /// </summary>
        private IReadOnlyList<FlowBlock>? ExpandBlockList(IReadOnlyList<FlowBlock> blocks, string path)
        {
            List<FlowBlock>? result = null;
            for (var i = 0; i < blocks.Count; i++)
            {
                var lowered = ExpandBlock(blocks[i], $"{path}[{i}]");
                if (lowered is null)
                {
                    result?.Add(blocks[i]);
                    continue;
                }
                result ??= blocks.Take(i).ToList();
                result.AddRange(lowered);
            }
            return result;
        }

        /// <summary>Lowers one block. Returns null for blocks that need no expansion.</summary>
        private IReadOnlyList<FlowBlock>? ExpandBlock(FlowBlock block, string path)
        {
            switch (block)
            {
                case CoverBlock cover:
                    return LowerCover(cover, path);
                case KpiRowBlock kpiRow:
                    return LowerKpiRow(kpiRow, path);
                case SemanticSectionBlock section:
                    return LowerSemanticSection(section, path);
                case ComparisonTableBlock comparison:
                    return [LowerComparisonTable(comparison, path)];
                case RoadmapBlock roadmap:
                    return [LowerRoadmap(roadmap, path)];
                case FlowContainerBlock group:
                    var expanded = ExpandBlockList(group.Blocks, $"{path}.blocks");
                    return expanded is null ? null : [group with { Blocks = expanded }];
                default:
                    return null;
            }
        }

        private IReadOnlyList<FlowBlock> LowerCover(CoverBlock cover, string path)
        {
            if (cover.Kpis is { } kpis)
            {
                WarnBudget(kpis.Count, 2, 4, $"{path}.kpis", "KPI items", "cover");
            }
            var blocks = new List<FlowBlock>();
            if (cover.Eyebrow is { } eyebrow)
            {
                blocks.Add(Paragraph(eyebrow, TextRole.Eyebrow));
            }
            blocks.Add(Paragraph(cover.Title, TextRole.Title));
            if (cover.Subtitle is { } subtitle)
            {
                blocks.Add(Paragraph(subtitle, TextRole.Subtitle));
            }
            if (cover.Metadata is { } metadata)
            {
                blocks.Add(Paragraph(metadata, TextRole.Muted));
            }
            if (cover.Kpis is { Count: > 0 } kpiItems)
            {
                blocks.Add(BuildKpiTable(kpiItems));
            }
            if (cover.PageBreak)
            {
                blocks.Add(new PageBreakBlock());
            }
            return blocks;
        }

        private IReadOnlyList<FlowBlock> LowerKpiRow(KpiRowBlock kpiRow, string path)
        {
            WarnBudget(kpiRow.Items.Count, 2, 4, $"{path}.items", "KPI items", "kpiRow");
            return [BuildKpiTable(kpiRow.Items)];
        }

        private IReadOnlyList<FlowBlock> LowerSemanticSection(SemanticSectionBlock section, string path)
        {
            var blocks = new List<FlowBlock>
            {
                new HeadingBlock
                {
                    Level = 1,
                    Content = section.Title with { Role = section.Title.Role ?? TextRole.Heading1 }
                }
            };
            if (section.Intro is { } intro)
            {
                blocks.Add(Paragraph(intro, TextRole.Body));
            }
            if (section.Blocks is { Count: > 0 } nested)
            {
                var expanded = ExpandBlockList(nested, $"{path}.blocks");
                blocks.AddRange(expanded ?? nested);
            }
            return blocks;
        }

        private TableBlock LowerComparisonTable(ComparisonTableBlock comparison, string path)
        {
            WarnBudget(comparison.Columns.Count, 2, 6, $"{path}.columns", "columns", "comparisonTable");
            WarnBudget(comparison.Rows.Count, 1, 12, $"{path}.rows", "rows", "comparisonTable");

            var rows = new List<TableRow>();
            var headerCells = new List<TableCell>(comparison.Columns.Count);
            foreach (var column in comparison.Columns)
            {
                headerCells.Add(new TableCell { Content = column with { Role = column.Role ?? TextRole.TableHeader } });
            }
            rows.Add(new TableRow { Cells = headerCells, IsHeader = true });

            foreach (var row in comparison.Rows)
            {
                var cells = new List<TableCell>(row.Cells.Count);
                for (var c = 0; c < row.Cells.Count; c++)
                {
                    var content = row.Cells[c];
                    if (comparison.EmphasisFirstColumn && c == 0)
                    {
                        content = content with { Role = content.Role ?? TextRole.Label };
                    }
                    else
                    {
                        content = content with { Role = content.Role ?? TextRole.TableBody };
                    }
                    cells.Add(new TableCell { Content = content });
                }
                rows.Add(new TableRow { Cells = cells });
            }
            return new TableBlock { Rows = rows };
        }

        private TableBlock LowerRoadmap(RoadmapBlock roadmap, string path)
        {
            WarnBudget(roadmap.Phases.Count, 1, 6, $"{path}.phases", "phases", "roadmap");

            var rows = new List<TableRow>
            {
                new()
                {
                    IsHeader = true,
                    Cells =
                    [
                        HeaderCell("Phase"),
                        HeaderCell("Window"),
                        HeaderCell("Action"),
                        HeaderCell("Evidence")
                    ]
                }
            };
            for (var i = 0; i < roadmap.Phases.Count; i++)
            {
                var phase = roadmap.Phases[i];
                rows.Add(new TableRow
                {
                    Cells =
                    [
                        new TableCell
                        {
                            Content = new TextModel
                            {
                                Runs = [new Run { Text = (i + 1).ToString(), Color = ToneColor(phase.Tone) }],
                                Role = TextRole.TableBody
                            },
                            Alignment = TextAlignment.Center
                        },
                        new TableCell { Content = phase.Window with { Role = phase.Window.Role ?? TextRole.TableBody } },
                        new TableCell { Content = phase.Action with { Role = phase.Action.Role ?? TextRole.TableBody } },
                        new TableCell
                        {
                            Content = phase.Evidence is { } evidence
                                ? evidence with { Role = evidence.Role ?? TextRole.TableBody }
                                : null
                        }
                    ]
                });
            }
            return new TableBlock { Rows = rows };
        }

        /// <summary>
        /// Builds a KPI band: a two-row table (values on top, labels below) with one column per
        /// item, all cells on the theme's pale surface. Values use the metric role tone-tinted to
        /// the item's semantic tone; labels use the metric-label role.
        /// </summary>
        private static TableBlock BuildKpiTable(IReadOnlyList<KpiItem> items)
        {
            var valueCells = new List<TableCell>(items.Count);
            var labelCells = new List<TableCell>(items.Count);
            for (var i = 0; i < items.Count; i++)
            {
                var item = items[i];
                var value = ApplyTone(item.Value, item.Tone);
                valueCells.Add(new TableCell
                {
                    Content = value with { Role = value.Role ?? TextRole.Metric, Alignment = TextAlignment.Center },
                    Alignment = TextAlignment.Center,
                    Fill = "pale"
                });
                labelCells.Add(new TableCell
                {
                    Content = item.Label with { Role = item.Label.Role ?? TextRole.MetricLabel, Alignment = TextAlignment.Center },
                    Alignment = TextAlignment.Center,
                    Fill = "pale"
                });
            }
            return new TableBlock
            {
                Rows =
                [
                    new TableRow { Cells = valueCells },
                    new TableRow { Cells = labelCells }
                ]
            };
        }

        /// <summary>A paragraph carrying a semantic role (the author's role wins over the archetype default).</summary>
        private static ParagraphBlock Paragraph(TextModel content, TextRole fallbackRole) =>
            new() { Content = content with { Role = content.Role ?? fallbackRole } };

        private static TableCell HeaderCell(string text) =>
            new() { Content = new TextModel { Text = text, Role = TextRole.TableHeader } };

        /// <summary>
        /// Tints a metric's runs toward the item's tone (positive → theme teal, negative → theme
        /// coral, neutral → untouched). Palette token names are used so the theme resolves them.
        /// </summary>
        private static TextModel ApplyTone(TextModel content, ReportTone tone)
        {
            var color = ToneColor(tone);
            if (color is null)
            {
                return content;
            }
            if (content.Runs is { } runs)
            {
                var tinted = new List<Run>(runs.Count);
                foreach (var run in runs)
                {
                    tinted.Add(run.Color is null ? run with { Color = color } : run);
                }
                return content with { Runs = tinted };
            }
            return content with
            {
                Runs = [new Run { Text = content.Text ?? string.Empty, Color = color }]
            };
        }

        private static string? ToneColor(ReportTone tone) => tone switch
        {
            ReportTone.Positive => "teal",
            ReportTone.Negative => "coral",
            _ => null
        };

        /// <summary>
        /// Advisory budget guardrail: warns (never rejects) when a component's item/row count
        /// falls outside its recommended budget so authors can split or trim oversized surfaces.
        /// </summary>
        private void WarnBudget(int count, int min, int max, string path, string what, string component)
        {
            if (count < min || count > max)
            {
                Warn(path, $"'{component}' has {count} {what.ToLowerInvariant()}; the recommended budget is {min}–{max}. It still renders, but consider splitting or trimming.");
            }
        }

        private void Warn(string path, string message) =>
            _warnings.Add(new DocxGenerationIssue(path, message, null, DocxGenerationIssueSeverity.Warning));
    }
}
