namespace DocxEditor.Core.Generation.Model;

/// <summary>
/// Semantic report archetypes. These are authoring-level flow blocks: they name what a piece
/// of a report <em>is</em> (a cover, a KPI band, a semantic section, a comparison, a roadmap)
/// rather than how it renders. A dedicated expansion stage
/// (<c>DocxEditor.Core.Generation.Expand.DocxGenerationExpander</c>) lowers every archetype
/// into the vocabulary's concrete flow blocks (paragraphs, headings, tables, page breaks)
/// before emission, so the emitter never sees an archetype. Archetypes are pure authoring
/// sugar over the existing flow tier and add no new rendering machinery.
/// </summary>
public abstract record ReportArchetype : FlowBlock;

/// <summary>
/// A document cover: eyebrow kicker, title, optional subtitle, optional metadata line,
/// optional 2–4 KPI items and an optional trailing page break. It expands to a title block
/// (eyebrow/title/subtitle/metadata paragraphs on the theme's editorial roles) and, when
/// KPIs are present, a single KPI band table.
/// </summary>
public sealed record CoverBlock : ReportArchetype
{
    /// <summary>Small kicker line above the title (role <see cref="TextRole.Eyebrow"/>). Optional.</summary>
    public TextModel? Eyebrow { get; init; }

    /// <summary>Cover title (role <see cref="TextRole.Title"/>). Required.</summary>
    public required TextModel Title { get; init; }

    /// <summary>Subtitle line (role <see cref="TextRole.Subtitle"/>). Optional.</summary>
    public TextModel? Subtitle { get; init; }

    /// <summary>Optional metadata line (role <see cref="TextRole.Muted"/>), e.g. "Prepared by · date".</summary>
    public TextModel? Metadata { get; init; }

    /// <summary>KPI band items (budget 2–4; other counts warn). Optional.</summary>
    public IReadOnlyList<KpiItem>? Kpis { get; init; }

    /// <summary>True to append a page break after the cover.</summary>
    public bool PageBreak { get; init; }
}

/// <summary>
/// A standalone KPI band: 2–4 value/label/tone items expanded into a single pale band table
/// (one value + label column pair per item). A KPI <em>row</em> is the building block shared
/// by <see cref="CoverBlock.Kpis"/> and <see cref="KpiRowBlock.Items"/>.
/// </summary>
public sealed record KpiRowBlock : ReportArchetype
{
    /// <summary>KPI band items (budget 2–4; other counts warn). Required.</summary>
    public required IReadOnlyList<KpiItem> Items { get; init; }
}

/// <summary>One KPI band item: a value, a label and a semantic tone that tints the value.</summary>
public sealed record KpiItem
{
    /// <summary>The figure (role <see cref="TextRole.Metric"/>, tone-tinted). Required.</summary>
    public required TextModel Value { get; init; }

    /// <summary>The caption under the value (role <see cref="TextRole.MetricLabel"/>). Required.</summary>
    public required TextModel Label { get; init; }

    /// <summary>Semantic tone: positive/neutral/negative (drives the value color).</summary>
    public ReportTone Tone { get; init; } = ReportTone.Neutral;
}

/// <summary>
/// A semantic report section: a title, an optional intro and nested flow blocks. It expands to
/// a level-1 heading, the intro paragraph and the nested flow (which may itself contain further
/// archetypes, expanded recursively). It is <em>not</em> a DOCX page section — it is an
/// outline/reading unit inside one document section.
/// </summary>
public sealed record SemanticSectionBlock : ReportArchetype
{
    /// <summary>Section title (expands to a level-1 heading). Required.</summary>
    public required TextModel Title { get; init; }

    /// <summary>Optional intro paragraph under the title.</summary>
    public TextModel? Intro { get; init; }

    /// <summary>Nested flow blocks in document order. Required, at least one.</summary>
    public required IReadOnlyList<FlowBlock> Blocks { get; init; }
}

/// <summary>
/// A column comparison: column labels plus text-model rows, with an optional emphasized first
/// column. It expands to a table whose header row carries the labels and whose body rows carry
/// the cells; the first column uses the label role when <see cref="EmphasisFirstColumn"/> is set.
/// </summary>
public sealed record ComparisonTableBlock : ReportArchetype
{
    /// <summary>Column labels (one per column; budget 2–6, other counts warn). Required.</summary>
    public required IReadOnlyList<TextModel> Columns { get; init; }

    /// <summary>Body rows; every row's cell count must match the column count. Required.</summary>
    public required IReadOnlyList<ComparisonTableRow> Rows { get; init; }

    /// <summary>True to render the first column as a label column (accent, bold).</summary>
    public bool EmphasisFirstColumn { get; init; }
}

/// <summary>One body row of a <see cref="ComparisonTableBlock"/>: a cell per column.</summary>
public sealed record ComparisonTableRow
{
    /// <summary>Cells in column order. Required, at least one.</summary>
    public required IReadOnlyList<TextModel> Cells { get; init; }
}

/// <summary>
/// A delivery roadmap: phases with a window, an action, optional evidence and a tone. It expands
/// to a Phase / Window / Action / Evidence table; the phase number cell is tone-tinted.
/// </summary>
public sealed record RoadmapBlock : ReportArchetype
{
    /// <summary>Phases in document order (budget 1–6, other counts warn). Required.</summary>
    public required IReadOnlyList<RoadmapPhase> Phases { get; init; }
}

/// <summary>One roadmap phase.</summary>
public sealed record RoadmapPhase
{
    /// <summary>Delivery window, e.g. "Q1 2026". Required.</summary>
    public required TextModel Window { get; init; }

    /// <summary>What happens in this phase. Required.</summary>
    public required TextModel Action { get; init; }

    /// <summary>Optional evidence / exit criterion for the phase.</summary>
    public TextModel? Evidence { get; init; }

    /// <summary>Semantic tone of the phase (drives the phase-number color).</summary>
    public ReportTone Tone { get; init; } = ReportTone.Neutral;
}
