namespace DocxEditor.Core.Generation.Model;

/// <summary>
/// One document section: optional page setup, optional header/footer flow, required body
/// flow blocks and an optional positioned (floating) tier. This mirrors the DOCX body →
/// <c>sectPr</c> model: a section is the unit that owns page geometry and breaks.
/// </summary>
public sealed record Section
{
    /// <summary>
    /// Page size, orientation, margins, columns and the section break applied before this
    /// section. Null = resolved from design tokens (or built-in A4/portrait/1in defaults).
    /// </summary>
    public PageSetup? PageSetup { get; init; }

    /// <summary>Optional header flow blocks (page-top content repeated for the section).</summary>
    public IReadOnlyList<FlowBlock> Header { get; init; } = [];

    /// <summary>Optional footer flow blocks (page-bottom content repeated for the section).</summary>
    public IReadOnlyList<FlowBlock> Footer { get; init; } = [];

    /// <summary>Body flow blocks in document order. Required, at least one.</summary>
    public required IReadOnlyList<FlowBlock> Blocks { get; init; }

    /// <summary>
    /// Optional positioned tier: absolutely anchored primitives (text boxes, shapes, lines,
    /// floating pictures, callouts) that float over the section's flow content.
    /// </summary>
    public IReadOnlyList<PositionedElement> Positioned { get; init; } = [];
}
