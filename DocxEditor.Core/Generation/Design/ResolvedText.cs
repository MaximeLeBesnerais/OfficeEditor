using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Design;

/// <summary>
/// Fully-resolved text content for a paragraph-like block (paragraph, heading, list item, callout,
/// table cell, text box). <see cref="Text"/> holds plain text when the source used <c>text</c>;
/// <see cref="Runs"/> holds the resolved runs otherwise (each run's formatting already overlays the
/// resolved typography token). <see cref="TokenFormatting"/> exposes the resolved token on its own
/// so the emitter can apply it to the paragraph mark, and <see cref="Paragraph"/> carries the
/// alignment/spacing. Emitters write this straight out without re-interpreting tokens.
/// </summary>
public sealed record ResolvedText
{
    /// <summary>Empty content (blank table cell, no runs).</summary>
    public static ResolvedText Empty { get; } = new();

    /// <summary>Plain text content. Null when <see cref="Runs"/> is populated.</summary>
    public string? Text { get; init; }

    /// <summary>Resolved runs; each formatting is already composed with the typography token.</summary>
    public IReadOnlyList<ResolvedRun> Runs { get; init; } = [];

    /// <summary>Resolved typography-token formatting (empty when the content has no token).</summary>
    public ResolvedRunFormatting TokenFormatting { get; init; } = ResolvedRunFormatting.Empty;

    /// <summary>Resolved paragraph formatting (alignment and spacing).</summary>
    public ResolvedParagraphFormatting Paragraph { get; init; } = ResolvedParagraphFormatting.Empty;

    /// <summary>True when plain text content is present.</summary>
    public bool HasText => Text is not null;

    /// <summary>True when runs are present.</summary>
    public bool HasRuns => Runs.Count > 0;
}

/// <summary>One resolved run: its text plus the final, fully-resolved run formatting.</summary>
public sealed record ResolvedRun(string Text, ResolvedRunFormatting Formatting);
