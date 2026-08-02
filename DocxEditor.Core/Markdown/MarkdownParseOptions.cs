using Markdig;
using Markdig.Extensions.Emoji;

namespace DocxEditor.Core.Markdown;

/// <summary>
/// Options controlling rich markdown parsing. Defaults enable the advanced Markdig
/// extensions plus YAML front matter and emoji shortcodes, and retain source spans.
/// </summary>
public sealed class MarkdownParseOptions
{
    /// <summary>
    /// When true, constructs that are retained verbatim rather than fully structured
    /// (raw HTML, unknown blocks/inlines, alerts collapsed to quotes, ...) produce
    /// <see cref="MarkdownDiagnosticSeverity.Warning"/> diagnostics. When false
    /// (permissive) they are retained as explicit nodes without diagnostics.
    /// </summary>
    public bool Strict { get; init; }

    /// <summary>Enables the Markdig advanced extension bundle (tables, footnotes, definition lists,
    /// task lists, emphasis extras, custom containers, ...).</summary>
    public bool UseAdvancedExtensions { get; init; } = true;

    /// <summary>Enables YAML front matter parsing at the start of the document.</summary>
    public bool UseYamlFrontMatter { get; init; } = true;

    /// <summary>Enables emoji shortcode expansion (shortcode-only mapping so pipe-table
    /// separators such as ":---" are not misinterpreted as smileys).</summary>
    public bool UseEmoji { get; init; } = true;

    /// <summary>Retains character spans and line numbers from the source where Markdig exposes them.</summary>
    public bool UseSourceSpans { get; init; } = true;

    public static MarkdownParseOptions Default { get; } = new();

    public static MarkdownParseOptions StrictMode { get; } = new() { Strict = true };

    /// <summary>Builds the Markdig pipeline for these options. Safe to call on any thread;
    /// the returned pipeline is immutable after <c>Build()</c>.</summary>
    internal Markdig.MarkdownPipeline BuildPipeline()
    {
        var builder = new Markdig.MarkdownPipelineBuilder();
        if (UseAdvancedExtensions)
        {
            builder.UseAdvancedExtensions();
        }

        if (UseYamlFrontMatter)
        {
            builder.UseYamlFrontMatter();
        }

        if (UseEmoji)
        {
            builder.UseEmojiAndSmiley(new EmojiMapping(false));
        }

        if (UseSourceSpans)
        {
            builder.UsePragmaLines();
        }

        return builder.Build();
    }
}
