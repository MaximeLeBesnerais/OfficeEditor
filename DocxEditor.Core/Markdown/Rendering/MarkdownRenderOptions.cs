using DocxEditor.Core.Generation.Assets;
using OfficeEditor.Core.Models;

namespace DocxEditor.Core.Markdown.Rendering;

/// <summary>
/// How a soft (single-newline) markdown line break is rendered inside a paragraph.
/// Hard breaks (two trailing spaces or a backslash) always produce an explicit
/// <c>&lt;w:br/&gt;</c> regardless of this setting.
/// </summary>
public enum MarkdownSoftBreakMode
{
    /// <summary>Rendered as a space — the common "soft wrap" interpretation.</summary>
    Space,

    /// <summary>Rendered as an explicit line break.</summary>
    LineBreak,

    /// <summary>Dropped entirely.</summary>
    None
}

/// <summary>
/// Options controlling rich markdown rendering into a DOCX document. Parsing is governed by
/// <see cref="ParseOptions"/> (defaults to the standard permissive configuration); this type
/// controls the OOXML half: style resolution, soft breaks, YAML core-property mapping, image
/// loading policy and strict/permissive diagnostics.
/// </summary>
public sealed record MarkdownRenderOptions
{
    /// <summary>
    /// Defaults: default parsing, <see cref="StyleMapping.Default"/> style resolution, soft
    /// breaks as spaces, YAML front matter mapped to core properties, images capped to a
    /// 6.5&Prime; display width, permissive diagnostics.
    /// </summary>
    public static MarkdownRenderOptions Default { get; } = new();

    /// <summary>
    /// Controls the markdown parse (advanced extensions, YAML, emoji, source spans, and
    /// whether retained-verbatim constructs produce parse diagnostics). Null uses
    /// <see cref="MarkdownParseOptions.Default"/>.
    /// </summary>
    public MarkdownParseOptions? ParseOptions { get; init; }

    /// <summary>
    /// Resolves markdown element names ("heading1", "paragraph", "blockquote", "codeBlock", …)
    /// to document style ids. Null uses <see cref="StyleMapping.Default"/>. The mapping is read
    /// only; existing style definitions are never mutated.
    /// </summary>
    public StyleMapping? StyleMapping { get; init; } = StyleMapping.Default;

    /// <summary>
    /// Style resolver used to map the style references produced by <see cref="StyleMapping"/>
    /// to actual document styles. When null, the renderer builds a resolver over the document's
    /// styles part, using <see cref="Strict"/> to choose permissive (fallback-generated) or
    /// strict (error, no style) resolution. Generated fallback styles are appended to the
    /// document styles part so every emitted reference exists in the output; existing styles
    /// are never mutated. An injected resolver must be built over the same document styles part.
    /// </summary>
    public MarkdownStyleResolver? StyleResolver { get; init; }

    /// <summary>
    /// When true, render fallbacks (unresolved images, HTML and unknown nodes converted to
    /// visible text, non-absolute links, table spans) emit
    /// <see cref="MarkdownDiagnosticSeverity.Warning"/> diagnostics. When false, the visible
    /// fallback is still produced but no render diagnostic is recorded.
    /// </summary>
    public bool Strict { get; init; }

    /// <summary>How soft line breaks render. Defaults to <see cref="MarkdownSoftBreakMode.Space"/>.</summary>
    public MarkdownSoftBreakMode SoftBreakMode { get; init; } = MarkdownSoftBreakMode.Space;

    /// <summary>
    /// When true, a YAML front-matter block maps recognized title / author / subject /
    /// keywords / description / language keys into the package core properties and renders
    /// no visible content. Defaults to true.
    /// </summary>
    public bool MapYamlFrontMatterToCoreProperties { get; init; } = true;

    /// <summary>
    /// Cap on the display width of embedded images in points. Larger images are scaled down
    /// preserving aspect ratio; the default (468&nbsp;pt = 6.5&Prime;) fits a letter/A4 body.
    /// </summary>
    public double MaxDisplayWidthPt { get; init; } = 468.0;

    /// <summary>
    /// Path policy for local image sources passed to the shared <see cref="ImageAssetLoader"/>.
    /// Null uses <see cref="ImageSourceOptions.Default"/> (absolute paths rejected).
    /// </summary>
    public ImageSourceOptions? ImageSourceOptions { get; init; }

    /// <summary>
    /// Asset-safety limits passed to the shared <see cref="ImageAssetLoader"/>. Null uses
    /// <see cref="ImageAssetOptions.Default"/>.
    /// </summary>
    public ImageAssetOptions? ImageAssetOptions { get; init; }
}
