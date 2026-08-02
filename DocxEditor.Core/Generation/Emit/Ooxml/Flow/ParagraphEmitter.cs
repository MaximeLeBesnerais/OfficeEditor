using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Generation.Design;
using DocxEditor.Core.Generation.Emit.Ooxml.Design;
using DocxEditor.Core.Generation.Model;
using Model = DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Flow;

/// <summary>
/// Build options that tune paragraph construction for its context. Table cells and list items
/// tune these so cell text does not inherit the body-flow spacing defaults.
/// </summary>
internal sealed record ParagraphBuildOptions
{
    /// <summary>Baseline style kind used when the content has neither a role nor a heading level.</summary>
    public BaselineStyleKind DefaultStyleKind { get; init; } = BaselineStyleKind.Body;

    /// <summary>True when a plain (role-less, token-less) paragraph receives the theme body spacing defaults.</summary>
    public bool ApplyBodySpacingDefaults { get; init; } = true;
}

/// <summary>
/// Emits flow paragraphs and headings: paragraph properties (style ID, outline level for
/// headings, keep-next/keep-lines for headings, widow/orphan control for reading text,
/// alignment, density-aware spacing) plus runs with direct character formatting resolved from
/// semantic roles, typography tokens and individual run properties. Style IDs reference
/// existing styles only; semantic roles resolve to the active theme's defaults.
/// </summary>
internal static class ParagraphEmitter
{
    /// <summary>Emits a body paragraph with the given content/style/heading level.</summary>
    public static void EmitParagraph(
        OoxmlEmitContext context,
        OpenXmlCompositeElement container,
        TextModel content,
        string? styleId,
        int? headingLevel,
        string path,
        ParagraphBuildOptions? options = null)
    {
        container.Append(BuildParagraph(context, content, styleId, headingLevel, path, options));
    }

    /// <summary>Emits a heading paragraph.</summary>
    public static void EmitHeading(OoxmlEmitContext context, OpenXmlCompositeElement container, HeadingBlock heading, string path)
    {
        var content = heading.Content;
        container.Append(BuildParagraph(context, content, heading.Style, heading.Level, path));
    }

    /// <summary>
    /// Builds a paragraph for a text model. Default run formatting comes from the content's
    /// typography token, then the content's semantic role, then the heading role inferred from the
    /// heading level — direct run properties always win. Paragraph spacing follows the same
    /// precedence (explicit spacing wins over the role's density-scaled spacing).
    /// </summary>
    public static Paragraph BuildParagraph(
        OoxmlEmitContext context,
        TextModel content,
        string? styleId,
        int? headingLevel,
        string path,
        ParagraphBuildOptions? options = null)
    {
        var paragraph = new Paragraph();
        var paragraphProperties = new ParagraphProperties();
        var opts = options ?? new ParagraphBuildOptions();
        var design = context.DesignResolver.ResolveAll();

        var role = content.Role;
        var resolvedRole = role is { } explicitRole ? context.DesignResolver.TryResolveRole(explicitRole, path) : null;
        var headingRole = role is null && headingLevel is { } level ? (TextRole?)TextRoleExtensions.ForHeading(level) : null;
        var resolvedHeadingRole = headingRole is { } inferredHeadingRole ? context.DesignResolver.TryResolveRole(inferredHeadingRole, path) : null;
        var effectiveRole = resolvedRole ?? resolvedHeadingRole;
        var isHeading = headingLevel is not null || role?.IsHeading() == true;

        var fallbackStyle = role is { } styledRole
            ? styledRole.ToBaselineStyleKind()
            : headingLevel is { } headingStyleLevel
                ? BaselineStyleKindExtensions.ForHeading(headingStyleLevel)
                : opts.DefaultStyleKind;
        if (context.ResolveStyle(styleId, StyleValues.Paragraph, path, fallbackStyle) is { } resolvedStyleId)
        {
            paragraphProperties.ParagraphStyleId = new ParagraphStyleId { Val = resolvedStyleId };
        }

        var outlineLevel = role?.IsHeading() == true
            ? role.Value.HeadingLevel() - 1
            : headingLevel is { } outlineHeadingLevel ? outlineHeadingLevel - 1 : (int?)null;
        if (outlineLevel is { } outline)
        {
            paragraphProperties.OutlineLevel = new OutlineLevel { Val = outline };
        }

        if (effectiveRole?.KeepNext == true || isHeading)
        {
            paragraphProperties.KeepNext = new KeepNext();
        }
        if (effectiveRole?.KeepLines == true || isHeading)
        {
            paragraphProperties.KeepLines = new KeepLines();
        }
        if (!isHeading)
        {
            paragraphProperties.WidowControl = new WidowControl();
        }

        var alignment = content.Alignment ?? effectiveRole?.Paragraph.Alignment;
        if (alignment is { } resolvedAlignment && FormattingHelpers.Justification(resolvedAlignment) is { } justification)
        {
            paragraphProperties.Justification = new Justification { Val = justification };
        }

        ApplySpacing(context, paragraphProperties, content, effectiveRole, opts, path);

        if (paragraphProperties.HasChildren)
        {
            paragraph.ParagraphProperties = paragraphProperties;
        }

        var defaults = content.Token is not null
            ? FormattingHelpers.ResolveTypographyToken(context, content.Token, path)
            : effectiveRole is { } roleFormatting && !roleFormatting.Run.IsEmpty
                ? FormattingHelpers.FromRoleFormatting(roleFormatting)
                : null;

        if (content.Runs is { Count: > 0 } runs)
        {
            WarnUndersizedBodyText(context, content, runs, isHeading, path);
            foreach (var run in runs)
            {
                paragraph.Append(FormattingHelpers.BuildRun(context, run, defaults, path));
            }
        }
        else if (content.Text is { } text)
        {
            paragraph.Append(FormattingHelpers.BuildTextRun(context, text, defaults, path));
        }

        return paragraph;
    }

    /// <summary>
    /// Applies paragraph spacing: explicit content spacing wins; otherwise the role/heading role's
    /// density-scaled spacing; otherwise the theme body spacing defaults (unless the context opted
    /// out, e.g. table cells).
    /// </summary>
    private static void ApplySpacing(
        OoxmlEmitContext context,
        ParagraphProperties paragraphProperties,
        TextModel content,
        ResolvedRoleFormatting? effectiveRole,
        ParagraphBuildOptions opts,
        string path)
    {
        double? before = null;
        double? after = null;
        double? line = null;

        if (content.Spacing is { } spacing)
        {
            before = spacing.BeforePt;
            after = spacing.AfterPt;
            line = spacing.LineMultiple;
        }
        else if (effectiveRole is { } role && !role.Paragraph.IsEmpty)
        {
            before = role.Paragraph.SpaceBeforePt;
            after = role.Paragraph.SpaceAfterPt;
            line = role.Paragraph.LineSpacingMultiple;
        }
        else if (opts.ApplyBodySpacingDefaults)
        {
            var body = context.DesignResolver.ResolveAll().RoleOrDefault(TextRole.Body).Paragraph;
            before = body.SpaceBeforePt;
            after = body.SpaceAfterPt;
            line = body.LineSpacingMultiple;
        }

        if (before is null && after is null && line is null)
        {
            return;
        }

        var between = new SpacingBetweenLines();
        if (before is { } beforePt)
        {
            between.Before = FormattingHelpers.Twips(beforePt);
        }
        if (after is { } afterPt)
        {
            between.After = FormattingHelpers.Twips(afterPt);
        }
        if (line is { } lineMultiple)
        {
            between.Line = FormattingHelpers.LineMultiple(lineMultiple);
            between.LineRule = LineSpacingRuleValues.Auto;
        }
        paragraphProperties.SpacingBetweenLines = between;
    }

    /// <summary>
    /// Guardrail: warns once per reading paragraph when a run's explicit size is below the layout's
    /// minimum body size (default 9pt). Advisory only — no pagination prediction.
    /// </summary>
    private static void WarnUndersizedBodyText(
        OoxmlEmitContext context,
        TextModel content,
        IReadOnlyList<Model.Run> runs,
        bool isHeading,
        string path)
    {
        if (isHeading)
        {
            return;
        }
        if (content.Role is { } role && !role.IsReadingRole())
        {
            return;
        }

        var minSize = context.DesignResolver.ResolveAll().Layout.MinBodySizePt;
        foreach (var run in runs)
        {
            // Only direct run sizes are flagged — a typography token is a deliberate author choice.
            if (run.FontSizePt is { } size && size < minSize)
            {
                context.Warn(
                    path,
                    $"run text size {FormatSize(size)}pt is below the recommended minimum body size of {FormatSize(minSize)}pt; body text this small hurts readability.");
                return;
            }
        }
    }

    private static string FormatSize(double size) => size.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
}
