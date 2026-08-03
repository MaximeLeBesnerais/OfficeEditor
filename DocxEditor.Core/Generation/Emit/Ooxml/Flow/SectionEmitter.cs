using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Generation.Design;
using DocxEditor.Core.Generation.Model;
using DocxEditor.Core.Generation.Schema;
using OfficeEditor.Core.Exceptions;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Flow;

/// <summary>
/// Emits sections into the body: flow blocks in document order, optional per-section
/// headers/footers (fresh parts + relationships), the positioned tier (via the seam) and a
/// schema-valid <c>sectPr</c>. Intermediate sections end with a paragraph carrying the
/// section properties (attached to the section's last paragraph when possible); the final
/// section's properties become the body-level <c>sectPr</c>. CT_SectPr child order is
/// preserved: headerReference*, footerReference*, type, pgSz, pgMar, cols.
/// </summary>
internal static class SectionEmitter
{
    private const double DefaultHeaderFooterDistancePt = 36; // 0.5 inch, Word's default.

    public static void EmitAll(OoxmlEmitContext context, IReadOnlyList<Section> sections)
    {
        var body = GetBody(context);
        if (context.FromTemplate)
        {
            PreserveTemplateSections(body);
        }

        for (var i = 0; i < sections.Count; i++)
        {
            EmitSection(context, body, sections, i);
        }
    }

    /// <summary>Emits a body with a single default A4 section (no content).</summary>
    public static void EmitEmpty(OoxmlEmitContext context)
    {
        var body = GetBody(context);
        var empty = new Section { Blocks = [] };
        if (context.FromTemplate)
        {
            PreserveTemplateSections(body);
        }
        var path = "sections[0]";
        var resolvedPage = context.DesignResolver.ResolvePage(empty.PageSetup, path);
        var headerIds = EmitHeaderParts(context, empty, path, resolvedPage);
        var footerIds = EmitFooterParts(context, empty, path, resolvedPage);
        body.Append(BuildSectionProperties(context, empty, index: 0, sections: [empty], headerIds, footerIds, path, resolvedPage));
    }

    private static Body GetBody(OoxmlEmitContext context) =>
        context.MainPart.Document?.Body
        ?? throw new OfficeEditorException("The main document part has no <w:body>; cannot emit sections.");

    private static void EmitSection(OoxmlEmitContext context, Body body, IReadOnlyList<Section> sections, int index)
    {
        var section = sections[index];
        var path = $"sections[{index}]";
        var isLast = index == sections.Count - 1;
        var resolvedPage = context.DesignResolver.ResolvePage(section.PageSetup, path);

        var headerIds = EmitHeaderParts(context, section, path, resolvedPage);
        var footerIds = EmitFooterParts(context, section, path, resolvedPage);

        FlowBlockEmitter.EmitBlocks(context, body, section.Blocks, $"{path}.blocks", resolvedPage);

        EmitPositioned(context, section, body, path);

        var sectionProperties = BuildSectionProperties(context, section, index, sections, headerIds, footerIds, path, resolvedPage);

        if (isLast)
        {
            body.Append(sectionProperties);
        }
        else
        {
            AttachSectionBreak(body, sectionProperties);
        }
    }

    private static void EmitPositioned(OoxmlEmitContext context, Section section, Body body, string sectionPath)
    {
        if (section.Positioned.Count == 0)
        {
            return;
        }

        var positionedPath = $"$.{sectionPath}.positioned";
        var result = context.PositionedEmitter.Emit(context.MainPart, body, section.Positioned, positionedPath);
        foreach (var warning in result.Warnings)
        {
            var warningPath = positionedPath + $"[{warning.Index}]";
            if (!string.IsNullOrEmpty(warning.PathSuffix))
            {
                warningPath += $".{warning.PathSuffix}";
            }
            context.Warnings.Add(new DocxGenerationIssue(
                warningPath,
                $"[{warning.ElementType}] {warning.Message}",
                null,
                DocxGenerationIssueSeverity.Warning));
        }
    }

    private static IReadOnlyList<string> EmitHeaderParts(OoxmlEmitContext context, Section section, string path, ResolvedPageFormat resolvedPage)
    {
        var headerPart = context.MainPart.AddNewPart<HeaderPart>();
        var header = new Header();
        context.RegisterPartContainer(header, headerPart);
        if (section.Header.Count == 0)
        {
            header.Append(new Paragraph());
        }
        else
        {
            FlowBlockEmitter.EmitBlocks(context, header, section.Header, $"{path}.header", resolvedPage);
        }
        headerPart.Header = header;
        return [context.MainPart.GetIdOfPart(headerPart)];
    }

    private static IReadOnlyList<string> EmitFooterParts(OoxmlEmitContext context, Section section, string path, ResolvedPageFormat resolvedPage)
    {
        var footerPart = context.MainPart.AddNewPart<FooterPart>();
        var footer = new Footer();
        context.RegisterPartContainer(footer, footerPart);
        if (section.Footer.Count == 0)
        {
            footer.Append(new Paragraph());
        }
        else
        {
            FlowBlockEmitter.EmitBlocks(context, footer, section.Footer, $"{path}.footer", resolvedPage);
        }
        footerPart.Footer = footer;
        return [context.MainPart.GetIdOfPart(footerPart)];
    }

    private static SectionProperties BuildSectionProperties(
        OoxmlEmitContext context,
        Section section,
        int index,
        IReadOnlyList<Section> sections,
        IReadOnlyList<string> headerIds,
        IReadOnlyList<string> footerIds,
        string path,
        ResolvedPageFormat pageSetup)
    {
        // CT_SectPr child order: headerReference*, footerReference*, type, pgSz, pgMar, cols.
        var sectionProperties = new SectionProperties();

        foreach (var id in headerIds)
        {
            sectionProperties.Append(new HeaderReference { Type = HeaderFooterValues.Default, Id = id });
            sectionProperties.Append(new HeaderReference { Type = HeaderFooterValues.Even, Id = id });
            sectionProperties.Append(new HeaderReference { Type = HeaderFooterValues.First, Id = id });
        }
        foreach (var id in footerIds)
        {
            sectionProperties.Append(new FooterReference { Type = HeaderFooterValues.Default, Id = id });
            sectionProperties.Append(new FooterReference { Type = HeaderFooterValues.Even, Id = id });
            sectionProperties.Append(new FooterReference { Type = HeaderFooterValues.First, Id = id });
        }

        // The break that starts the FOLLOWING section is recorded on this section's sectPr.
        if (index < sections.Count - 1)
        {
            var breakType = sections[index + 1].PageSetup?.BreakType ?? SectionBreakType.NextPage;
            sectionProperties.Append(new SectionType { Val = SectionMarkValue(breakType) });
        }

        sectionProperties.Append(new DocumentFormat.OpenXml.Wordprocessing.PageSize
        {
            Width = (uint)FormattingHelpers.TwipsInt(pageSetup.WidthPt),
            Height = (uint)FormattingHelpers.TwipsInt(pageSetup.HeightPt),
            Orient = pageSetup.Orientation == PageOrientation.Landscape
                ? PageOrientationValues.Landscape
                : PageOrientationValues.Portrait
        });

        sectionProperties.Append(new PageMargin
        {
            Top = FormattingHelpers.TwipsInt(pageSetup.Margins.TopPt),
            Right = (uint)FormattingHelpers.TwipsInt(pageSetup.Margins.RightPt),
            Bottom = FormattingHelpers.TwipsInt(pageSetup.Margins.BottomPt),
            Left = (uint)FormattingHelpers.TwipsInt(pageSetup.Margins.LeftPt),
            Header = (uint)FormattingHelpers.TwipsInt(DefaultHeaderFooterDistancePt),
            Footer = (uint)FormattingHelpers.TwipsInt(DefaultHeaderFooterDistancePt)
        });

        if (pageSetup.Columns is { Count: > 1 } columns)
        {
            sectionProperties.Append(new Columns
            {
                EqualWidth = true,
                ColumnCount = (short)columns.Count,
                Space = FormattingHelpers.Twips(columns.SpacingPt),
                Separator = columns.SeparatorLine
            });
        }

        return sectionProperties;
    }

    /// <summary>
    /// Attaches an intermediate section's <c>sectPr</c> to the last paragraph of the section
    /// when possible (avoiding a trailing blank line); otherwise appends a dedicated section
    /// break paragraph.
    /// </summary>
    private static void AttachSectionBreak(Body body, SectionProperties sectionProperties)
    {
        if (body.Elements().LastOrDefault() is Paragraph lastParagraph
            && lastParagraph.ParagraphProperties?.SectionProperties is null)
        {
            lastParagraph.ParagraphProperties ??= new ParagraphProperties();
            lastParagraph.ParagraphProperties.SectionProperties = sectionProperties;
            return;
        }

        body.Append(new Paragraph(new ParagraphProperties(sectionProperties)));
    }

    private static void PreserveTemplateSections(Body body)
    {
        // A body-level sectPr describes the template's final section. Convert it to an
        // intermediate section break before appending generated content, otherwise the old
        // template body would inherit the first generated section's geometry and headers.
        foreach (var trailing in body.Elements<SectionProperties>().ToList())
        {
            trailing.Remove();
            body.Append(new Paragraph(new ParagraphProperties(trailing)));
        }
    }

    private static SectionMarkValues SectionMarkValue(SectionBreakType breakType) => breakType switch
    {
        SectionBreakType.Continuous => SectionMarkValues.Continuous,
        SectionBreakType.OddPage => SectionMarkValues.OddPage,
        SectionBreakType.EvenPage => SectionMarkValues.EvenPage,
        _ => SectionMarkValues.NextPage
    };

}
