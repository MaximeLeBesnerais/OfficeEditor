using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Generation.Assets;
using DocxEditor.Core.Generation.Emit.Ooxml.Flow;
using DocxEditor.Core.Generation.Emit.Ooxml.Images;
using DocxEditor.Core.Markdown.Model;
using DocxGenerationModel = DocxEditor.Core.Generation.Model;
using OfficeEditor.Core.Models;
using YamlDotNet.Core;
using YamlDotNet.RepresentationModel;
using A = DocumentFormat.OpenXml.Drawing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocxEditor.Core.Markdown.Rendering;

/// <summary>
/// Active run-level formatting accumulated while walking nested markdown inlines. Emphasis
/// combines recursively (e.g. bold inside italic) and is applied to the leaf text runs.
/// </summary>
[Flags]
internal enum MarkdownRunFlags
{
    None = 0,
    Italic = 1,
    Bold = 2,
    Strikethrough = 4,
    Subscript = 8,
    Superscript = 16,
    Inserted = 32,
    Marked = 64,
    Code = 128,
    Hyperlink = 256
}

/// <summary>
/// Renders the rich recursive markdown IR (see <see cref="RichMarkdownParser"/>) into DOCX
/// OOXML. Blocks render as paragraphs, tables, numbering-backed lists, a FootnotesPart and
/// core properties; inlines render as runs with nested formatting, real hyperlink
/// relationships and embedded local/data-URI images. All image loading flows through the
/// shared <see cref="ImageAssetLoader"/> and <see cref="DocxImagePartManager"/>; nothing is
/// ever fetched remotely, and an unresolved image renders as a visible alt fallback plus (in
/// strict mode) a diagnostic instead of vanishing.
/// </summary>
public sealed class RichMarkdownRenderer
{
    private readonly MainDocumentPart _mainPart;
    private readonly MarkdownRenderOptions _options;
    private readonly ImageAssetLoader _imageLoader;
    private readonly DocxImagePartManager _imagePartManager;
    private readonly Func<bool, int, int>? _allocateNumbering;
    private readonly NumberingAllocator? _fallbackNumbering;
    private MarkdownStyleResolver _styleResolver;
    private readonly HashSet<string> _appendedFallbackStyleIds = new(StringComparer.Ordinal);
    private readonly HashSet<string> _surfacedStyleDiagnostics = new(StringComparer.Ordinal);
    private readonly Dictionary<string, string?> _characterStyleIds = new(StringComparer.Ordinal);

    private readonly List<MarkdownDiagnostic> _diagnostics = [];
    private readonly List<OpenXmlElement> _rendered = [];

    private Body? _body;
    private uint _nextDrawingId;
    private Dictionary<string, int>? _footnoteIdsByLabel;

    private const string CorePropertiesNamespace =
        "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private const string DublinCoreNamespace = "http://purl.org/dc/elements/1.1/";

    /// <param name="allocateNumbering">
    /// Allocates a fresh, collision-free numbering instance id for one list (ordered, start).
    /// Each list gets its own instance so counters restart and lists never bind to existing
    /// numbering. When null, the renderer allocates through the shared
    /// <see cref="NumberingAllocator"/> against <paramref name="mainDocumentPart"/>.
    /// </param>
    public RichMarkdownRenderer(
        MainDocumentPart mainDocumentPart,
        MarkdownRenderOptions? options = null,
        Func<bool, int, int>? allocateNumbering = null)
    {
        _mainPart = mainDocumentPart ?? throw new ArgumentNullException(nameof(mainDocumentPart));
        _options = options ?? MarkdownRenderOptions.Default;
        _imageLoader = new ImageAssetLoader();
        _imagePartManager = new DocxImagePartManager(mainDocumentPart);
        _allocateNumbering = allocateNumbering;
        if (allocateNumbering is null)
        {
            _fallbackNumbering = new NumberingAllocator();
        }
        _styleResolver = _options.StyleResolver ?? new MarkdownStyleResolver(
            mainDocumentPart.StyleDefinitionsPart,
            _options.Strict ? MarkdownStyleResolverOptions.StrictMode : MarkdownStyleResolverOptions.Default);
    }

    /// <summary>Render diagnostics produced by the last <see cref="Render"/> call.</summary>
    public IReadOnlyList<MarkdownDiagnostic> Diagnostics => _diagnostics;

    /// <summary>The top-level body elements appended by the last <see cref="Render"/> call.</summary>
    public IReadOnlyList<OpenXmlElement> RenderedElements => _rendered;

    /// <summary>
    /// Renders a parsed markdown document into <paramref name="body"/> and returns the
    /// top-level elements appended (so a caller performing an in-place replace can relocate
    /// them). Relationships (hyperlinks, images, footnotes) are created directly against the
    /// document part, keeping every reference valid for callers that reopen or replace.
    /// </summary>
    public IReadOnlyList<OpenXmlElement> Render(Body body, MarkdownDocument document)
    {
        ArgumentNullException.ThrowIfNull(body);
        ArgumentNullException.ThrowIfNull(document);

        _body = body;
        _rendered.Clear();
        _diagnostics.Clear();
        _nextDrawingId = ComputeNextDrawingId(_mainPart);
        _footnoteIdsByLabel = BuildFootnoteIdMap(document);

        EnsureTrustedFallbackStyles(document);

        RenderBlocks(body, document.Blocks, depth: 0);
        return _rendered;
    }

    // ---- Trusted built-in styles (blank documents in strict mode) ----------

    /// <summary>
    /// Collects the semantic style keys (e.g. "heading1", "paragraph", "codeInline") the
    /// renderer actually resolves while rendering <paramref name="document"/>. Used to scope
    /// strict template style validation to the constructs an input really uses, and to seed
    /// trusted built-in fallback styles on blank documents. Must mirror the <c>ResolveStyle</c> /
    /// <c>ResolveCharacterStyle</c> call sites below.
    /// </summary>
    public static IReadOnlySet<string> CollectReferencedStyleKeys(MarkdownDocument document)
    {
        var keys = new HashSet<string>(StringComparer.Ordinal);
        CollectBlockStyleKeys(document.Blocks, ParagraphStyleContext.Default, keys);
        return keys;
    }

    private enum ParagraphStyleContext
    {
        Default,
        Quote,
        List,
        TableHeader,
        Footnote,
        DefinitionDescription
    }

    private static string ParagraphKeyFor(ParagraphStyleContext context) => context switch
    {
        ParagraphStyleContext.Quote => "blockquote",
        ParagraphStyleContext.List => "list",
        ParagraphStyleContext.TableHeader => "tableHeader",
        ParagraphStyleContext.Footnote => "footnoteText",
        ParagraphStyleContext.DefinitionDescription => "definitionDescription",
        _ => "paragraph"
    };

    private static void CollectBlockStyleKeys(IEnumerable<MarkdownBlock> blocks, ParagraphStyleContext context, ISet<string> keys)
    {
        foreach (var block in blocks)
        {
            switch (block)
            {
                case MarkdownHeading heading:
                    keys.Add($"heading{heading.Level}");
                    CollectInlineStyleKeys(heading.Inlines, keys);
                    break;
                case MarkdownParagraph paragraph:
                    keys.Add(ParagraphKeyFor(context));
                    CollectInlineStyleKeys(paragraph.Inlines, keys);
                    break;
                case MarkdownList list:
                    keys.Add("list");
                    foreach (var item in list.Items)
                    {
                        if (item.Blocks.Count > 0 && item.Blocks[0] is MarkdownParagraph first)
                        {
                            // The first block of a list item renders with the "list" key; the
                            // remaining blocks render under their own (default) context.
                            CollectInlineStyleKeys(first.Inlines, keys);
                            CollectBlockStyleKeys(item.Blocks.Skip(1), ParagraphStyleContext.Default, keys);
                        }
                        else
                        {
                            CollectBlockStyleKeys(item.Blocks, ParagraphStyleContext.Default, keys);
                        }
                    }
                    break;
                case MarkdownQuote quote:
                    keys.Add("blockquote");
                    CollectBlockStyleKeys(quote.Blocks, ParagraphStyleContext.Quote, keys);
                    break;
                case MarkdownCodeBlock:
                    keys.Add("codeBlock");
                    break;
                case MarkdownTable table:
                    keys.Add("table");
                    foreach (var row in table.Rows)
                    {
                        var cellContext = row.IsHeader ? ParagraphStyleContext.TableHeader : ParagraphStyleContext.Default;
                        if (row.IsHeader)
                        {
                            keys.Add("tableHeader");
                        }
                        foreach (var cell in row.Cells)
                        {
                            CollectBlockStyleKeys(cell.Blocks, cellContext, keys);
                        }
                    }
                    break;
                case MarkdownDefinitionList definitionList:
                    foreach (var item in definitionList.Items)
                    {
                        keys.Add("definitionTerm");
                        foreach (var definition in item.Definitions)
                        {
                            if (definition is MarkdownParagraph definitionParagraph)
                            {
                                keys.Add("definitionDescription");
                                CollectInlineStyleKeys(definitionParagraph.Inlines, keys);
                            }
                            else
                            {
                                CollectBlockStyleKeys([definition], ParagraphStyleContext.Default, keys);
                            }
                        }
                    }
                    break;
                case MarkdownFootnotesBlock footnotes:
                    keys.Add("footnoteText");
                    foreach (var footnote in footnotes.Footnotes)
                    {
                        CollectBlockStyleKeys(footnote.Blocks, ParagraphStyleContext.Footnote, keys);
                    }
                    break;
                case MarkdownHtmlBlock:
                case MarkdownUnknownBlock:
                    // Raw/unknown blocks render through RenderFallbackTextBlock, which resolves
                    // the plain "paragraph" style.
                    keys.Add("paragraph");
                    break;
                case MarkdownYamlFrontMatter:
                case MarkdownThematicBreak:
                case MarkdownLinkReferenceDefinitions:
                    break;
            }
        }
    }

    private static void CollectInlineStyleKeys(IEnumerable<MarkdownInline> inlines, ISet<string> keys)
    {
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case MarkdownCode:
                    keys.Add("codeInline");
                    break;
                case MarkdownLink link:
                    keys.Add("hyperlink");
                    CollectInlineStyleKeys(link.Children, keys);
                    break;
                case MarkdownEmphasis emphasis:
                    CollectInlineStyleKeys(emphasis.Children, keys);
                    break;
                case MarkdownImage image:
                    CollectInlineStyleKeys(image.Children, keys);
                    break;
            }
        }
    }

    /// <summary>
    /// In strict mode a freshly created (blank) document has no styles part, so every reference
    /// from the style mapping would be "unresolved" and the whole conversion would fail. Before
    /// strict resolution runs, the trusted built-in fallback styles for exactly the constructs
    /// the document uses are appended to a new styles part, making them real document styles the
    /// strict resolver finds. Only references that match the built-in default mapping are seeded:
    /// a user-supplied style map carries the author's explicit intent, so a reference it names
    /// must exist or strict mode reports it, even on a blank document. Existing template styles
    /// are never touched — this path only ever runs when the document has no styles part at all.
    /// </summary>
    private void EnsureTrustedFallbackStyles(MarkdownDocument document)
    {
        if (!_options.Strict)
        {
            return;
        }
        if (_options.StyleResolver is not null)
        {
            // A caller-injected resolver owns its style view; do not replace it.
            return;
        }
        if (_mainPart.StyleDefinitionsPart is not null)
        {
            return;
        }

        var styles = new Styles();
        foreach (var key in CollectReferencedStyleKeys(document))
        {
            var reference = _options.StyleMapping?.GetStyle(key);
            if (string.IsNullOrWhiteSpace(reference))
            {
                continue;
            }

            // Only the built-in default references are trusted on a blank document. Custom
            // mappings must resolve against a real styles part (or fail in strict mode).
            var defaultValue = StyleMapping.Default.GetStyle(key);
            if (!string.Equals(reference, defaultValue, StringComparison.Ordinal))
            {
                continue;
            }

            if (styles.Elements<Style>().Any(s => s.StyleId?.Value == reference))
            {
                continue;
            }

            styles.Append(MarkdownFallbackStyles.Create(
                reference, reference, MarkdownStyleKinds.ForElement(key), key));
        }

        if (styles.ChildElements.Count == 0)
        {
            return;
        }

        var stylesPart = _mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = styles;
        _styleResolver = new MarkdownStyleResolver(stylesPart, MarkdownStyleResolverOptions.StrictMode);
    }

    // ---- Style resolution -------------------------------------------------

    /// <summary>
    /// Resolves the style reference for a paragraph/table semantic key against the document
    /// styles and returns the style id to emit (or null when no style applies). Diagnostics
    /// from the resolver surface into <see cref="Diagnostics"/>, and a generated fallback of
    /// the expected kind is appended to the document styles part exactly once.
    /// </summary>
    private string? ResolveStyle(string key)
    {
        var reference = _options.StyleMapping?.GetStyle(key);
        if (reference is null)
        {
            return null;
        }
        return ApplyResolution(_styleResolver.ResolveElement(key, reference));
    }

    /// <summary>
    /// Resolves a character style (inline code, hyperlink) once per key and caches the result,
    /// applying the same diagnostic/fallback plumbing as <see cref="ResolveStyle"/>.
    /// </summary>
    private string? ResolveCharacterStyle(string key)
    {
        if (_characterStyleIds.TryGetValue(key, out var cached))
        {
            return cached;
        }

        string? id = null;
        var reference = _options.StyleMapping?.GetStyle(key);
        if (reference is not null)
        {
            id = ApplyResolution(_styleResolver.ResolveElement(key, reference));
        }
        _characterStyleIds[key] = id;
        return id;
    }

    private string? ApplyResolution(MarkdownStyleResolution resolution)
    {
        foreach (var diagnostic in resolution.Diagnostics)
        {
            // A single unresolved reference is reported once even when it applies to many
            // elements (every paragraph would otherwise re-emit the same fallback warning).
            if (_surfacedStyleDiagnostics.Add(diagnostic.Message))
            {
                _diagnostics.Add(diagnostic.ToMarkdownDiagnostic());
            }
        }

        if (resolution.FallbackStyle is { } fallback
            && resolution.StyleId is { } id
            && _appendedFallbackStyleIds.Add(id))
        {
            AppendFallbackStyle(fallback);
        }

        return resolution.StyleId;
    }

    /// <summary>
    /// Appends a generated fallback style to the document styles part so every style
    /// referenced in the rendered output actually exists. Existing template styles are
    /// never modified — only this new, resolver-owned style is added.
    /// </summary>
    private void AppendFallbackStyle(Style style)
    {
        var stylesPart = _mainPart.StyleDefinitionsPart;
        if (stylesPart is null)
        {
            stylesPart = _mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles();
        }

        stylesPart.Styles ??= new Styles();
        stylesPart.Styles.Append(style);
    }

    // ---- Blocks ----------------------------------------------------------

    private void RenderBlocks(OpenXmlCompositeElement container, IEnumerable<MarkdownBlock> blocks, int depth)
    {
        foreach (var block in blocks)
        {
            RenderBlock(container, block, depth);
        }
    }

    private void RenderBlock(OpenXmlCompositeElement container, MarkdownBlock block, int depth)
    {
        if (depth > RichMarkdownParser.MaxNestingDepth)
        {
            // The parser already flattens beyond the limit; this guard is defense in depth for
            // hand-constructed IR so rendering can never recurse past the process stack.
            FlattenBlock(container);
            return;
        }

        switch (block)
        {
            case MarkdownHeading heading:
                RenderHeading(container, heading, depth);
                break;
            case MarkdownParagraph paragraph:
                RenderParagraph(container, paragraph, depth);
                break;
            case MarkdownList list:
                RenderList(container, list, depth);
                break;
            case MarkdownQuote quote:
                RenderQuote(container, quote, depth);
                break;
            case MarkdownCodeBlock code:
                RenderCodeBlock(container, code, depth);
                break;
            case MarkdownTable table:
                RenderTable(container, table, depth);
                break;
            case MarkdownThematicBreak:
                RenderThematicBreak(container);
                break;
            case MarkdownDefinitionList definitionList:
                RenderDefinitionList(container, definitionList, depth);
                break;
            case MarkdownFootnotesBlock footnotes:
                RenderFootnotes(footnotes);
                break;
            case MarkdownYamlFrontMatter yaml:
                RenderYamlFrontMatter(yaml);
                break;
            case MarkdownHtmlBlock html:
                RenderHtmlBlock(container, html);
                break;
            case MarkdownUnknownBlock unknown:
                RenderUnknownBlock(container, unknown);
                break;
            case MarkdownLinkReferenceDefinitions:
                // Reference definitions are resolved into links by the parser; nothing renders.
                break;
        }
    }

    private void RenderHeading(OpenXmlCompositeElement container, MarkdownHeading heading, int depth)
    {
        var paragraph = new Paragraph();
        var style = ResolveStyle($"heading{heading.Level}");
        var pPr = new ParagraphProperties();
        if (!string.IsNullOrEmpty(style))
        {
            pPr.Append(new ParagraphStyleId { Val = style });
        }
        if (depth > 0)
        {
            pPr.Append(new Indentation { Left = (720 * depth).ToString() });
        }
        if (pPr.ChildElements.Count > 0)
        {
            paragraph.ParagraphProperties = pPr;
        }

        RenderInlines(paragraph, heading.Inlines, MarkdownRunFlags.None, depth: depth);
        AppendBlock(container, paragraph);
    }

    private void RenderParagraph(
        OpenXmlCompositeElement container,
        MarkdownParagraph paragraph,
        int depth,
        bool quote = false,
        int? numberingId = null,
        JustificationValues? alignment = null,
        MarkdownRunFlags forced = MarkdownRunFlags.None,
        string? styleKey = null)
    {
        var p = new Paragraph();
        var pPr = new ParagraphProperties();

        var style = ResolveStyle(styleKey ?? (quote ? "blockquote" : "paragraph"));
        if (!string.IsNullOrEmpty(style))
        {
            pPr.Append(new ParagraphStyleId { Val = style });
        }

        // Widow/orphan control keeps the last line of a multi-line paragraph together with the
        // previous line at a page break, instead of stranding a single trailing word on the
        // next page. Word's own default, applied explicitly so it holds regardless of the
        // document's Normal/body style definition.
        pPr.Append(new WidowControl());

        if (numberingId is { } numId)
        {
            pPr.Append(new NumberingProperties(
                new NumberingLevelReference { Val = 0 },
                new NumberingId { Val = numId }));
            pPr.Append(NumberedIndentation(depth));
        }
        else if (quote)
        {
            pPr.Append(QuoteIndentation(depth));
        }
        else if (depth > 0)
        {
            pPr.Append(ContinuationIndentation(depth));
        }

        if (alignment is { } jc)
        {
            pPr.Append(new Justification { Val = jc });
        }

        if (pPr.ChildElements.Count > 0)
        {
            p.ParagraphProperties = pPr;
        }

        RenderInlines(p, paragraph.Inlines, MarkdownRunFlags.None, forced, depth);
        AppendBlock(container, p);
    }

    private void RenderList(OpenXmlCompositeElement container, MarkdownList list, int depth)
    {
        var start = ParseStart(list.OrderedStart);
        var numberingId = AllocateNumbering(list.Ordered, start);

        foreach (var item in list.Items)
        {
            RenderListItem(container, item, depth, numberingId);
        }
    }

    private void RenderListItem(OpenXmlCompositeElement container, MarkdownListItem item, int depth, int numberingId)
    {
        var blocks = item.Blocks;
        if (blocks.Count == 0)
        {
            // An empty list item still carries its bullet/number marker.
            RenderParagraph(container, new MarkdownParagraph(), depth, numberingId: numberingId, styleKey: "list");
            return;
        }

        if (blocks[0] is MarkdownParagraph first)
        {
            RenderParagraph(container, first, depth, numberingId: numberingId, styleKey: "list");
            for (var i = 1; i < blocks.Count; i++)
            {
                RenderBlock(container, blocks[i], depth + 1);
            }
        }
        else
        {
            // A list item whose first block is not a paragraph (fenced code, a nested list,
            // a table, …) would otherwise lose its bullet/number marker entirely. Emit an
            // empty numbered paragraph first so the visible marker survives, then render the
            // item's blocks.
            RenderParagraph(container, new MarkdownParagraph(), depth, numberingId: numberingId, styleKey: "list");
            foreach (var block in blocks)
            {
                RenderBlock(container, block, depth + 1);
            }
        }
    }

    private void RenderQuote(OpenXmlCompositeElement container, MarkdownQuote quote, int depth)
    {
        if (depth > RichMarkdownParser.MaxNestingDepth)
        {
            // Nested quotes recurse through RenderQuote (bypassing the RenderBlock guard), so the
            // depth bound is enforced here as well; see FlattenBlock.
            FlattenBlock(container);
            return;
        }

        foreach (var block in quote.Blocks)
        {
            RenderQuoteInner(container, block, depth);
        }
    }

    private void RenderQuoteInner(OpenXmlCompositeElement container, MarkdownBlock block, int depth)
    {
        switch (block)
        {
            case MarkdownParagraph paragraph:
                RenderParagraph(container, paragraph, depth, quote: true);
                break;
            case MarkdownQuote nested:
                RenderQuote(container, nested, depth + 1);
                break;
            case MarkdownHeading heading:
                RenderHeading(container, heading, depth);
                break;
            case MarkdownList list:
                RenderList(container, list, depth + 1);
                break;
            case MarkdownCodeBlock code:
                RenderCodeBlock(container, code, depth);
                break;
            case MarkdownThematicBreak:
                RenderThematicBreak(container);
                break;
            default:
                RenderBlock(container, block, depth + 1);
                break;
        }
    }

    private void RenderCodeBlock(OpenXmlCompositeElement container, MarkdownCodeBlock code, int depth)
    {
        var style = ResolveStyle("codeBlock");
        foreach (var line in SplitLines(code.Text))
        {
            var paragraph = new Paragraph();
            var pPr = new ParagraphProperties();
            if (!string.IsNullOrEmpty(style))
            {
                pPr.Append(new ParagraphStyleId { Val = style });
            }
            if (depth > 0)
            {
                pPr.Append(new Indentation { Left = (720 * depth).ToString() });
            }
            if (pPr.ChildElements.Count > 0)
            {
                paragraph.ParagraphProperties = pPr;
            }
            paragraph.Append(new Run(new Text(line) { Space = SpaceProcessingModeValues.Preserve }));
            AppendBlock(container, paragraph);
        }
    }

    private void RenderTable(OpenXmlCompositeElement container, MarkdownTable table, int depth)
    {
        var columnCount = table.Columns.Count;
        if (columnCount == 0 && table.Rows.Count > 0)
        {
            columnCount = table.Rows.Max(r => r.Cells.Count);
        }
        if (columnCount == 0)
        {
            // An empty table renders nothing: CT_Tbl requires a tblGrid and rows after tblPr.
            return;
        }

        var tbl = new Table();
        var tableStyle = ResolveStyle("table");
        var tblPr = new TableProperties();
        if (!string.IsNullOrEmpty(tableStyle))
        {
            // CT_TblPr requires tblStyle to precede tblBorders.
            tblPr.Append(new TableStyle { Val = tableStyle });
        }
        tblPr.Append(new TableBorders(
            new TopBorder { Val = BorderValues.Single, Size = 4 },
            new LeftBorder { Val = BorderValues.Single, Size = 4 },
            new BottomBorder { Val = BorderValues.Single, Size = 4 },
            new RightBorder { Val = BorderValues.Single, Size = 4 },
            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
            new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }));
        tbl.Append(tblPr);

        var grid = new TableGrid();
        for (var i = 0; i < columnCount; i++)
        {
            grid.Append(new GridColumn());
        }
        tbl.Append(grid);

        foreach (var row in table.Rows)
        {
            var tableRow = new TableRow();

            // Keep the row intact across a page boundary (no empty continuation rows) and
            // repeat the header row at the top of every page the table spans. Per CT_TrPr
            // ordering w:cantSplit must precede w:tblHeader.
            var rowProperties = new TableRowProperties(new CantSplit());
            if (row.IsHeader)
            {
                rowProperties.Append(new TableHeader());
            }
            tableRow.TableRowProperties = rowProperties;

            var cellIndex = 0;
            foreach (var cell in row.Cells)
            {
                if (cell.ColumnSpan > 1 || cell.RowSpan > 1)
                {
                    Warn($"Table cell spanning {cell.ColumnSpan}×{cell.RowSpan} rendered without row/column merge.");
                }

                var tableCell = new TableCell();
                if (row.IsHeader)
                {
                    tableCell.TableCellProperties = new TableCellProperties(
                        new Shading { Val = ShadingPatternValues.Clear, Fill = "D9E2F3" });
                }

                // Pipe-table cells do not always carry a column index; fall back to the
                // ordinal position within the row so alignment still resolves.
                var effectiveColumnIndex = cell.ColumnIndex >= 0 ? cell.ColumnIndex : cellIndex;
                var alignment = GetColumnAlignment(table, effectiveColumnIndex);
                var forced = row.IsHeader ? MarkdownRunFlags.Bold : MarkdownRunFlags.None;
                foreach (var block in cell.Blocks)
                {
                    if (block is MarkdownParagraph paragraph)
                    {
                        RenderParagraph(
                            tableCell,
                            paragraph,
                            depth: 0,
                            alignment: alignment,
                            forced: forced,
                            styleKey: row.IsHeader ? "tableHeader" : null);
                    }
                    else
                    {
                        RenderBlock(tableCell, block, depth);
                    }
                }

                if (tableCell.ChildElements.Count == 0)
                {
                    tableCell.Append(new Paragraph());
                }
                tableRow.Append(tableCell);
                cellIndex++;
            }
            tbl.Append(tableRow);
        }

        AppendBlock(container, tbl);
    }

    private void RenderThematicBreak(OpenXmlCompositeElement container)
    {
        AppendBlock(container, new Paragraph(
            new ParagraphProperties(
                new ParagraphBorders(
                    new BottomBorder
                    {
                        Val = BorderValues.Single,
                        Size = 6,
                        Space = 1
                    }))));
    }

    private void RenderDefinitionList(OpenXmlCompositeElement container, MarkdownDefinitionList definitionList, int depth)
    {
        foreach (var item in definitionList.Items)
        {
            foreach (var term in item.Terms)
            {
                var paragraph = new Paragraph();
                var termStyle = ResolveStyle("definitionTerm");
                if (!string.IsNullOrEmpty(termStyle))
                {
                    paragraph.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = termStyle });
                }
                RenderInlines(paragraph, term.Inlines, MarkdownRunFlags.Bold, depth: depth);
                AppendBlock(container, paragraph);
            }

            foreach (var definition in item.Definitions)
            {
                if (definition is MarkdownParagraph paragraph)
                {
                    RenderParagraph(container, paragraph, depth + 1, styleKey: "definitionDescription");
                }
                else
                {
                    RenderBlock(container, definition, depth + 1);
                }
            }
        }
    }

    private void RenderHtmlBlock(OpenXmlCompositeElement container, MarkdownHtmlBlock html)
    {
        Warn($"Raw HTML block rendered as escaped visible text (type '{html.Type}').");
        RenderFallbackTextBlock(container, html.Text);
    }

    private void RenderUnknownBlock(OpenXmlCompositeElement container, MarkdownUnknownBlock unknown)
    {
        Warn($"Unrecognized markdown block '{unknown.Kind}' rendered as escaped visible text.");
        RenderFallbackTextBlock(container, unknown.Raw);
    }

    private void RenderFallbackTextBlock(OpenXmlCompositeElement container, string text)
    {
        var paragraph = new Paragraph();
        var style = ResolveStyle("paragraph");
        if (!string.IsNullOrEmpty(style))
        {
            paragraph.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = style });
        }
        paragraph.Append(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        AppendBlock(container, paragraph);
    }

    /// <summary>
    /// Renders a block whose nesting exceeded <see cref="RichMarkdownParser.MaxNestingDepth"/> as
    /// a single visible marker plus a diagnostic instead of recursing. Extracting the block's
    /// text would itself recurse over its (deep) children, so the marker is deliberately flat.
    /// </summary>
    private void FlattenBlock(OpenXmlCompositeElement container)
    {
        _diagnostics.Add(new MarkdownDiagnostic(
            MarkdownDiagnosticSeverity.Warning,
            $"Markdown nesting exceeds the maximum supported depth of {RichMarkdownParser.MaxNestingDepth}; block content omitted."));
        RenderFallbackTextBlock(container, $"[omitted: markdown nesting exceeds the maximum supported depth of {RichMarkdownParser.MaxNestingDepth}]");
    }

    private void RenderYamlFrontMatter(MarkdownYamlFrontMatter yaml)
    {
        if (!_options.MapYamlFrontMatterToCoreProperties)
        {
            return;
        }

        SetCorePropertiesFromYaml(yaml.Yaml);
    }

    // ---- Footnotes and core properties -----------------------------------

    private void RenderFootnotes(MarkdownFootnotesBlock block)
    {
        var footnotesPart = _mainPart.FootnotesPart ?? _mainPart.AddNewPart<FootnotesPart>();
        var footnotes = footnotesPart.Footnotes ?? new Footnotes();

        if (!footnotes.Elements<Footnote>().Any(f => f.Type?.Value == FootnoteEndnoteValues.Separator))
        {
            footnotes.Append(CreateSpecialFootnote(0, FootnoteEndnoteValues.Separator));
        }
        if (!footnotes.Elements<Footnote>().Any(f => f.Type?.Value == FootnoteEndnoteValues.ContinuationSeparator))
        {
            footnotes.Append(CreateSpecialFootnote(1, FootnoteEndnoteValues.ContinuationSeparator));
        }

        foreach (var footnote in block.Footnotes)
        {
            if (!_footnoteIdsByLabel!.TryGetValue(footnote.Label, out var id))
            {
                continue;
            }

            var footnoteElement = new Footnote { Id = id };
            foreach (var footnoteBlock in footnote.Blocks)
            {
                if (footnoteBlock is MarkdownParagraph paragraph)
                {
                    RenderParagraph(footnoteElement, paragraph, depth: 0, styleKey: "footnoteText");
                }
                else
                {
                    RenderBlock(footnoteElement, footnoteBlock, depth: 0);
                }
            }
            if (footnoteElement.ChildElements.Count == 0)
            {
                footnoteElement.Append(new Paragraph());
            }
            else
            {
                PrependFootnoteMark(footnoteElement);
            }
            footnotes.Append(footnoteElement);
        }

        footnotesPart.Footnotes = footnotes;
    }

    private static Footnote CreateSpecialFootnote(int id, FootnoteEndnoteValues type)
    {
        var paragraph = new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }),
            new Run(type == FootnoteEndnoteValues.Separator ? new SeparatorMark() : new ContinuationSeparatorMark()));
        return new Footnote(paragraph) { Id = id, Type = type };
    }

    /// <summary>
    /// Inserts the footnote number mark as the first run of the first paragraph, followed by a
    /// space so the note text does not run flush against the reference number (Word's convention).
    /// </summary>
    private static void PrependFootnoteMark(Footnote footnote)
    {
        var first = footnote.Elements<Paragraph>().FirstOrDefault();
        if (first is null)
        {
            return;
        }
        var mark = new Run(new FootnoteReferenceMark());
        var spacer = new Run(new Text(" ") { Space = SpaceProcessingModeValues.Preserve });
        if (first.ParagraphProperties is not null)
        {
            first.InsertAfter(mark, first.ParagraphProperties);
            first.InsertAfter(spacer, mark);
        }
        else
        {
            first.InsertAt(mark, 0);
            first.InsertAt(spacer, 1);
        }
    }

    private Dictionary<string, int> BuildFootnoteIdMap(MarkdownDocument document)
    {
        var map = new Dictionary<string, int>(StringComparer.Ordinal);
        var footnotesBlock = document.Blocks.OfType<MarkdownFootnotesBlock>().FirstOrDefault();
        if (footnotesBlock is null)
        {
            return map;
        }

        var nextId = ComputeNextFootnoteId();
        foreach (var footnote in footnotesBlock.Footnotes)
        {
            if (string.IsNullOrEmpty(footnote.Label))
            {
                continue;
            }
            map[footnote.Label] = nextId++;
        }
        return map;
    }

    private int ComputeNextFootnoteId()
    {
        if (_mainPart.FootnotesPart?.Footnotes is not { } footnotes)
        {
            return 2;
        }
        var maxId = footnotes.Elements<Footnote>()
            .Select(f => (int)(f.Id?.Value ?? 0))
            .DefaultIfEmpty(0)
            .Max();
        return Math.Max(maxId + 1, 2);
    }

    private void SetCorePropertiesFromYaml(string yaml)
    {
        string? title = null, author = null, subject = null, keywords = null, description = null, language = null;
        try
        {
            var stream = new YamlStream();
            using (var reader = new StringReader(yaml))
            {
                stream.Load(reader);
            }
            if (stream.Documents.Count == 0 || stream.Documents[0].RootNode is not YamlMappingNode yamlRoot)
            {
                return;
            }

            string? GetString(string key) =>
                yamlRoot.Children.TryGetValue(new YamlScalarNode(key), out var node) && node is YamlScalarNode scalar
                    ? scalar.Value
                    : null;

            title = GetString("title");
            author = GetString("author");
            subject = GetString("subject");
            keywords = GetString("keywords");
            description = GetString("description");
            language = GetString("language");
        }
        catch (YamlException)
        {
            Warn("YAML front matter could not be parsed; core properties were not set.");
            return;
        }

        if (title is null && author is null && subject is null
            && keywords is null && description is null && language is null)
        {
            return;
        }

        if (_mainPart.OpenXmlPackage is not WordprocessingDocument package)
        {
            throw new InvalidOperationException(
                "Rich markdown core-property mapping requires a WordprocessingDocument package.");
        }
        var corePart = package.CoreFilePropertiesPart ?? package.AddCoreFilePropertiesPart();

        XDocument document;
        try
        {
            using var readStream = corePart.GetStream(FileMode.OpenOrCreate, FileAccess.Read);
            document = XDocument.Load(readStream);
        }
        catch
        {
            document = new XDocument(new XElement(XNamespace.Get(CorePropertiesNamespace) + "coreProperties"));
        }

        var root = document.Root;
        if (root is null)
        {
            return;
        }

        void Set(string localName, string nsUri, string? value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return;
            }
            var ns = XNamespace.Get(nsUri);
            var element = root.Elements().FirstOrDefault(e => e.Name == ns + localName);
            if (element is null)
            {
                element = new XElement(ns + localName);
                root.Add(element);
            }
            element.Value = value;
        }

        Set("title", DublinCoreNamespace, title);
        Set("creator", DublinCoreNamespace, author);
        Set("subject", DublinCoreNamespace, subject);
        Set("keywords", CorePropertiesNamespace, keywords);
        Set("description", DublinCoreNamespace, description);
        Set("language", DublinCoreNamespace, language);

        using var writeStream = corePart.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(writeStream, new System.Text.UTF8Encoding(false));
        document.Save(writer);
    }

    // ---- Inlines ---------------------------------------------------------

    private void RenderInlines(
        OpenXmlCompositeElement container,
        IEnumerable<MarkdownInline> inlines,
        MarkdownRunFlags flags,
        MarkdownRunFlags forced = MarkdownRunFlags.None,
        int depth = 0)
    {
        foreach (var inline in inlines)
        {
            RenderInline(container, inline, flags | forced, depth);
        }
    }

    private void RenderInline(OpenXmlCompositeElement container, MarkdownInline inline, MarkdownRunFlags flags, int depth)
    {
        if (depth > RichMarkdownParser.MaxNestingDepth)
        {
            // Inline nesting (emphasis inside emphasis, links, images) is bounded by the same
            // depth cap so adversarial or hand-constructed IR can never overflow the stack; the
            // content degrades to plain text instead of recursing further.
            _diagnostics.Add(new MarkdownDiagnostic(
                MarkdownDiagnosticSeverity.Warning,
                $"Markdown nesting exceeds the maximum supported depth of {RichMarkdownParser.MaxNestingDepth}; inline content rendered as plain text."));
            AppendTextRun(container, PlainText([inline]), flags);
            return;
        }

        switch (inline)
        {
            case MarkdownText text:
                AppendTextRun(container, text.Text, flags);
                break;
            case MarkdownEntity entity:
                AppendTextRun(container, entity.Decoded, flags);
                break;
            case MarkdownEmoji emoji:
                AppendTextRun(container, emoji.Text, flags);
                break;
            case MarkdownEmphasis emphasis:
                RenderInlines(container, emphasis.Children, flags | FlagFor(emphasis.Kind), depth: depth + 1);
                break;
            case MarkdownCode code:
                AppendCodeRun(container, code.Content, flags);
                break;
            case MarkdownLink link:
                RenderLink(container, link, flags, depth);
                break;
            case MarkdownImage image:
                RenderImage(container, image, flags);
                break;
            case MarkdownLineBreak lineBreak:
                RenderLineBreak(container, lineBreak);
                break;
            case MarkdownFootnoteReference reference:
                RenderFootnoteReference(container, reference);
                break;
            case MarkdownTaskCheckbox checkbox:
                AppendTextRun(container, checkbox.Checked ? "\u2611" : "\u2610", flags);
                break;
            case MarkdownHtml html:
                if (IsEmptyAnchorTag(html.Tag))
                {
                    // Markdig's AutoIdentifiers injects an empty <a id="…"></a> anchor into
                    // headings; it carries no visible text, so it renders nothing.
                    break;
                }
                Warn("Raw inline HTML rendered as escaped visible text.");
                AppendTextRun(container, html.Tag, flags);
                break;
            case MarkdownUnknownInline unknown:
                Warn($"Unrecognized inline '{unknown.Kind}' rendered as escaped visible text.");
                AppendTextRun(container, unknown.Raw, flags);
                break;
        }
    }

    private void RenderLink(OpenXmlCompositeElement container, MarkdownLink link, MarkdownRunFlags flags, int depth)
    {
        var url = link.Url;
        if (string.IsNullOrWhiteSpace(url))
        {
            RenderInlines(container, link.Children, flags, depth: depth + 1);
            return;
        }

        if (IsInternalAnchor(url))
        {
            // Internal anchors are not supported (no bookmark emission); the label renders
            // as plain text with a diagnostic instead of a dangling external relationship.
            RenderLinkAsPlainText(
                container, link, flags, depth,
                $"Link URL '{url}' is an internal anchor, which is not supported; rendered as plain text.");
            return;
        }

        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri))
        {
            RenderLinkAsPlainText(
                container, link, flags, depth,
                $"Link URL '{url}' is not absolute; rendered as plain text.");
            return;
        }

        if (!IsSafeExternalScheme(uri.Scheme))
        {
            RenderLinkAsPlainText(
                container, link, flags, depth,
                $"Link URL '{url}' uses scheme '{uri.Scheme}', which is not an allowed external link scheme (http, https, mailto); rendered as plain text.");
            return;
        }

        var relationship = _mainPart.AddHyperlinkRelationship(uri, isExternal: true);
        var hyperlink = new Hyperlink { Id = relationship.Id, History = true };
        RenderInlines(hyperlink, link.Children, flags | MarkdownRunFlags.Hyperlink, depth: depth + 1);
        container.Append(hyperlink);
    }

    private void RenderLinkAsPlainText(
        OpenXmlCompositeElement container,
        MarkdownLink link,
        MarkdownRunFlags flags,
        int depth,
        string message)
    {
        _diagnostics.Add(new MarkdownDiagnostic(MarkdownDiagnosticSeverity.Warning, message));
        RenderInlines(container, link.Children, flags, depth: depth + 1);
    }

    private static bool IsInternalAnchor(string url) =>
        url.StartsWith("#", StringComparison.Ordinal);

    private static bool IsSafeExternalScheme(string scheme) =>
        scheme.Equals("http", StringComparison.OrdinalIgnoreCase)
        || scheme.Equals("https", StringComparison.OrdinalIgnoreCase)
        || scheme.Equals("mailto", StringComparison.OrdinalIgnoreCase);

    private void RenderImage(OpenXmlCompositeElement container, MarkdownImage image, MarkdownRunFlags flags)
    {
        var alt = PlainText(image.Children);
        var source = image.Url;

        if (string.IsNullOrWhiteSpace(source))
        {
            RenderImageFallback(container, alt);
            return;
        }

        ImageAsset asset;
        try
        {
            asset = _imageLoader.Load(source, _options.ImageSourceOptions, _options.ImageAssetOptions);
        }
        catch (Exception ex) when (ex is ImageSourceException or IOException or UnauthorizedAccessException)
        {
            RenderImageFallback(container, alt, ex.Message);
            return;
        }

        var registered = _imagePartManager.Register(asset);
        var (widthPt, heightPt) = FitToMaxWidth(asset, _options.MaxDisplayWidthPt);
        var geometry = ResolvedImageGeometry.Resolve(
            asset, DocxGenerationModel.ImageFitMode.Contain, null, widthPt, heightPt);
        var docPrId = _nextDrawingId++;
        var pictureNonVisualId = _nextDrawingId++;
        var drawing = InlinePictureXmlFactory.BuildInline(registered, geometry, docPrId, pictureNonVisualId, alt);
        container.Append(new Run(drawing));
    }

    private void RenderImageFallback(OpenXmlCompositeElement container, string alt, string? reason = null)
    {
        var text = string.IsNullOrEmpty(alt) ? "[image]" : $"[image: {alt}]";
        Warn(reason is null
            ? $"Image could not be embedded; rendered as visible fallback '{text}'."
            : $"Image could not be embedded ({reason}); rendered as visible fallback '{text}'.");

        var run = new Run();
        var rPr = new RunProperties();
        rPr.Append(new Italic());
        rPr.Append(new Color { Val = "808080" });
        run.Append(rPr);
        run.Append(new Text(text));
        container.Append(run);
    }

    private void RenderFootnoteReference(OpenXmlCompositeElement container, MarkdownFootnoteReference reference)
    {
        if (reference.IsBackLink)
        {
            // The markdown backlink marker inside a footnote definition is navigation scaffolding,
            // not footnote content. Word's own backlink is the footnote reference mark rendered at
            // the start of the footnote body; emitting the literal caret would show a spurious
            // visible "^" artifact, so it is omitted.
            return;
        }

        if (_footnoteIdsByLabel is { } map && map.TryGetValue(reference.Label, out var id))
        {
            container.Append(new Run(new FootnoteReference { Id = id }));
            return;
        }

        Warn($"Footnote reference '{reference.Label}' has no matching definition; rendered as literal text.");
        AppendTextRun(container, $"[{reference.Label}]", MarkdownRunFlags.None);
    }

    private void RenderLineBreak(OpenXmlCompositeElement container, MarkdownLineBreak lineBreak)
    {
        if (lineBreak.IsHard)
        {
            container.Append(new Run(new Break()));
            return;
        }

        switch (_options.SoftBreakMode)
        {
            case MarkdownSoftBreakMode.Space:
                AppendTextRun(container, " ", MarkdownRunFlags.None);
                break;
            case MarkdownSoftBreakMode.LineBreak:
                container.Append(new Run(new Break()));
                break;
            case MarkdownSoftBreakMode.None:
                break;
        }
    }

    private void AppendTextRun(OpenXmlCompositeElement container, string text, MarkdownRunFlags flags)
    {
        if (string.IsNullOrEmpty(text))
        {
            return;
        }
        var run = new Run();
        var rPr = BuildRunProperties(flags);
        if (rPr is not null)
        {
            run.Append(rPr);
        }
        run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        container.Append(run);
    }

    private void AppendCodeRun(OpenXmlCompositeElement container, string content, MarkdownRunFlags flags)
    {
        var run = new Run();
        run.Append(BuildRunProperties(flags | MarkdownRunFlags.Code) ?? new RunProperties());
        run.Append(new Text(content) { Space = SpaceProcessingModeValues.Preserve });
        container.Append(run);
    }

    /// <summary>
    /// Appends <paramref name="element"/> to <paramref name="container"/>. When the container is
    /// the document body, the element is inserted immediately before the final body-level
    /// <c>w:sectPr</c> instead of after it: CT_Body requires the section properties to be the
    /// last child, so content appended after a template's trailing <c>w:sectPr</c> would be
    /// schema-invalid. Section breaks nested inside paragraphs (paragraph-level <c>w:sectPr</c>)
    /// are not body children and are unaffected.
    /// </summary>
    private void AppendBlock(OpenXmlCompositeElement container, OpenXmlElement element)
    {
        if (ReferenceEquals(container, _body))
        {
            var finalSectionProperties = _body?.Elements<SectionProperties>().LastOrDefault();
            if (finalSectionProperties is not null)
            {
                _body!.InsertBefore(element, finalSectionProperties);
            }
            else
            {
                container.Append(element);
            }
            _rendered.Add(element);
        }
        else
        {
            container.Append(element);
        }
    }

    // ---- Formatting helpers ----------------------------------------------

    private RunProperties? BuildRunProperties(MarkdownRunFlags flags)
    {
        // CT_RPr child order: rStyle, rFonts, b, bCs, i, iCs, …, strike, …, color, …,
        // highlight, u, …, shd, …, vertAlign, …
        var rPr = new RunProperties();

        if (flags.HasFlag(MarkdownRunFlags.Code) || flags.HasFlag(MarkdownRunFlags.Hyperlink))
        {
            // Inline code and hyperlinks are character styles, so the resolved style id goes
            // into rStyle (the first run-property element) rather than into paragraph styles.
            var key = flags.HasFlag(MarkdownRunFlags.Code) ? "codeInline" : "hyperlink";
            var styleId = ResolveCharacterStyle(key);
            if (styleId is not null)
            {
                rPr.Append(new RunStyle { Val = styleId });
            }
        }
        if (flags.HasFlag(MarkdownRunFlags.Code))
        {
            rPr.Append(new RunFonts { Ascii = "Consolas", HighAnsi = "Consolas", ComplexScript = "Consolas" });
        }
        if (flags.HasFlag(MarkdownRunFlags.Bold))
        {
            rPr.Append(new Bold());
        }
        if (flags.HasFlag(MarkdownRunFlags.Italic))
        {
            rPr.Append(new Italic());
        }
        if (flags.HasFlag(MarkdownRunFlags.Strikethrough))
        {
            rPr.Append(new Strike());
        }
        if (flags.HasFlag(MarkdownRunFlags.Hyperlink))
        {
            rPr.Append(new Color { Val = "0563C1" });
        }
        if (flags.HasFlag(MarkdownRunFlags.Marked))
        {
            rPr.Append(new Highlight { Val = HighlightColorValues.Yellow });
        }
        if (flags.HasFlag(MarkdownRunFlags.Inserted) || flags.HasFlag(MarkdownRunFlags.Hyperlink))
        {
            rPr.Append(new Underline { Val = UnderlineValues.Single });
        }
        if (flags.HasFlag(MarkdownRunFlags.Code))
        {
            rPr.Append(new Shading { Val = ShadingPatternValues.Clear, Fill = "F2F2F2" });
        }
        if (flags.HasFlag(MarkdownRunFlags.Subscript))
        {
            rPr.Append(new VerticalTextAlignment { Val = VerticalPositionValues.Subscript });
        }
        else if (flags.HasFlag(MarkdownRunFlags.Superscript))
        {
            rPr.Append(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript });
        }

        return rPr.ChildElements.Count == 0 ? null : rPr;
    }

    private static MarkdownRunFlags FlagFor(EmphasisKind kind) => kind switch
    {
        EmphasisKind.Italic => MarkdownRunFlags.Italic,
        EmphasisKind.Bold => MarkdownRunFlags.Bold,
        EmphasisKind.Strikethrough => MarkdownRunFlags.Strikethrough,
        EmphasisKind.Subscript => MarkdownRunFlags.Subscript,
        EmphasisKind.Superscript => MarkdownRunFlags.Superscript,
        EmphasisKind.Inserted => MarkdownRunFlags.Inserted,
        EmphasisKind.Marked => MarkdownRunFlags.Marked,
        _ => MarkdownRunFlags.None
    };

    private static Indentation NumberedIndentation(int depth) =>
        new() { Left = (720 + depth * 720).ToString(), Hanging = "360" };

    private static Indentation ContinuationIndentation(int depth) =>
        new() { Left = (360 + depth * 720).ToString() };

    private static Indentation QuoteIndentation(int depth) =>
        new() { Left = (720 * (depth + 1)).ToString() };

    private static int ParseStart(string? orderedStart) =>
        int.TryParse(orderedStart, out var value) ? Math.Max(value, 1) : 1;

    /// <summary>
    /// True for an empty, self-contained anchor like <c>&lt;a id="…"&gt;&lt;/a&gt;. Markdig's
    /// AutoIdentifiers injects exactly this shape into headings as navigation scaffolding; it
    /// carries no visible text, so the renderer omits it.
    /// </summary>
    private static bool IsEmptyAnchorTag(string tag)
    {
        var trimmed = tag.Trim();
        if (!trimmed.StartsWith("<a", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }
        var closing = trimmed.IndexOf('>');
        if (closing < 0)
        {
            return false;
        }
        return trimmed[(closing + 1)..].TrimStart().StartsWith("</a>", StringComparison.OrdinalIgnoreCase);
    }

    private static JustificationValues? GetColumnAlignment(MarkdownTable table, int columnIndex)
    {
        if (columnIndex < 0 || columnIndex >= table.Columns.Count || table.Columns[columnIndex].Alignment is not { } alignment)
        {
            return null;
        }
        return alignment switch
        {
            MarkdownTableAlignment.Center => JustificationValues.Center,
            MarkdownTableAlignment.Right => JustificationValues.Right,
            _ => JustificationValues.Left
        };
    }

    private static IEnumerable<string> SplitLines(string text)
    {
        if (text.Length == 0)
        {
            return [""];
        }
        return text.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n').Split('\n');
    }

    private static string PlainText(IEnumerable<MarkdownInline> inlines)
    {
        var sb = new System.Text.StringBuilder();
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case MarkdownText t:
                    sb.Append(t.Text);
                    break;
                case MarkdownEntity e:
                    sb.Append(e.Decoded);
                    break;
                case MarkdownEmoji em:
                    sb.Append(em.Text);
                    break;
                case MarkdownCode c:
                    sb.Append(c.Content);
                    break;
                case MarkdownEmphasis m:
                    sb.Append(PlainText(m.Children));
                    break;
                case MarkdownLink l:
                    sb.Append(PlainText(l.Children));
                    break;
                case MarkdownImage img:
                    sb.Append(PlainText(img.Children));
                    break;
            }
        }
        return sb.ToString();
    }

    private int AllocateNumbering(bool ordered, int start)
    {
        if (_allocateNumbering is not null)
        {
            return _allocateNumbering(ordered, start);
        }
        return _fallbackNumbering!.Allocate(_mainPart, ordered, start);
    }

    private static (double WidthPt, double HeightPt) FitToMaxWidth(ImageAsset asset, double maxWidthPt)
    {
        var widthPt = asset.NaturalWidthPt;
        var heightPt = asset.NaturalHeightPt;
        if (maxWidthPt > 0 && widthPt > maxWidthPt)
        {
            var ratio = maxWidthPt / widthPt;
            widthPt = maxWidthPt;
            heightPt *= ratio;
        }
        return (widthPt, heightPt);
    }

    private static uint ComputeNextDrawingId(MainDocumentPart mainPart)
    {
        uint max = 0;
        max = Math.Max(max, MaxDrawingId(mainPart.Document));
        foreach (var header in mainPart.HeaderParts)
        {
            max = Math.Max(max, MaxDrawingId(header.Header));
        }
        foreach (var footer in mainPart.FooterParts)
        {
            max = Math.Max(max, MaxDrawingId(footer.Footer));
        }
        if (mainPart.FootnotesPart?.Footnotes is { } footnotes)
        {
            max = Math.Max(max, MaxDrawingId(footnotes));
        }
        return max + 1;
    }

    private static uint MaxDrawingId(OpenXmlElement? root)
    {
        uint max = 0;
        if (root is null)
        {
            return max;
        }
        foreach (var docProperties in root.Descendants<Wp.DocProperties>())
        {
            max = Math.Max(max, docProperties.Id?.Value ?? 0);
        }
        foreach (var nonVisual in root.Descendants<A.NonVisualDrawingProperties>())
        {
            max = Math.Max(max, nonVisual.Id?.Value ?? 0);
        }
        return max;
    }

    private void Warn(string message)
    {
        if (!_options.Strict)
        {
            return;
        }
        _diagnostics.Add(new MarkdownDiagnostic(MarkdownDiagnosticSeverity.Warning, message));
    }
}
