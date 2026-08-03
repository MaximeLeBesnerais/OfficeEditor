using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocxEditor.Core.Markdown;

/// <summary>
/// The outcome of a DOCX → Markdown conversion: the generated Markdown plus any warnings
/// raised while degrading unsupported constructs, and the absolute paths of image assets
/// extracted next to the output (empty when images were embedded as data URIs).
/// </summary>
public sealed record DocxToMarkdownResult
{
    public string Markdown { get; init; } = string.Empty;

    public IReadOnlyList<string> Warnings { get; init; } = [];

    public IReadOnlyList<string> ExtractedAssetPaths { get; init; } = [];
}

/// <summary>
/// Reads a WordprocessingML document and emits CommonMark-ish Markdown. This is the reverse
/// of <see cref="Rendering.RichMarkdownRenderer"/>: headings, paragraphs with run emphasis,
/// inline code, hyperlinks, lists (numbering-aware, with nesting from <c>w:ilvl</c> and
/// <c>w:ind</c>), pipe tables, block quotes, horizontal rules, fenced code blocks, images
/// (extracted to assets or embedded as data URIs), hard line breaks, tabs and escaped special
/// characters. Constructs that have no Markdown equivalent are degraded losslessly-enough
/// (the visible text is kept) and reported through <see cref="DocxToMarkdownResult.Warnings"/>.
/// </summary>
public sealed class DocxToMarkdownWriter
{
    private readonly WordprocessingDocument _document;
    private readonly Dictionary<string, Style> _stylesById = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _numberingInstances = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, Dictionary<int, Level>> _abstractLevels = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<string> _warnings = [];
    private readonly HashSet<string> _warnedMessages = new(StringComparer.Ordinal);
    private readonly List<string> _extractedAssetPaths = [];
    private string? _assetsDirectory;

    /// <summary>Creates a writer over an opened document. The document is not owned or disposed.</summary>
    public DocxToMarkdownWriter(WordprocessingDocument document)
    {
        _document = document ?? throw new ArgumentNullException(nameof(document));
        LoadStyles();
        LoadNumbering();
    }

    /// <summary>
    /// Converts the document to Markdown. Images are embedded as <c>data:</c> URIs so the result
    /// is fully self-contained.
    /// </summary>
    public DocxToMarkdownResult Convert() => ConvertCore(assetsDirectory: null);

    /// <summary>
    /// Converts the document to Markdown, extracting images into <paramref name="assetsDirectory"/>
    /// and referencing them by <c>&lt;folder&gt;/&lt;file&gt;</c> relative paths. The directory is
    /// created on first use; the Markdown file must live in its parent directory for the relative
    /// references to resolve.
    /// </summary>
    public DocxToMarkdownResult Convert(string assetsDirectory)
    {
        ArgumentNullException.ThrowIfNull(assetsDirectory);
        if (string.IsNullOrWhiteSpace(assetsDirectory))
        {
            throw new ArgumentException("Assets directory must not be empty or whitespace.", nameof(assetsDirectory));
        }

        return ConvertCore(assetsDirectory);
    }

    /// <summary>
    /// Opens <paramref name="sourcePath"/>, converts it to Markdown (images extracted next to the
    /// output) and writes the result to <paramref name="outputPath"/>. Assets land in a
    /// <c>&lt;output-name&gt;_assets</c> sibling directory.
    /// </summary>
    public static DocxToMarkdownResult ConvertToFile(string sourcePath, string outputPath)
    {
        ValidatePath(sourcePath, nameof(sourcePath));
        ValidatePath(outputPath, nameof(outputPath));

        string outputDirectory = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
        string assetsDirectory = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(outputPath) + "_assets");

        using WordprocessingDocument document = WordprocessingDocument.Open(sourcePath, false);
        DocxToMarkdownResult result = new DocxToMarkdownWriter(document).Convert(assetsDirectory);
        File.WriteAllText(outputPath, result.Markdown, new UTF8Encoding(false));
        return result;
    }

    /// <summary>
    /// Opens <paramref name="sourcePath"/> and returns its Markdown representation with images
    /// embedded as data URIs (convenience wrapper that also owns document lifecycle).
    /// </summary>
    public static DocxToMarkdownResult ConvertFile(string sourcePath)
    {
        ValidatePath(sourcePath, nameof(sourcePath));

        using WordprocessingDocument document = WordprocessingDocument.Open(sourcePath, false);
        return new DocxToMarkdownWriter(document).Convert();
    }

    private static void ValidatePath(string path, string parameterName)
    {
        ArgumentNullException.ThrowIfNull(path);
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("Path must not be empty or whitespace.", parameterName);
        }
    }

    private DocxToMarkdownResult ConvertCore(string? assetsDirectory)
    {
        _warnings.Clear();
        _warnedMessages.Clear();
        _extractedAssetPaths.Clear();
        _assetsDirectory = assetsDirectory;

        Body? body = _document.MainDocumentPart?.Document?.Body;
        if (body is null)
        {
            return new DocxToMarkdownResult { Warnings = _warnings.ToList() };
        }

        List<string> blocks = [];
        List<OpenXmlElement> elements = body.Elements().ToList();
        int i = 0;
        while (i < elements.Count)
        {
            switch (elements[i])
            {
                case Paragraph paragraph when IsListParagraph(paragraph):
                    (int listEnd, string listMarkdown) = ConvertList(elements, i);
                    if (listMarkdown.Length > 0)
                    {
                        blocks.Add(listMarkdown);
                    }

                    i = listEnd + 1;
                    break;
                case Paragraph paragraph when IsCodeBlockParagraph(paragraph):
                    (int codeEnd, string codeMarkdown) = ConvertCodeBlock(elements, i);
                    if (codeMarkdown.Length > 0)
                    {
                        blocks.Add(codeMarkdown);
                    }

                    i = codeEnd + 1;
                    break;
                case Paragraph paragraph when IsQuoteParagraph(paragraph):
                    (int quoteEnd, string quoteMarkdown) = ConvertQuote(elements, i);
                    if (quoteMarkdown.Length > 0)
                    {
                        blocks.Add(quoteMarkdown);
                    }

                    i = quoteEnd + 1;
                    break;
                case Paragraph paragraph:
                    blocks.AddRange(ConvertBodyParagraph(paragraph).Where(block => block.Length > 0));
                    i++;
                    break;
                case Table table:
                    string tableMarkdown = ConvertTable(table);
                    if (tableMarkdown.Length > 0)
                    {
                        blocks.Add(tableMarkdown);
                    }

                    i++;
                    break;
                default:
                    i++;
                    break;
            }
        }

        string markdown = string.Join("\n\n", blocks);
        if (markdown.Length > 0)
        {
            markdown += "\n";
        }

        return new DocxToMarkdownResult
        {
            Markdown = markdown,
            Warnings = _warnings.ToList(),
            ExtractedAssetPaths = _extractedAssetPaths.ToList()
        };
    }

    // ---- Block conversion -------------------------------------------------

    private List<string> ConvertBodyParagraph(Paragraph paragraph)
    {
        if (GetHeadingLevel(paragraph) is int level)
        {
            string content = RenderInlineItems(CollectInlineItems(paragraph), suppressFirstLine: true);
            return [new string('#', level) + (content.Length > 0 ? " " + content : string.Empty)];
        }

        if (IsThematicBreakParagraph(paragraph))
        {
            return ["---"];
        }

        List<string> blocks = [];
        string text = RenderInlineItems(CollectInlineItems(paragraph));
        if (text.Length > 0)
        {
            blocks.Add(text);
        }

        // Text boxes hold real content that would otherwise be dropped; flatten it as
        // paragraphs (positioning is not representable in Markdown).
        foreach (Paragraph textboxParagraph in GetTextboxParagraphs(paragraph))
        {
            string textboxText = RenderInlineItems(CollectInlineItems(textboxParagraph, inTextbox: true));
            if (textboxText.Length > 0)
            {
                Warn("Text box content flattened into a paragraph (positioning lost).");
                blocks.Add(textboxText);
            }
        }

        return blocks;
    }

    private (int LastIndex, string Markdown) ConvertList(IReadOnlyList<OpenXmlElement> elements, int startIndex)
    {
        List<ListItemInfo> items = [];
        int i = startIndex;
        while (i < elements.Count && elements[i] is Paragraph paragraph && IsListParagraph(paragraph))
        {
            int ilvl = GetNumberingLevel(paragraph);
            int left = GetIndentLeftTwips(paragraph) ?? 720;
            bool ordered = IsOrderedListParagraph(paragraph);
            int start = GetListStart(paragraph);
            int depth;
            if (items.Count == 0)
            {
                depth = ilvl > 0 ? ilvl : Math.Max(0, (left - 720) / 720);
            }
            else
            {
                ListItemInfo previous = items[^1];
                if (ilvl != previous.Ilvl)
                {
                    depth = ilvl;
                }
                else if (left != previous.Left)
                {
                    depth = left > previous.Left ? previous.Depth + 1 : Math.Max(0, previous.Depth - 1);
                }
                else
                {
                    depth = previous.Depth;
                }
            }

            items.Add(new ListItemInfo(paragraph, depth, ordered, start, ilvl, left));
            i++;
        }

        var sb = new StringBuilder();
        var orderedCounters = new Dictionary<int, int>();
        for (int k = 0; k < items.Count; k++)
        {
            ListItemInfo item = items[k];
            string indent = new string(' ', item.Depth * 4);
            string marker;
            int markerWidth;
            if (item.Ordered)
            {
                bool continuesSequence = k > 0 && items[k - 1].Ordered && items[k - 1].Depth == item.Depth;
                if (!continuesSequence)
                {
                    orderedCounters[item.Depth] = 0;
                }

                orderedCounters[item.Depth]++;
                int number = orderedCounters[item.Depth];
                int markerNumber = number == 1 && item.Start != 1 ? item.Start : number;
                marker = markerNumber.ToString(CultureInfo.InvariantCulture) + ". ";
            }
            else
            {
                marker = "- ";
            }

            markerWidth = marker.Length;
            string content = RenderInlineItems(CollectInlineItems(item.Paragraph));
            if (content.Contains('\n', StringComparison.Ordinal))
            {
                content = IndentContinuationLines(content, item.Depth * 4 + markerWidth);
            }

            if (k > 0)
            {
                ListItemInfo previous = items[k - 1];
                bool nestedEnded = item.Depth < previous.Depth;
                bool markerTypeChanged = item.Depth == previous.Depth && item.Ordered != previous.Ordered;
                if (nestedEnded || markerTypeChanged)
                {
                    sb.Append('\n');
                }
            }

            sb.Append(indent).Append(marker).Append(content).Append('\n');
        }

        // Lazy continuation: a non-list paragraph indented more than the last item is the
        // second (or later) block of that item in Markdown.
        ListItemInfo last = items[^1];
        while (i < elements.Count && elements[i] is Paragraph continuation && IsContinuationParagraph(continuation, last))
        {
            string continuationText = RenderInlineItems(CollectInlineItems(continuation));
            if (continuationText.Length > 0)
            {
                sb.Append(new string(' ', (last.Depth + 1) * 4)).Append(continuationText).Append('\n');
            }

            i++;
        }

        return (i - 1, sb.ToString().TrimEnd('\n'));
    }

    private bool IsContinuationParagraph(Paragraph paragraph, ListItemInfo last)
    {
        if (IsListParagraph(paragraph)
            || GetHeadingLevel(paragraph) is not null
            || IsCodeBlockParagraph(paragraph)
            || IsQuoteParagraph(paragraph)
            || IsThematicBreakParagraph(paragraph))
        {
            return false;
        }

        int left = GetIndentLeftTwips(paragraph) ?? 720;
        return left > last.Left;
    }

    private (int LastIndex, string Markdown) ConvertCodeBlock(IReadOnlyList<OpenXmlElement> elements, int startIndex)
    {
        List<string> lines = [];
        int i = startIndex;
        while (i < elements.Count && elements[i] is Paragraph paragraph && IsCodeBlockParagraph(paragraph))
        {
            lines.Add(paragraph.InnerText);
            i++;
        }

        // Pick a fence longer than any backtick run in the content so the block stays open.
        int longestRun = 0;
        int currentRun = 0;
        foreach (string line in lines)
        {
            foreach (char c in line)
            {
                if (c == '`')
                {
                    currentRun++;
                    longestRun = Math.Max(longestRun, currentRun);
                }
                else
                {
                    currentRun = 0;
                }
            }
        }

        string fence = new string('`', Math.Max(3, longestRun + 1));
        var sb = new StringBuilder();
        sb.Append(fence).Append('\n');
        foreach (string line in lines)
        {
            sb.Append(line).Append('\n');
        }

        sb.Append(fence);
        return (i - 1, sb.ToString());
    }

    private (int LastIndex, string Markdown) ConvertQuote(IReadOnlyList<OpenXmlElement> elements, int startIndex)
    {
        var sb = new StringBuilder();
        int i = startIndex;
        while (i < elements.Count && elements[i] is Paragraph paragraph && IsQuoteParagraph(paragraph))
        {
            string text = RenderInlineItems(CollectInlineItems(paragraph), suppressFirstLine: true);
            foreach (string line in text.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n'))
            {
                sb.Append("> ").Append(line).Append('\n');
            }

            i++;
        }

        return (i - 1, sb.ToString().TrimEnd('\n'));
    }

    private string ConvertTable(Table table)
    {
        List<TableRow> rows = table.Elements<TableRow>().ToList();
        if (rows.Count == 0)
        {
            return string.Empty;
        }

        int columnCount = rows.Max(row => row.Elements<TableCell>().Count());
        if (columnCount == 0)
        {
            return string.Empty;
        }

        var headerRows = new HashSet<int>();
        for (int r = 0; r < rows.Count; r++)
        {
            if (rows[r].TableRowProperties?.GetFirstChild<TableHeader>() is not null)
            {
                headerRows.Add(r);
            }
        }

        if (headerRows.Count == 0)
        {
            Warn("Table has no header row; the first row is emitted as the Markdown header.");
            headerRows.Add(0);
        }

        int headerIndex = headerRows.Min();
        if (headerRows.Count > 1)
        {
            Warn("Table has multiple header rows; only the first is emitted as the Markdown header.");
        }

        string[] alignments = ReadColumnAlignments(rows[headerIndex], columnCount);
        var sb = new StringBuilder();
        sb.Append(RenderTableRow(rows[headerIndex], columnCount, isHeader: true)).Append('\n');
        sb.Append(RenderTableSeparator(alignments)).Append('\n');
        for (int r = 0; r < rows.Count; r++)
        {
            if (r == headerIndex)
            {
                continue;
            }

            sb.Append(RenderTableRow(rows[r], columnCount)).Append('\n');
        }

        return sb.ToString().TrimEnd('\n');
    }

    private string RenderTableRow(TableRow row, int columnCount, bool isHeader = false)
    {
        List<string> cells = [];
        foreach (TableCell cell in row.Elements<TableCell>())
        {
            cells.Add(RenderTableCell(cell, isHeader));
        }

        while (cells.Count < columnCount)
        {
            cells.Add(string.Empty);
        }

        return "| " + string.Join(" | ", cells) + " |";
    }

    private string RenderTableCell(TableCell cell, bool isHeader = false)
    {
        List<string> parts = [];
        foreach (Paragraph paragraph in cell.Elements<Paragraph>())
        {
            string text = RenderInlineItems(CollectInlineItems(paragraph, suppressBold: isHeader), escapePipe: true, suppressFirstLine: true);
            if (text.Length > 0)
            {
                parts.Add(text);
            }
        }

        foreach (Table nested in cell.Elements<Table>())
        {
            Warn("Nested table flattened to inline text inside a Markdown table cell.");
            string inner = nested.InnerText.Trim();
            if (inner.Length > 0)
            {
                parts.Add(inner);
            }
        }

        return string.Join("<br>", parts);
    }

    private static string RenderTableSeparator(IReadOnlyList<string> alignments)
    {
        List<string> parts = [];
        foreach (string alignment in alignments)
        {
            parts.Add(alignment switch
            {
                "center" => ":---:",
                "right" => "---:",
                _ => "---"
            });
        }

        return "| " + string.Join(" | ", parts) + " |";
    }

    private static string[] ReadColumnAlignments(TableRow row, int columnCount)
    {
        var result = new string[columnCount];
        int index = 0;
        foreach (TableCell cell in row.Elements<TableCell>())
        {
            if (index >= columnCount)
            {
                break;
            }

            JustificationValues? justification = cell.Elements<Paragraph>()
                .FirstOrDefault()?
                .ParagraphProperties?
                .GetFirstChild<Justification>()?
                .Val?.Value;
            result[index] = justification == JustificationValues.Center
                ? "center"
                : justification == JustificationValues.Right
                    ? "right"
                    : "left";
            index++;
        }

        return result;
    }

    // ---- Inline conversion -------------------------------------------------

    private List<InlineItem> CollectInlineItems(Paragraph paragraph, bool inTextbox = false, bool suppressBold = false)
    {
        var items = new List<InlineItem>();
        foreach (OpenXmlElement child in paragraph.ChildElements)
        {
            ConvertInlineElement(child, items, linkUrl: null, inTextbox, suppressBold);
        }

        return items;
    }

    private void ConvertInlineElement(OpenXmlElement element, List<InlineItem> items, string? linkUrl, bool inTextbox, bool suppressBold)
    {
        switch (element)
        {
            case Run run:
                ConvertRun(run, items, linkUrl, inTextbox, suppressBold);
                break;
            case Hyperlink hyperlink:
                string? resolvedUrl = ResolveHyperlinkUrl(hyperlink);
                foreach (OpenXmlElement child in hyperlink.ChildElements)
                {
                    ConvertInlineElement(child, items, resolvedUrl, inTextbox, suppressBold);
                }

                break;
            case SimpleField field:
                foreach (OpenXmlElement child in field.ChildElements)
                {
                    ConvertInlineElement(child, items, linkUrl, inTextbox, suppressBold);
                }

                break;
            case Deleted:
                // Tracked deletions are not document content; skip them.
                break;
            case BookmarkStart or BookmarkEnd or ParagraphProperties:
                break;
            default:
                foreach (OpenXmlElement child in element.ChildElements)
                {
                    ConvertInlineElement(child, items, linkUrl, inTextbox, suppressBold);
                }

                break;
        }
    }

    private void ConvertRun(Run run, List<InlineItem> items, string? linkUrl, bool inTextbox, bool suppressBold)
    {
        if (IsInTextBoxContent(run) && !inTextbox)
        {
            return;
        }

        (bool bold, bool italic, bool strike, bool code, bool underline) = ComputeRunFormatting(run);
        if (suppressBold)
        {
            // Pipe-table header cells are rendered bold by Markdown convention, so the bold the
            // renderer forces on them is not re-emitted as emphasis.
            bold = false;
        }

        if (underline)
        {
            Warn("Underline formatting has no Markdown equivalent; text preserved.");
        }

        foreach (OpenXmlElement child in run.ChildElements)
        {
            switch (child)
            {
                case Text text when !string.IsNullOrEmpty(text.Text):
                    items.Add(new InlineItem(InlineItemKind.Text, text.Text, bold, italic, strike, code, linkUrl));
                    break;
                case Break breakElement when breakElement.Type?.Value != BreakValues.Page:
                    items.Add(new InlineItem(InlineItemKind.LineBreak));
                    break;
                case TabChar:
                    items.Add(new InlineItem(InlineItemKind.Tab));
                    break;                case SymbolChar symbol:
                    if (symbol.Char?.Value is string hex
                        && int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int codePoint))
                    {
                        items.Add(new InlineItem(InlineItemKind.Text, char.ConvertFromUtf32(codePoint), bold, italic, strike, code, linkUrl));
                    }

                    break;
                case Drawing drawing:
                    string image = ConvertDrawing(drawing);
                    if (image.Length > 0)
                    {
                        items.Add(new InlineItem(InlineItemKind.Image, image));
                    }

                    break;
                case Picture:
                    Warn("Legacy VML picture skipped; no Markdown equivalent.");
                    break;
            }
        }
    }

    private string ConvertDrawing(Drawing drawing)
    {
        A.Blip? blip = drawing.Descendants<A.Blip>().FirstOrDefault();
        return blip is null ? string.Empty : ConvertBlip(blip);
    }

    private string ConvertBlip(A.Blip blip)
    {
        string? relationshipId = blip.Embed?.Value ?? blip.Link?.Value;
        if (string.IsNullOrWhiteSpace(relationshipId))
        {
            Warn("Image without a relationship id was skipped.");
            return string.Empty;
        }

        string alt = ExtractAltText(blip);
        MainDocumentPart mainPart = _document.MainDocumentPart!;

        OpenXmlPart? part;
        try
        {
            part = mainPart.GetPartById(relationshipId);
        }
        catch (ArgumentOutOfRangeException)
        {
            part = null;
        }

        if (part is ImagePart imagePart)
        {
            string contentType = imagePart.ContentType;
            string extension = ContentTypeToExtension(contentType);
            byte[] bytes;
            using (Stream stream = imagePart.GetStream(FileMode.Open, FileAccess.Read))
            using (var buffer = new MemoryStream())
            {
                stream.CopyTo(buffer);
                bytes = buffer.ToArray();
            }

            if (IsUnsupportedImageExtension(extension))
            {
                Warn($"Image format '{extension}' may not render in all Markdown viewers.");
            }

            if (_assetsDirectory is null)
            {
                return $"![{alt}](data:{contentType};base64,{System.Convert.ToBase64String(bytes)})";
            }

            string fileName = $"image-{_extractedAssetPaths.Count + 1}{extension}";
            Directory.CreateDirectory(_assetsDirectory);
            string path = Path.Combine(_assetsDirectory, fileName);
            File.WriteAllBytes(path, bytes);
            _extractedAssetPaths.Add(path);
            string folderName = Path.GetFileName(_assetsDirectory.TrimEnd('/', '\\'));
            return $"![{alt}]({EscapeUrl(folderName + "/" + fileName)})";
        }

        ExternalRelationship? external = mainPart.ExternalRelationships.FirstOrDefault(rel => rel.Id == relationshipId);
        if (external is not null)
        {
            return $"![{alt}]({EscapeUrl(external.Uri.OriginalString)})";
        }

        Warn($"Image relationship '{relationshipId}' could not be resolved; image omitted.");
        return string.Empty;
    }

    private static string ExtractAltText(A.Blip blip)
    {
        // The descriptive wp:docPr is a sibling of the graphic inside wp:inline / wp:anchor,
        // not an ancestor of the blip, so it is read from those containers.
        DW.DocProperties? docProperties = blip.Ancestors<DW.Inline>().FirstOrDefault()?.GetFirstChild<DW.DocProperties>()
            ?? blip.Ancestors<DW.Anchor>().FirstOrDefault()?.GetFirstChild<DW.DocProperties>();
        string? alt = docProperties?.Description?.Value;
        if (string.IsNullOrWhiteSpace(alt))
        {
            alt = docProperties?.Name?.Value;
        }

        if (string.IsNullOrWhiteSpace(alt))
        {
            A.NonVisualDrawingProperties? drawingProperties = blip.Ancestors<A.NonVisualDrawingProperties>().FirstOrDefault();
            alt = drawingProperties?.Description?.Value ?? drawingProperties?.Name?.Value;
        }

        return string.IsNullOrWhiteSpace(alt) ? "image" : EscapeInlineText(alt.Trim());
    }

    private string? ResolveHyperlinkUrl(Hyperlink hyperlink)
    {
        string? anchor = hyperlink.Anchor?.Value;
        if (!string.IsNullOrWhiteSpace(anchor))
        {
            return "#" + anchor;
        }

        string? relationshipId = hyperlink.Id?.Value;
        if (string.IsNullOrWhiteSpace(relationshipId))
        {
            return null;
        }

        HyperlinkRelationship? relationship = _document.MainDocumentPart?.HyperlinkRelationships
            .FirstOrDefault(rel => rel.Id == relationshipId);
        if (relationship is null)
        {
            Warn($"Hyperlink relationship '{relationshipId}' could not be resolved; rendered as plain text.");
            return null;
        }

        return relationship.IsExternal
            ? relationship.Uri.OriginalString
            : "#" + relationship.Uri?.OriginalString.TrimStart('#');
    }

    private string RenderInlineItems(List<InlineItem> items, bool escapePipe = false, bool suppressFirstLine = false)
    {
        // Keep emphasis markers away from whitespace so they always open/close cleanly, then
        // merge adjacent text runs with identical formatting into a single span.
        var normalized = new List<InlineItem>();
        foreach (InlineItem item in items)
        {
            if (item.Kind == InlineItemKind.Text)
            {
                normalized.AddRange(NormalizeTextBoundaries(item));
            }
            else
            {
                normalized.Add(item);
            }
        }

        var merged = new List<InlineItem>();
        foreach (InlineItem item in normalized)
        {
            if (item.Kind == InlineItemKind.Text
                && merged.Count > 0
                && merged[^1].Kind == InlineItemKind.Text
                && SameFormatting(merged[^1], item))
            {
                merged[^1] = merged[^1] with { Text = merged[^1].Text + item.Text };
            }
            else
            {
                merged.Add(item);
            }
        }

        var sb = new StringBuilder();
        foreach (InlineItem item in merged)
        {
            switch (item.Kind)
            {
                case InlineItemKind.Text:
                    sb.Append(RenderSegment(item, escapePipe, suppressFirstLine));
                    break;
                case InlineItemKind.LineBreak:
                    sb.Append("  \n");
                    break;
                case InlineItemKind.Tab:
                    sb.Append('\t');
                    break;
                case InlineItemKind.Image:
                    sb.Append(item.Text);
                    break;
            }
        }

        return sb.ToString();
    }

    private static string RenderSegment(InlineItem segment, bool escapePipe, bool suppressFirstLine)
    {
        string text = segment.Code
            ? BuildCodeSpan(segment.Text)
            : EscapeInlineText(segment.Text, escapePipe, suppressFirstLine);
        if (segment.Strike)
        {
            text = "~~" + text + "~~";
        }

        if (segment.Italic)
        {
            text = "*" + text + "*";
        }

        if (segment.Bold)
        {
            text = "**" + text + "**";
        }

        if (segment.LinkUrl is not null)
        {
            text = "[" + text + "](" + EscapeUrl(segment.LinkUrl) + ")";
        }

        return text;
    }

    private static IEnumerable<InlineItem> NormalizeTextBoundaries(InlineItem item)
    {
        bool formatted = item.Bold || item.Italic || item.Strike || item.Code || item.LinkUrl is not null;
        if (!formatted || item.Text.Length == 0)
        {
            yield return item;
            yield break;
        }

        int leading = 0;
        while (leading < item.Text.Length && char.IsWhiteSpace(item.Text[leading]))
        {
            leading++;
        }

        int trailing = 0;
        while (trailing < item.Text.Length - leading && char.IsWhiteSpace(item.Text[item.Text.Length - 1 - trailing]))
        {
            trailing++;
        }

        int coreLength = item.Text.Length - leading - trailing;
        if (coreLength <= 0)
        {
            yield return PlainText(item.Text);
            yield break;
        }

        if (leading > 0)
        {
            yield return PlainText(item.Text[..leading]);
        }

        yield return item with { Text = item.Text.Substring(leading, coreLength) };
        if (trailing > 0)
        {
            yield return PlainText(item.Text.Substring(item.Text.Length - trailing));
        }
    }

    private static InlineItem PlainText(string text) => new(InlineItemKind.Text, text);

    private static bool SameFormatting(InlineItem left, InlineItem right)
        => left.Bold == right.Bold
            && left.Italic == right.Italic
            && left.Strike == right.Strike
            && left.Code == right.Code
            && left.LinkUrl == right.LinkUrl;

    // ---- Construct detection -------------------------------------------------

    private int? GetHeadingLevel(Paragraph paragraph)
    {
        int? outline = paragraph.ParagraphProperties?.GetFirstChild<OutlineLevel>()?.Val?.Value;
        if (outline is int outlineValue)
        {
            return Math.Clamp(outlineValue + 1, 1, 6);
        }

        string? styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (TryParseHeadingStyleId(styleId, out int directLevel))
        {
            return directLevel;
        }

        foreach (Style style in GetStyleChain(styleId))
        {
            if (TryParseHeadingStyleId(style.StyleId?.Value, out int level))
            {
                return level;
            }

            if (TryParseHeadingStyleName(style.StyleName?.Val?.Value, out level))
            {
                return level;
            }

            int? styleOutline = style.StyleParagraphProperties?.GetFirstChild<OutlineLevel>()?.Val?.Value;
            if (styleOutline is int styleOutlineValue)
            {
                return Math.Clamp(styleOutlineValue + 1, 1, 6);
            }
        }

        return null;
    }

    private static bool TryParseHeadingStyleId(string? styleId, out int level)
    {
        Match match = Regex.Match(styleId ?? string.Empty, "^heading(\\d+)$", RegexOptions.IgnoreCase);
        if (match.Success && int.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value))
        {
            level = Math.Clamp(value, 1, 6);
            return true;
        }

        level = 0;
        return false;
    }

    private static bool TryParseHeadingStyleName(string? styleName, out int level)
    {
        Match match = Regex.Match(styleName ?? string.Empty, "^heading\\s*(\\d+)$", RegexOptions.IgnoreCase);
        if (match.Success && int.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value))
        {
            level = Math.Clamp(value, 1, 6);
            return true;
        }

        level = 0;
        return false;
    }

    private bool IsCodeBlockParagraph(Paragraph paragraph)
    {
        string? styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (IsCodeStyleId(styleId))
        {
            return true;
        }

        foreach (Style style in GetStyleChain(styleId))
        {
            if (IsCodeStyleId(style.StyleId?.Value)
                || style.StyleName?.Val?.Value is string name && name.Contains("code", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private bool IsQuoteParagraph(Paragraph paragraph)
    {
        string? styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (styleId?.Equals("Quote", StringComparison.OrdinalIgnoreCase) == true)
        {
            return true;
        }

        foreach (Style style in GetStyleChain(styleId))
        {
            if (string.Equals(style.StyleId?.Value, "Quote", StringComparison.OrdinalIgnoreCase)
                || style.StyleName?.Val?.Value is string name && name.Contains("quote", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsThematicBreakParagraph(Paragraph paragraph)
    {
        if (paragraph.InnerText.Length > 0)
        {
            return false;
        }

        BorderType? bottomBorder = paragraph.ParagraphProperties?.GetFirstChild<ParagraphBorders>()?.BottomBorder;
        return bottomBorder is not null
            && bottomBorder.Val?.Value is not null
            && bottomBorder.Val.Value != BorderValues.None
            && bottomBorder.Val.Value != BorderValues.Nil;
    }

    private static bool IsListParagraph(Paragraph paragraph)
        => paragraph.ParagraphProperties?.NumberingProperties?.NumberingId?.Val is not null;

    private static int GetNumberingLevel(Paragraph paragraph)
        => paragraph.ParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val?.Value ?? 0;

    private static string? GetNumberingId(Paragraph paragraph)
        => paragraph.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value.ToString(CultureInfo.InvariantCulture);

    private static int? GetIndentLeftTwips(Paragraph paragraph)
    {
        string? left = paragraph.ParagraphProperties?.GetFirstChild<Indentation>()?.Left?.Value;
        return int.TryParse(left, NumberStyles.Integer, CultureInfo.InvariantCulture, out int twips) ? twips : null;
    }

    private bool IsOrderedListParagraph(Paragraph paragraph)
    {
        Level? level = ResolveLevel(GetNumberingId(paragraph), GetNumberingLevel(paragraph));
        string? format = level?.GetFirstChild<NumberingFormat>()?.Val?.InnerText;
        return !string.Equals(format, "bullet", StringComparison.OrdinalIgnoreCase)
            && !string.IsNullOrEmpty(format);
    }

    private int GetListStart(Paragraph paragraph)
    {
        Level? level = ResolveLevel(GetNumberingId(paragraph), GetNumberingLevel(paragraph));
        return level?.GetFirstChild<StartNumberingValue>()?.Val?.Value ?? 1;
    }

    private Level? ResolveLevel(string? numberingId, int levelIndex)
    {
        if (numberingId is null
            || !_numberingInstances.TryGetValue(numberingId, out string? abstractId)
            || !_abstractLevels.TryGetValue(abstractId, out Dictionary<int, Level>? levels))
        {
            return null;
        }

        return levels.GetValueOrDefault(levelIndex);
    }

    // ---- Run formatting ----------------------------------------------------

    private (bool Bold, bool Italic, bool Strike, bool Code, bool Underline) ComputeRunFormatting(Run run)
    {
        RunProperties? runProperties = run.RunProperties;
        string? runStyleId = runProperties?.RunStyle?.Val?.Value;
        List<Style> characterStyles = GetStyleChain(runStyleId);

        bool? bold = null;
        if (runProperties?.GetFirstChild<Bold>() is Bold directBold)
        {
            bold = IsOn(directBold);
        }
        else
        {
            foreach (Style style in characterStyles)
            {
                if (style.StyleRunProperties?.GetFirstChild<Bold>() is Bold styleBold)
                {
                    bold = IsOn(styleBold);
                    break;
                }
            }
        }

        bool? italic = null;
        if (runProperties?.GetFirstChild<Italic>() is Italic directItalic)
        {
            italic = IsOn(directItalic);
        }
        else
        {
            foreach (Style style in characterStyles)
            {
                if (style.StyleRunProperties?.GetFirstChild<Italic>() is Italic styleItalic)
                {
                    italic = IsOn(styleItalic);
                    break;
                }
            }
        }

        bool? strike = null;
        if (runProperties?.GetFirstChild<Strike>() is Strike directStrike)
        {
            strike = IsOn(directStrike);
        }
        else if (runProperties?.GetFirstChild<DoubleStrike>() is DoubleStrike directDoubleStrike)
        {
            strike = IsOn(directDoubleStrike);
        }
        else
        {
            foreach (Style style in characterStyles)
            {
                if (style.StyleRunProperties?.GetFirstChild<Strike>() is Strike styleStrike)
                {
                    strike = IsOn(styleStrike);
                    break;
                }

                if (style.StyleRunProperties?.GetFirstChild<DoubleStrike>() is DoubleStrike styleDoubleStrike)
                {
                    strike = IsOn(styleDoubleStrike);
                    break;
                }
            }
        }

        bool code = IsCodeStyleId(runStyleId) || IsCodeFontShading(runProperties);
        if (!code)
        {
            foreach (Style style in characterStyles)
            {
                if (IsCodeStyleId(style.StyleId?.Value))
                {
                    code = true;
                    break;
                }
            }
        }

        Underline? underline = runProperties?.GetFirstChild<Underline>();
        bool underlineOn = underline is not null
            && underline.Val?.Value is not null
            && underline.Val.Value != UnderlineValues.None;
        if (!underlineOn)
        {
            foreach (Style style in characterStyles)
            {
                if (style.StyleRunProperties?.GetFirstChild<Underline>() is Underline styleUnderline
                    && styleUnderline.Val?.Value is not null
                    && styleUnderline.Val.Value != UnderlineValues.None)
                {
                    underlineOn = true;
                    break;
                }
            }
        }

        return (bold ?? false, italic ?? false, strike ?? false, code, underlineOn);
    }

    private static bool IsCodeStyleId(string? styleId)
        => styleId is not null && styleId.Contains("code", StringComparison.OrdinalIgnoreCase);

    private static bool IsCodeFontShading(RunProperties? runProperties)
    {
        if (runProperties?.GetFirstChild<Shading>() is null)
        {
            return false;
        }

        RunFonts? fonts = runProperties.GetFirstChild<RunFonts>();
        string? font = fonts?.Ascii?.Value ?? fonts?.HighAnsi?.Value;
        return IsMonospaceFont(font);
    }

    private static bool IsMonospaceFont(string? font)
        => font is not null
            && (font.Contains("Consolas", StringComparison.OrdinalIgnoreCase)
                || font.Contains("Courier", StringComparison.OrdinalIgnoreCase)
                || font.Contains("Mono", StringComparison.OrdinalIgnoreCase));

    // ---- Shared infrastructure -------------------------------------------------

    private void LoadStyles()
    {
        Styles? styles = _document.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        if (styles is null)
        {
            return;
        }

        foreach (Style style in styles.Elements<Style>())
        {
            string? id = style.StyleId?.Value;
            if (!string.IsNullOrWhiteSpace(id))
            {
                _stylesById[id] = style;
            }
        }
    }

    private void LoadNumbering()
    {
        Numbering? numbering = _document.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
        if (numbering is null)
        {
            return;
        }

        foreach (AbstractNum abstractNum in numbering.Elements<AbstractNum>())
        {
            string? abstractId = abstractNum.AbstractNumberId?.Value.ToString(CultureInfo.InvariantCulture);
            if (abstractId is null)
            {
                continue;
            }

            var levels = new Dictionary<int, Level>();
            foreach (Level level in abstractNum.Elements<Level>())
            {
                levels[level.LevelIndex?.Value ?? 0] = level;
            }

            _abstractLevels[abstractId] = levels;
        }

        foreach (NumberingInstance instance in numbering.Elements<NumberingInstance>())
        {
            string? numId = instance.NumberID?.Value.ToString(CultureInfo.InvariantCulture);
            string? abstractId = instance.AbstractNumId?.Val?.Value.ToString(CultureInfo.InvariantCulture);
            if (numId is not null && abstractId is not null)
            {
                _numberingInstances[numId] = abstractId;
            }
        }
    }

    private List<Style> GetStyleChain(string? styleId)
    {
        var chain = new List<Style>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        while (!string.IsNullOrWhiteSpace(styleId) && seen.Add(styleId) && _stylesById.TryGetValue(styleId, out Style? style))
        {
            chain.Insert(0, style);
            styleId = style.BasedOn?.Val?.Value;
        }

        return chain;
    }

    private void Warn(string message)
    {
        if (_warnedMessages.Add(message))
        {
            _warnings.Add(message);
        }
    }

    // ---- Text rendering helpers -------------------------------------------------

    private static string IndentContinuationLines(string text, int indent)
    {
        string pad = new string(' ', indent);
        string[] lines = text.Split('\n');
        for (int i = 1; i < lines.Length; i++)
        {
            lines[i] = pad + lines[i];
        }

        return string.Join("\n", lines);
    }

    private static string BuildCodeSpan(string content)
    {
        int longestRun = 0;
        int currentRun = 0;
        foreach (char c in content)
        {
            if (c == '`')
            {
                currentRun++;
                longestRun = Math.Max(longestRun, currentRun);
            }
            else
            {
                currentRun = 0;
            }
        }

        string delimiter = new string('`', Math.Max(longestRun + 1, 1));
        return delimiter + content + delimiter;
    }

    /// <summary>
    /// Escapes characters that would change meaning if the text were re-parsed as Markdown.
    /// Always-escaped: backslash, backtick, emphasis, link brackets, angle brackets, heading
    /// markers, subscript/superscript markers. Additionally at the start of a line: list
    /// markers (<c>-</c>, <c>+</c>, digits followed by <c>.</c>), block quotes (<c>&gt;</c>)
    /// and setext markers (<c>=</c>). When <paramref name="escapePipe"/> is set, pipe
    /// characters are escaped for table cells.
    /// </summary>
    private static string EscapeInlineText(string text, bool escapePipe = false, bool suppressFirstLine = false)
    {
        var sb = new StringBuilder(text.Length);
        int column = 0;
        bool onlySpaces = true;
        bool lineIsNumeric = true;
        bool firstLine = true;
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            if (c == '\n')
            {
                sb.Append('\n');
                column = 0;
                onlySpaces = true;
                lineIsNumeric = true;
                firstLine = false;
                continue;
            }

            bool atLineStart = !(suppressFirstLine && firstLine) && onlySpaces && column <= 3;
            switch (c)
            {
                case '\\':
                    sb.Append("\\\\");
                    break;
                case '`':
                    sb.Append("\\`");
                    break;
                case '*':
                    sb.Append("\\*");
                    break;
                case '_':
                    sb.Append("\\_");
                    break;
                case '[':
                    sb.Append("\\[");
                    break;
                case ']':
                    sb.Append("\\]");
                    break;
                case '<':
                    sb.Append("\\<");
                    break;
                case '#':
                    sb.Append("\\#");
                    break;
                case '~':
                    sb.Append("\\~");
                    break;
                case '^':
                    sb.Append("\\^");
                    break;
                case '|' when escapePipe:
                    sb.Append("\\|");
                    break;
                case '-' when atLineStart:
                case '+' when atLineStart:
                case '>' when atLineStart:
                case '=' when atLineStart:
                    sb.Append('\\').Append(c);
                    break;
                case '.' when !(suppressFirstLine && firstLine) && column <= 3 && lineIsNumeric && i > 0 && char.IsAsciiDigit(text[i - 1]):
                    // A leading "N." would re-parse as an ordered list marker.
                    sb.Append("\\.");
                    break;
                default:
                    sb.Append(c);
                    break;
            }

            column++;
            if (!char.IsWhiteSpace(c))
            {
                onlySpaces = false;
                if (!char.IsAsciiDigit(c))
                {
                    lineIsNumeric = false;
                }
            }
        }

        return sb.ToString();
    }

    private static string EscapeUrl(string url)
        => url.Replace("\\", "%5C", StringComparison.Ordinal)
            .Replace("(", "%28", StringComparison.Ordinal)
            .Replace(")", "%29", StringComparison.Ordinal)
            .Replace(" ", "%20", StringComparison.Ordinal);

    private static bool IsOn(OnOffType? value) => value is not null && (value.Val is null || value.Val.Value);

    private static bool IsInTextBoxContent(OpenXmlElement element)
        => element.Ancestors().Any(ancestor => ancestor.LocalName == "txbxContent");

    private static IEnumerable<Paragraph> GetTextboxParagraphs(Paragraph paragraph)
        => paragraph.Descendants<Paragraph>().Where(IsInTextBoxContent);

    private static string ContentTypeToExtension(string contentType) => contentType.ToLowerInvariant() switch
    {
        "image/png" => ".png",
        "image/jpeg" or "image/jpg" => ".jpg",
        "image/gif" => ".gif",
        "image/bmp" => ".bmp",
        "image/tiff" => ".tiff",
        "image/x-emf" or "image/emf" => ".emf",
        "image/x-wmf" or "image/wmf" => ".wmf",
        "image/svg+xml" => ".svg",
        _ => ".bin"
    };

    private static bool IsUnsupportedImageExtension(string extension)
        => extension.Equals(".emf", StringComparison.OrdinalIgnoreCase)
            || extension.Equals(".wmf", StringComparison.OrdinalIgnoreCase);

    private enum InlineItemKind
    {
        Text,
        LineBreak,
        Tab,
        Image
    }

    private sealed record InlineItem(InlineItemKind Kind, string Text = "", bool Bold = false, bool Italic = false, bool Strike = false, bool Code = false, string? LinkUrl = null);

    private sealed record ListItemInfo(Paragraph Paragraph, int Depth, bool Ordered, int Start, int Ilvl, int Left);
}
