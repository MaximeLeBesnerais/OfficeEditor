using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Content;
using DocxEditor.Core.Markdown;
using DocxEditor.Core.Markdown.Rendering;
using DocxEditor.Core.Models;
using OfficeEditor.Core.Exceptions;
using OfficeEditor.Core.Models;

namespace DocxEditor.Core.Builders;

public interface IDocumentBuilder : IDisposable
{
    IDocumentBuilder AddParagraph(string text, string? style = null);
    IDocumentBuilder InsertAfter(string targetText, string text, string? style = null);
    IDocumentBuilder InsertBefore(string targetText, string text, string? style = null);
    IDocumentBuilder ReplaceText(string find, string replace);
    IDocumentBuilder ReplaceParagraph(string targetText, string newText, string? style = null);
    IDocumentBuilder DeleteParagraph(string targetText);
    IDocumentBuilder ApplyStyle(string styleId);
    
    // Rich content
    IDocumentBuilder AddRichContent(List<ContentBlock> blocks);
    IDocumentBuilder ReplaceWithRichContent(string targetText, List<ContentBlock> blocks);
    
    // Hyperlinks
    IDocumentBuilder AddHyperlink(string url, string displayText, string? style = null);
    
    // Markdown
    IDocumentBuilder AddMarkdown(string markdown, StyleMapping? styleMap = null);
    IDocumentBuilder ReplaceWithMarkdown(string targetText, string markdown, StyleMapping? styleMap = null);
    
    // Rich Markdown (v2 IR, recursive rendering)
    // These members have default implementations so external IDocumentBuilder implementations
    // written against the pre-rich-markdown surface stay source-compatible: they compile
    // without implementing the new members, and only the default (a descriptive
    // NotSupportedException) runs for them. DocumentBuilder overrides all three, so its
    // behavior is unchanged.
    IDocumentBuilder AddRichMarkdown(string markdown, Markdown.Rendering.MarkdownRenderOptions? options = null)
    {
        throw new NotSupportedException(
            "This IDocumentBuilder implementation does not support rich markdown rendering (AddRichMarkdown).");
    }

    IDocumentBuilder ReplaceWithRichMarkdown(string targetText, string markdown, Markdown.Rendering.MarkdownRenderOptions? options = null)
    {
        throw new NotSupportedException(
            "This IDocumentBuilder implementation does not support rich markdown rendering (ReplaceWithRichMarkdown).");
    }

    Markdown.Rendering.MarkdownRenderResult? LastRichMarkdownResult => null;
    
    // Variables
    List<VariableInfo> DetectVariables();
    IDocumentBuilder MergeVariables(Dictionary<string, string> data);
    
    // Publipostage
    void MergeBatch(List<Dictionary<string, string>> records, string outputPattern, string? templatePath = null);

    void Save(string? path = null);
    void Save(Stream stream);
    byte[] SaveToBytes();

    static abstract IDocumentBuilder Create();
    static abstract IDocumentBuilder Open(Stream stream);
    static abstract IDocumentBuilder Open(byte[] bytes);
}

public class DocumentBuilder : IDocumentBuilder
{
    private readonly WordprocessingDocument _document;
    private readonly Body _body;
    private readonly bool _isNewDocument;
    private readonly Dictionary<string, Style> _cachedStyles;
    private readonly Dictionary<(bool Ordered, int Start), int> _generatedAbstractNumberingIds;
    private readonly string? _filePath;
    private readonly MemoryStream? _documentStream;

    private DocumentBuilder(WordprocessingDocument document, bool isNew, string? filePath, MemoryStream? documentStream = null)
    {
        _document = document;
        _isNewDocument = isNew;
        _filePath = filePath;
        _documentStream = documentStream;

        var mainPart = GetRequiredMainPart(document);
        var documentRoot = mainPart.Document
            ?? throw new OfficeEditorException(
                "The main document part does not contain a <w:document> root element.");
        _body = documentRoot.Body
            ?? throw new OfficeEditorException(
                "The document does not contain a <w:body> element, so it cannot be edited.");
        _cachedStyles = LoadStyles();
        _generatedAbstractNumberingIds = new Dictionary<(bool Ordered, int Start), int>();
    }

    /// <summary>
    /// Returns the required main document part. The SDK returns null when the package has no
    /// officeDocument relationship, and throws <see cref="InvalidOperationException"/> when that
    /// relationship targets a part that does not exist in the package. Both structural failures
    /// surface as the DOCX domain exception instead of an NRE or an unrelated SDK failure.
    /// </summary>
    private static MainDocumentPart GetRequiredMainPart(WordprocessingDocument document)
    {
        try
        {
            return document.MainDocumentPart
                ?? throw new OfficeEditorException(
                    "The package does not contain a WordprocessingML main document part, so it is not a valid DOCX document.");
        }
        catch (InvalidOperationException ex) when (ex is not ObjectDisposedException)
        {
            throw new OfficeEditorException(
                "The package's main document relationship targets a part that does not exist, so it is not a valid DOCX document.",
                ex);
        }
    }

    /// <summary>
    /// Creates a new DOCX document at the given path and returns a builder over it.
    /// The path must be a non-empty, non-whitespace file path, mirroring
    /// <see cref="Open(string)"/>; null, empty, or whitespace paths are rejected up front as
    /// argument exceptions. If opening or initializing the package fails, the opened document
    /// is disposed so no file handle leaks, and the original failure propagates unchanged —
    /// ordinary path/permission and IO errors are never normalized into the domain exception.
    /// On success the returned builder owns the document; callers must dispose it.
    /// </summary>
    public static IDocumentBuilder Create(string path)
    {
        ArgumentNullException.ThrowIfNull(path);
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("Path must not be empty or whitespace.", nameof(path));
        }

        WordprocessingDocument? document = null;
        try
        {
            document = WordprocessingDocument.Create(path, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = new Body();
            mainPart.Document.Append(body);
            mainPart.Document.Save();

            return new DocumentBuilder(document, true, path);
        }
        catch
        {
            // A failed create must never leak the partially opened package handle; the
            // caller-visible failure is preserved and rethrown unchanged, and teardown is
            // best-effort so it can never mask that primary failure.
            TryDispose(document);
            throw;
        }
    }

    public static IDocumentBuilder Create()
    {
        var memoryStream = new MemoryStream();
        var document = WordprocessingDocument.Create(memoryStream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new Document();
        var body = new Body();
        mainPart.Document.Append(body);
        mainPart.Document.Save();

        return new DocumentBuilder(document, true, null, memoryStream);
    }

    public static IDocumentBuilder Open(string path)
    {
        ArgumentNullException.ThrowIfNull(path);
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("Path must not be empty or whitespace.", nameof(path));
        }

        WordprocessingDocument? document = null;
        try
        {
            document = WordprocessingDocument.Open(path, true);
            return new DocumentBuilder(document, false, path);
        }
        catch (Exception ex)
        {
            TryDispose(document);
            if (IsMalformedPackageFailure(ex))
            {
                throw new OfficeEditorException(
                    $"Could not open a valid DOCX document from the file '{path}'. The content is missing, corrupt, or not a WordprocessingML package.",
                    ex);
            }
            throw;
        }
    }

    /// <summary>
    /// Opens an existing DOCX document from a stream.
    /// The stream content is copied to an internal buffer; the caller retains ownership of the original stream.
    /// </summary>
    public static IDocumentBuilder Open(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);

        MemoryStream? buffer = null;
        WordprocessingDocument? document = null;
        try
        {
            buffer = new MemoryStream();
            stream.CopyTo(buffer);
            buffer.Position = 0;
            document = WordprocessingDocument.Open(buffer, true);
            return new DocumentBuilder(document, false, null, buffer);
        }
        catch (Exception ex)
        {
            TryDispose(document);
            TryDispose(buffer);
            if (IsMalformedPackageFailure(ex))
            {
                throw new OfficeEditorException(
                    $"Could not open a valid DOCX document from the supplied stream. The content is missing, corrupt, or not a WordprocessingML package.",
                    ex);
            }
            throw;
        }
    }

    /// <summary>
    /// Opens an existing DOCX document from a byte array (convenience overload).
    /// </summary>
    public static IDocumentBuilder Open(byte[] bytes)
    {
        ArgumentNullException.ThrowIfNull(bytes);

        MemoryStream? buffer = null;
        WordprocessingDocument? document = null;
        try
        {
            buffer = new MemoryStream(bytes.Length);
            buffer.Write(bytes, 0, bytes.Length);
            buffer.Position = 0;
            document = WordprocessingDocument.Open(buffer, true);
            return new DocumentBuilder(document, false, null, buffer);
        }
        catch (Exception ex)
        {
            TryDispose(document);
            TryDispose(buffer);
            if (IsMalformedPackageFailure(ex))
            {
                throw new OfficeEditorException(
                    $"Could not open a valid DOCX document from the supplied byte array. The content is missing, corrupt, or not a WordprocessingML package.",
                    ex);
            }
            throw;
        }
    }

    /// <summary>
    /// True for the failure types the OpenXml SDK raises for structurally corrupt, wrong-format,
    /// or malformed-XML packages. Only these are normalized to <see cref="OfficeEditorException"/>;
    /// everything else — <see cref="FileNotFoundException"/> and ordinary path/permission errors,
    /// the constructor's own <see cref="OfficeEditorException"/>, and programmer/fatal failures —
    /// must propagate unchanged after the opened document and internal buffer are cleaned up.
    /// </summary>
    private static bool IsMalformedPackageFailure(Exception failure)
    {
        return failure is OpenXmlPackageException
            or FileFormatException
            or InvalidDataException
            or XmlException;
    }

    /// <summary>
    /// Best-effort teardown of a partially opened package or internal buffer after an open/create
    /// failure. Teardown is never allowed to replace the primary open/validation exception: a
    /// dispose failure here is deliberately swallowed (the GC reclaims the handle) so the original
    /// failure — normalized or raw — propagates unchanged and the caller-stream ownership contract
    /// is unaffected.
    /// </summary>
    private static void TryDispose(IDisposable? disposable)
    {
        if (disposable is null)
        {
            return;
        }

        try
        {
            disposable.Dispose();
        }
        catch
        {
            // Swallowed: a teardown failure must not mask the primary open/validation exception.
        }
    }

    /// <summary>
    /// Best-effort removal of a record's temporary output after a failed write. A deletion failure
    /// is deliberately swallowed so it can never mask the record's primary failure; the OS temp
    /// cleaner is the backstop for any survivor.
    /// </summary>
    private static void TryDeleteFile(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
            // Swallowed: best-effort cleanup must not mask the primary failure.
        }
    }

    public IDocumentBuilder AddParagraph(string text, string? style = null)
    {
        var paragraph = new Paragraph();
        var run = new Run(new Text(text));
        paragraph.Append(run);

        if (!string.IsNullOrEmpty(style))
        {
            EnsureStyle(style);
            paragraph.ParagraphProperties = new ParagraphProperties(
                new ParagraphStyleId { Val = style }
            );
        }

        _body.Append(paragraph);
        return this;
    }

    public IDocumentBuilder InsertAfter(string targetText, string text, string? style = null)
    {
        var targetParagraph = FindParagraphByText(targetText);
        if (targetParagraph == null)
        {
            throw new InvalidOperationException($"Paragraph containing '{targetText}' not found.");
        }

        var newParagraph = CreateParagraph(text, style);
        _body.InsertAfter(newParagraph, targetParagraph);
        return this;
    }

    public IDocumentBuilder InsertBefore(string targetText, string text, string? style = null)
    {
        var targetParagraph = FindParagraphByText(targetText);
        if (targetParagraph == null)
        {
            throw new InvalidOperationException($"Paragraph containing '{targetText}' not found.");
        }

        var newParagraph = CreateParagraph(text, style);
        _body.InsertBefore(newParagraph, targetParagraph);
        return this;
    }

    public IDocumentBuilder ReplaceText(string find, string replace)
    {
        // Empty find would insert the replacement between every character of the document.
        if (string.IsNullOrEmpty(find))
        {
            throw new ArgumentException("Find must be a non-empty string.", nameof(find));
        }

        // Descendants (not Elements) covers table cells, nested runs, and any other
        // container; header/footer parts are included so templated letterheads work.
        foreach (var text in EnumerateReplaceableTexts())
        {
            if (text.Text.Contains(find))
            {
                text.Text = text.Text.Replace(find, replace);
            }
        }
        return this;
    }

    private IEnumerable<Text> EnumerateReplaceableTexts()
    {
        foreach (var text in _body.Descendants<Text>())
        {
            yield return text;
        }

        var mainPart = _document.MainDocumentPart!;
        foreach (var headerPart in mainPart.HeaderParts)
        {
            if (headerPart.Header == null)
            {
                continue;
            }

            foreach (var text in headerPart.Header.Descendants<Text>())
            {
                yield return text;
            }
        }

        foreach (var footerPart in mainPart.FooterParts)
        {
            if (footerPart.Footer == null)
            {
                continue;
            }

            foreach (var text in footerPart.Footer.Descendants<Text>())
            {
                yield return text;
            }
        }
    }

    public IDocumentBuilder ReplaceParagraph(string targetText, string newText, string? style = null)
    {
        var targetParagraph = FindParagraphByText(targetText);
        if (targetParagraph == null)
        {
            throw new InvalidOperationException($"Paragraph containing '{targetText}' not found.");
        }

        // Preserve existing style if no new style specified
        var existingStyle = targetParagraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        
        targetParagraph.RemoveAllChildren<Run>();
        var run = new Run(new Text(newText));
        targetParagraph.Append(run);

        if (!string.IsNullOrEmpty(style))
        {
            EnsureStyle(style);
            targetParagraph.ParagraphProperties ??= new ParagraphProperties();
            targetParagraph.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId { Val = style };
        }
        else if (!string.IsNullOrEmpty(existingStyle))
        {
            // Preserve original style
            targetParagraph.ParagraphProperties ??= new ParagraphProperties();
            targetParagraph.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId { Val = existingStyle };
        }

        return this;
    }

    public IDocumentBuilder DeleteParagraph(string targetText)
    {
        var targetParagraph = FindParagraphByText(targetText);
        if (targetParagraph != null)
        {
            targetParagraph.Remove();
        }
        return this;
    }

    public IDocumentBuilder ApplyStyle(string styleId)
    {
        EnsureStyle(styleId);

        // Apply to last paragraph if no specific target
        var lastParagraph = _body.Elements<Paragraph>().LastOrDefault();
        if (lastParagraph != null)
        {
            lastParagraph.ParagraphProperties ??= new ParagraphProperties();
            lastParagraph.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId { Val = styleId };
        }

        return this;
    }

    public void Save(string? path = null)
    {
        // Reject empty/whitespace destinations before flushing so an invalid path fails fast
        // without touching the source document; matches the Create/Open path contract.
        if (path is not null && string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("Path must not be empty or whitespace.", nameof(path));
        }

        _document.Save();

        if (!string.IsNullOrEmpty(path) && !string.Equals(_filePath, path, StringComparison.OrdinalIgnoreCase))
        {
            using var clone = _document.Clone(path);
        }
        else if (string.IsNullOrEmpty(path) && string.IsNullOrEmpty(_filePath))
        {
            throw new InvalidOperationException(
                "No file path is associated with this document. Use Save(string path), Save(Stream stream), or SaveToBytes() to specify a destination.");
        }
    }

    /// <summary>
    /// Writes the current document content to the provided stream and leaves it open.
    /// </summary>
    public void Save(Stream stream)
    {
        _document.Save();

        if (_documentStream != null)
        {
            _documentStream.Position = 0;
            _documentStream.CopyTo(stream);
            _documentStream.Position = 0;
        }
        else
        {
            using var fileStream = File.OpenRead(_filePath!);
            fileStream.CopyTo(stream);
        }
    }

    /// <summary>
    /// Returns the current document content as a byte array.
    /// </summary>
    public byte[] SaveToBytes()
    {
        _document.Save();

        if (_documentStream != null)
        {
            return _documentStream.ToArray();
        }

        return File.ReadAllBytes(_filePath!);
    }

    private Paragraph CreateParagraph(string text, string? style)
    {
        var paragraph = new Paragraph();
        var run = new Run(new Text(text));
        paragraph.Append(run);

        if (!string.IsNullOrEmpty(style))
        {
            EnsureStyle(style);
            paragraph.ParagraphProperties = new ParagraphProperties(
                new ParagraphStyleId { Val = style }
            );
        }

        return paragraph;
    }

    /// <summary>
    /// Locates the first paragraph whose text contains <paramref name="text"/>. Because every
    /// paragraph's text contains the empty string, an empty target would silently select the
    /// first paragraph of the document — deleting, replacing, or inserting around the wrong
    /// element — so null and empty targets are rejected as argument errors at this choke point
    /// shared by all find-by-text operations.
    /// </summary>
    private Paragraph? FindParagraphByText(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        if (string.IsNullOrEmpty(text))
        {
            throw new ArgumentException("Target text must be a non-empty string.", nameof(text));
        }

        return _body.Elements<Paragraph>()
            .FirstOrDefault(p => p.InnerText.Contains(text));
    }

    private Dictionary<string, Style> LoadStyles()
    {
        var styles = new Dictionary<string, Style>();
        var stylesPart = _document.MainDocumentPart?.StyleDefinitionsPart;
        
        if (stylesPart?.Styles != null)
        {
            foreach (var style in stylesPart.Styles.Elements<Style>())
            {
                if (style.StyleId?.Value != null)
                {
                    styles[style.StyleId.Value] = style;
                }
            }
        }

        return styles;
    }

    private void EnsureStyle(string styleId)
    {
        if (_cachedStyles.ContainsKey(styleId))
            return;

        var stylesPart = _document.MainDocumentPart!.StyleDefinitionsPart;
        if (stylesPart == null)
        {
            stylesPart = _document.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles();
        }

        var style = CreateDefaultStyle(styleId);
        stylesPart.Styles ??= new Styles();
        stylesPart.Styles.Append(style);
        _cachedStyles[styleId] = style;
    }

    /// <summary>
    /// Allocates a fresh numbering instance for a single list and returns its id. Each list
    /// gets its own instance so its counter restarts at 1 and bullets can never bind to a
    /// decimal (or vice versa) definition from the host document. The abstract numbering
    /// definition behind the instance is reused across lists with matching semantics: one
    /// generated definition per (ordered, start) kind is created on demand and shared by
    /// every later list of the same kind instead of duplicating identical definitions.
    /// Existing numbering in the document is never touched or reused (its formatting and
    /// levels may differ); ids are allocated as max(existing) + 1 so they can never collide,
    /// and the abstractNum is inserted before the first numbering instance because
    /// CT_Numbering requires all abstractNum elements to precede all num elements.
    /// </summary>
    private int AllocateNumberingInstance(bool ordered) => AllocateNumberingInstance(ordered, 1);

    private int AllocateNumberingInstance(bool ordered, int start)
    {
        var mainPart = _document.MainDocumentPart!;
        var numberingPart = mainPart.NumberingDefinitionsPart;
        if (numberingPart == null)
        {
            numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
        }

        var numbering = numberingPart.Numbering;
        if (numbering == null)
        {
            numbering = new Numbering();
            numberingPart.Numbering = numbering;
        }

        // Reuse the abstract definition this builder generated for the same list semantics
        // (ordered vs bullet, with the same start value); create a fresh one only on first
        // use of that kind.
        var key = (ordered, start);
        if (!_generatedAbstractNumberingIds.TryGetValue(key, out int abstractNumId))
        {
            abstractNumId = numbering.Elements<AbstractNum>()
                .Select(n => n.AbstractNumberId?.Value ?? -1)
                .DefaultIfEmpty(-1)
                .Max() + 1;

            var abstractNum = ordered
                ? CreateOrderedAbstractNum(abstractNumId, start)
                : CreateBulletAbstractNum(abstractNumId);

            var firstInstance = numbering.Elements<NumberingInstance>().FirstOrDefault();
            if (firstInstance != null)
            {
                numbering.InsertBefore(abstractNum, firstInstance);
            }
            else
            {
                numbering.Append(abstractNum);
            }

            _generatedAbstractNumberingIds[key] = abstractNumId;
        }

        int numberingId = numbering.Elements<NumberingInstance>()
            .Select(n => n.NumberID?.Value ?? 0)
            .DefaultIfEmpty(0)
            .Max() + 1;

        numbering.Append(new NumberingInstance(
            new AbstractNumId { Val = abstractNumId }
        ) { NumberID = numberingId });

        return numberingId;
    }

    private static AbstractNum CreateOrderedAbstractNum(int abstractNumId, int start)
    {
        return new AbstractNum(
            new Level(
                new StartNumberingValue { Val = start },
                new NumberingFormat { Val = NumberFormatValues.Decimal },
                new LevelText { Val = "%1." },
                new LevelJustification { Val = LevelJustificationValues.Left },
                new PreviousParagraphProperties(
                    new Indentation { Left = "720", Hanging = "360" }
                )
            ) { LevelIndex = 0 }
        ) { AbstractNumberId = abstractNumId };
    }

    private static AbstractNum CreateBulletAbstractNum(int abstractNumId)
    {
        return new AbstractNum(
            new Level(
                new StartNumberingValue { Val = 1 },
                new NumberingFormat { Val = NumberFormatValues.Bullet },
                new LevelText { Val = "\u2022" },
                new LevelJustification { Val = LevelJustificationValues.Left },
                new PreviousParagraphProperties(
                    new Indentation { Left = "720", Hanging = "360" }
                )
            ) { LevelIndex = 0 }
        ) { AbstractNumberId = abstractNumId };
    }

    private static void ValidateContentBlocks(List<ContentBlock> blocks)
    {
        ArgumentNullException.ThrowIfNull(blocks);
        for (int i = 0; i < blocks.Count; i++)
        {
            // A null entry would be silently skipped by the renderer's type switch, claiming
            // success while rendering fewer blocks than requested.
            if (blocks[i] is null)
            {
                throw new ArgumentException($"blocks[{i}]: each content block must be non-null.", nameof(blocks));
            }
        }
    }

    private static Style CreateDefaultStyle(string styleId)
    {
        return styleId switch
        {
            "Heading1" => CreateHeadingStyle(1, "heading 1", "32", 0),
            "Heading2" => CreateHeadingStyle(2, "heading 2", "26", 1),
            "Heading3" => CreateHeadingStyle(3, "heading 3", "24", 2),
            "Heading4" => CreateHeadingStyle(4, "heading 4", "22", 3),
            "Heading5" => CreateHeadingStyle(5, "heading 5", "20", 4),
            "Heading6" => CreateHeadingStyle(6, "heading 6", "20", 5),
            "Normal" => new Style(
                new StyleName { Val = "Normal" }
            ) { Type = StyleValues.Paragraph, StyleId = styleId, Default = true },
            "Quote" => new Style(
                new StyleName { Val = "Quote" },
                new BasedOn { Val = "Normal" },
                new StyleParagraphProperties(
                    new Indentation { Left = "720" }
                ),
                new StyleRunProperties(
                    new Italic()
                )
            ) { Type = StyleValues.Paragraph, StyleId = styleId },
            "Code" => new Style(
                new StyleName { Val = "Code" },
                new BasedOn { Val = "Normal" },
                new StyleRunProperties(
                    new RunFonts { Ascii = "Consolas", HighAnsi = "Consolas" }
                )
            ) { Type = StyleValues.Paragraph, StyleId = styleId },
            "Hyperlink" => new Style(
                new StyleName { Val = "Hyperlink" },
                new StyleRunProperties(
                    new Color { Val = "0563C1", ThemeColor = ThemeColorValues.Hyperlink },
                    new Underline { Val = UnderlineValues.Single }
                )
            ) { Type = StyleValues.Character, StyleId = styleId },
            _ => new Style(
                new StyleName { Val = styleId }
            ) { Type = StyleValues.Paragraph, StyleId = styleId }
        };
    }

    private static Style CreateHeadingStyle(int level, string name, string fontSize, int outlineLevel)
    {
        var before = (240 - (level - 1) * 30).ToString();
        return new Style(
            new StyleName { Val = name },
            new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new KeepNext(),
                new SpacingBetweenLines { Before = before, After = "60" },
                new OutlineLevel { Val = outlineLevel }
            ),
            new StyleRunProperties(
                new Bold(),
                new BoldComplexScript(),
                new FontSize { Val = fontSize },
                new FontSizeComplexScript { Val = fontSize }
            )
        ) { Type = StyleValues.Paragraph, StyleId = $"Heading{level}" };
    }

    public IDocumentBuilder AddRichContent(List<ContentBlock> blocks)
    {
        ValidateContentBlocks(blocks);

        // Numbering definitions are allocated lazily per list block via the callback,
        // so documents without lists never get a numbering part.
        var renderer = new ContentBlockRenderer(StyleMapping.Default, _cachedStyles, EnsureStyle, AllocateNumberingInstance);
        renderer.Render(_body, blocks);
        return this;
    }

    public IDocumentBuilder ReplaceWithRichContent(string targetText, List<ContentBlock> blocks)
    {
        ValidateContentBlocks(blocks);

        var targetParagraph = FindParagraphByText(targetText);
        if (targetParagraph == null)
        {
            throw new InvalidOperationException($"Paragraph containing '{targetText}' not found.");
        }

        var renderer = new ContentBlockRenderer(StyleMapping.Default, _cachedStyles, EnsureStyle, AllocateNumberingInstance);
        
        // Remove target paragraph and insert rich content before its position
        var parent = targetParagraph.Parent;
        if (parent != null)
        {
            // Create a temporary body to render blocks
            var tempBody = new Body();
            renderer.Render(tempBody, blocks);
            
            // Insert after the last inserted node so block order is preserved.
            OpenXmlElement insertAfter = targetParagraph;
            foreach (var element in tempBody.Elements().ToList())
            {
                insertAfter = parent.InsertAfter(element.CloneNode(true), insertAfter)!;
            }
            
            targetParagraph.Remove();
        }

        return this;
    }

    public IDocumentBuilder AddHyperlink(string url, string displayText, string? style = null)
    {
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri))
        {
            throw new OfficeEditorException(
                $"Invalid hyperlink URL '{url}'. Hyperlink URLs must be absolute (e.g. https://example.com/page).");
        }

        EnsureStyle("Hyperlink");

        var mainPart = _document.MainDocumentPart!;
        var hyperlinkRelationship = mainPart.AddHyperlinkRelationship(uri, true);
        var paragraph = new Paragraph();
        var hyperlink = new Hyperlink() { History = true, Id = hyperlinkRelationship.Id };
        var run = new Run(
            new RunProperties(new RunStyle { Val = "Hyperlink" }),
            new Text(displayText));
        hyperlink.Append(run);
        paragraph.Append(hyperlink);

        if (!string.IsNullOrEmpty(style))
        {
            EnsureStyle(style);
            paragraph.ParagraphProperties = new ParagraphProperties(
                new ParagraphStyleId { Val = style }
            );
        }

        _body.Append(paragraph);
        return this;
    }

    public IDocumentBuilder AddMarkdown(string markdown, StyleMapping? styleMap = null)
    {
        // Routed through the rich recursive renderer: markdown parses into the structured
        // Markdown IR and renders directly to OOXML (headings with inline formatting, nested
        // lists, tables, footnotes, images, hyperlinks) without the lossy ContentBlock pass.
        return AddRichMarkdown(markdown, new MarkdownRenderOptions
        {
            StyleMapping = styleMap ?? StyleMapping.Default
        });
    }

    public IDocumentBuilder ReplaceWithMarkdown(string targetText, string markdown, StyleMapping? styleMap = null)
    {
        return ReplaceWithRichMarkdown(targetText, markdown, new MarkdownRenderOptions
        {
            StyleMapping = styleMap ?? StyleMapping.Default
        });
    }

    /// <summary>
    /// Appends rich markdown to the document with the given rendering options. The last
    /// conversion outcome (parse + render diagnostics) is exposed via
    /// <see cref="LastRichMarkdownResult"/>.
    /// </summary>
    public IDocumentBuilder AddRichMarkdown(string markdown, MarkdownRenderOptions? options = null)
    {
        RenderRichCore(markdown, options, _body);
        return this;
    }

    /// <summary>
    /// Replaces the paragraph containing <paramref name="targetText"/> with rich markdown.
    /// Content renders into a detached body so relationships (hyperlinks, images, the
    /// footnotes part) are created against the live main part — where every <c>r:id</c>
    /// resolves — and the resulting elements are cloned into the target's place, keeping the
    /// caller's package relationships valid.
    /// </summary>
    public IDocumentBuilder ReplaceWithRichMarkdown(string targetText, string markdown, MarkdownRenderOptions? options = null)
    {
        // Resolve the target before rendering so an empty/missing target fails fast without
        // touching the document (the empty target is rejected by FindParagraphByText).
        var targetParagraph = FindParagraphByText(targetText);
        if (targetParagraph == null)
        {
            throw new InvalidOperationException($"Paragraph containing '{targetText}' not found.");
        }

        var outcome = RenderRichCore(markdown, options, new Body());

        var parent = targetParagraph.Parent;
        if (parent != null)
        {
            // Insert after the last inserted node so block order is preserved, then drop
            // the target. Cloning preserves the relationship ids created during render.
            OpenXmlElement insertAfter = targetParagraph;
            foreach (var element in outcome.Elements)
            {
                insertAfter = parent.InsertAfter(element.CloneNode(true), insertAfter)!;
            }
            targetParagraph.Remove();
        }

        return this;
    }

    public MarkdownRenderResult? LastRichMarkdownResult { get; private set; }

    private sealed record RichRenderOutcome(IReadOnlyList<OpenXmlElement> Elements, MarkdownRenderResult Result);

    private RichRenderOutcome RenderRichCore(string markdown, MarkdownRenderOptions? options, Body targetBody)
    {
        ArgumentNullException.ThrowIfNull(markdown);

        var renderOptions = options ?? MarkdownRenderOptions.Default;
        var parser = new Markdown.RichMarkdownParser(renderOptions.ParseOptions ?? MarkdownParseOptions.Default);
        var parseResult = parser.Parse(markdown);

        var renderer = new RichMarkdownRenderer(
            _document.MainDocumentPart!,
            renderOptions,
            (ordered, start) => AllocateNumberingInstance(ordered, start));
        var elements = renderer.Render(targetBody, parseResult.Document);

        // The renderer may have appended generated fallback styles to the styles part so every
        // referenced style exists. Merge them into the style cache so a later EnsureStyle call
        // reuses those styles instead of appending a duplicate style id.
        RefreshCachedStyles();

        var result = new MarkdownRenderResult(
            parseResult,
            parseResult.Diagnostics.Concat(renderer.Diagnostics).ToArray());
        LastRichMarkdownResult = result;
        return new RichRenderOutcome(elements, result);
    }

    private void RefreshCachedStyles()
    {
        var stylesPart = _document.MainDocumentPart?.StyleDefinitionsPart;
        if (stylesPart?.Styles is not { } styles)
        {
            return;
        }

        foreach (var style in styles.Elements<Style>())
        {
            if (style.StyleId?.Value is { } id)
            {
                _cachedStyles[id] = style;
            }
        }
    }

    public List<VariableInfo> DetectVariables()
    {
        var detector = new Variables.DocxVariableDetector();
        return detector.Scan(_document);
    }

    public IDocumentBuilder MergeVariables(Dictionary<string, string> data)
    {
        ArgumentNullException.ThrowIfNull(data);

        var replacer = new Variables.DocxVariableReplacer();
        replacer.Replace(_document, data);
        return this;
    }

    /// <summary>
    /// Publipostage: generates one document per record by merging the record's variables into the
    /// template. Every record is produced with per-record atomicity: it is fully written to a
    /// temporary file in the destination directory, opened/edited/saved/closed there, and only then
    /// atomically moved onto the final output path. A record that fails leaves no corrupt partial
    /// output and no temp artifact behind, and any pre-existing final file for that record survives;
    /// records already completed before the failure remain on disk (this is per-record, not
    /// whole-batch, atomicity). All records are prevalidated — non-null records, keys, and values —
    /// before any output is written so a malformed batch fails before touching disk.
    /// </summary>
    public void MergeBatch(List<Dictionary<string, string>> records, string outputPattern, string? templatePath = null)
    {
        ArgumentNullException.ThrowIfNull(records);
        ArgumentNullException.ThrowIfNull(outputPattern);
        if (string.IsNullOrWhiteSpace(outputPattern))
        {
            throw new ArgumentException("Output pattern must not be empty or whitespace.", nameof(outputPattern));
        }
        for (int i = 0; i < records.Count; i++)
        {
            if (records[i] is null)
            {
                throw new ArgumentException($"records[{i}]: each record must be non-null.", nameof(records));
            }

            foreach (var entry in records[i])
            {
                // Dictionary<string,string> permits null values (and rejects null keys), but both
                // are rejected here so a malformed record fails up front instead of silently
                // dropping the variable from the generated file name.
                if (entry.Key is null || entry.Value is null)
                {
                    throw new ArgumentException(
                        $"records[{i}]: each record key and value must be non-null.", nameof(records));
                }
            }
        }

        if (templatePath is not null && string.IsNullOrWhiteSpace(templatePath))
        {
            throw new ArgumentException("Template path must not be empty or whitespace.", nameof(templatePath));
        }

        // Use provided template path or try to get from document
        var originalPath = templatePath ?? GetDocumentPath();
        
        if (string.IsNullOrEmpty(originalPath))
        {
            throw new InvalidOperationException("Document path not available. Please provide templatePath parameter.");
        }

        if (records.Count > 0)
        {
            // Route the template through the hardened open contract before any output file is
            // written. A corrupt, wrong-format, or structurally-invalid template previously bypassed
            // the Open boundary: it was copied to every output path and then failed mid-loop with a
            // raw SDK exception, leaving corrupt partial outputs on disk. Validating first makes the
            // malformed template fail up front with the domain exception (missing-file and other
            // ordinary IO errors still propagate unchanged) while no output file has been created.
            using var template = Open(File.ReadAllBytes(originalPath));
        }
        
        for (int i = 0; i < records.Count; i++)
        {
            // Resolve the final output path for this record from the pattern.
            var outputPath = outputPattern.Replace("{index}", i.ToString());
            foreach (var kvp in records[i])
            {
                outputPath = outputPath.Replace($"{{{kvp.Key}}}", kvp.Value);
            }

            // The temp file lives in the destination directory so the final move is a same-volume
            // rename (atomic); the final output is never touched until it is atomically replaced.
            var destinationDirectory = Path.GetDirectoryName(outputPath) ?? string.Empty;
            var tempPath = Path.Combine(
                destinationDirectory,
                $"{Path.GetFileName(outputPath)}.{Guid.NewGuid():N}.tmp");

            try
            {
                File.Copy(originalPath, tempPath, true);
                using (var doc = WordprocessingDocument.Open(tempPath, true))
                {
                    var replacer = new Variables.DocxVariableReplacer();
                    replacer.Replace(doc, records[i]);
                    doc.Save();
                }
                File.Move(tempPath, outputPath, overwrite: true);
            }
            finally
            {
                // A failed record must not leave its partial temp file behind; a pre-existing final
                // output is preserved because it is only replaced by the atomic move above.
                TryDeleteFile(tempPath);
            }
        }
    }

    private string? GetDocumentPath() => _filePath;

    public void Dispose()
    {
        _document.Dispose();
        _documentStream?.Dispose();
    }
}
