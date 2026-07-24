using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Content;
using DocxEditor.Core.Models;
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
    private readonly string? _filePath;
    private readonly MemoryStream? _documentStream;

    private DocumentBuilder(WordprocessingDocument document, bool isNew, string? filePath, MemoryStream? documentStream = null)
    {
        _document = document;
        _isNewDocument = isNew;
        _filePath = filePath;
        _documentStream = documentStream;
        _body = document.MainDocumentPart!.Document!.Body!;
        _cachedStyles = LoadStyles();
    }

    public static IDocumentBuilder Create(string path)
    {
        var document = WordprocessingDocument.Create(path, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new Document();
        var body = new Body();
        mainPart.Document.Append(body);
        mainPart.Document.Save();
        
        return new DocumentBuilder(document, true, path);
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
        var document = WordprocessingDocument.Open(path, true);
        return new DocumentBuilder(document, false, path);
    }

    /// <summary>
    /// Opens an existing DOCX document from a stream.
    /// The stream content is copied to an internal buffer; the caller retains ownership of the original stream.
    /// </summary>
    public static IDocumentBuilder Open(Stream stream)
    {
        var memoryStream = new MemoryStream();
        stream.CopyTo(memoryStream);
        memoryStream.Position = 0;
        var document = WordprocessingDocument.Open(memoryStream, true);
        return new DocumentBuilder(document, false, null, memoryStream);
    }

    /// <summary>
    /// Opens an existing DOCX document from a byte array (convenience overload).
    /// </summary>
    public static IDocumentBuilder Open(byte[] bytes)
    {
        var memoryStream = new MemoryStream(bytes.Length);
        memoryStream.Write(bytes, 0, bytes.Length);
        memoryStream.Position = 0;
        var document = WordprocessingDocument.Open(memoryStream, true);
        return new DocumentBuilder(document, false, null, memoryStream);
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
        var paragraphs = _body.Elements<Paragraph>();
        foreach (var paragraph in paragraphs)
        {
            var runs = paragraph.Elements<Run>();
            foreach (var run in runs)
            {
                var texts = run.Elements<Text>();
                foreach (var text in texts)
                {
                    if (text.Text.Contains(find))
                    {
                        text.Text = text.Text.Replace(find, replace);
                    }
                }
            }
        }
        return this;
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

    private Paragraph? FindParagraphByText(string text)
    {
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

    private bool _numberingDefinitionsEnsured;

    private void EnsureNumberingDefinitions()
    {
        if (_numberingDefinitionsEnsured)
            return;

        var mainPart = _document.MainDocumentPart!;
        var numberingPart = mainPart.NumberingDefinitionsPart;
        if (numberingPart == null)
        {
            numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
        }

        var numbering = numberingPart.Numbering ?? new Numbering();

        if (!numbering.Elements<AbstractNum>().Any(n => n.AbstractNumberId?.Value == 1))
        {
            numbering.Append(CreateOrderedAbstractNum());
        }

        if (!numbering.Elements<AbstractNum>().Any(n => n.AbstractNumberId?.Value == 2))
        {
            numbering.Append(CreateBulletAbstractNum());
        }

        if (!numbering.Elements<NumberingInstance>().Any(n => n.NumberID?.Value == 1))
        {
            numbering.Append(new NumberingInstance(
                new AbstractNumId { Val = 1 }
            ) { NumberID = 1 });
        }

        if (!numbering.Elements<NumberingInstance>().Any(n => n.NumberID?.Value == 2))
        {
            numbering.Append(new NumberingInstance(
                new AbstractNumId { Val = 2 }
            ) { NumberID = 2 });
        }

        numberingPart.Numbering = numbering;
        _numberingDefinitionsEnsured = true;
    }

    private static AbstractNum CreateOrderedAbstractNum()
    {
        return new AbstractNum(
            new Level(
                new StartNumberingValue { Val = 1 },
                new NumberingFormat { Val = NumberFormatValues.Decimal },
                new LevelText { Val = "%1." },
                new LevelJustification { Val = LevelJustificationValues.Left },
                new PreviousParagraphProperties(
                    new Indentation { Left = "720", Hanging = "360" }
                )
            ) { LevelIndex = 0 }
        ) { AbstractNumberId = 1 };
    }

    private static AbstractNum CreateBulletAbstractNum()
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
        ) { AbstractNumberId = 2 };
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
        EnsureNumberingDefinitions();
        var renderer = new ContentBlockRenderer(StyleMapping.Default, _cachedStyles, EnsureStyle, EnsureNumberingDefinitions);
        renderer.Render(_body, blocks);
        return this;
    }

    public IDocumentBuilder ReplaceWithRichContent(string targetText, List<ContentBlock> blocks)
    {
        var targetParagraph = FindParagraphByText(targetText);
        if (targetParagraph == null)
        {
            throw new InvalidOperationException($"Paragraph containing '{targetText}' not found.");
        }

        EnsureNumberingDefinitions();
        var renderer = new ContentBlockRenderer(StyleMapping.Default, _cachedStyles, EnsureStyle, EnsureNumberingDefinitions);
        
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
        var mainPart = _document.MainDocumentPart!;
        var hyperlinkRelationship = mainPart.AddHyperlinkRelationship(new Uri(url), true);
        var paragraph = new Paragraph();
        var hyperlink = new Hyperlink() { History = true, Id = hyperlinkRelationship.Id };
        var run = new Run(new Text(displayText));
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
        var parser = new Markdown.MarkdownParser();
        var blocks = parser.Parse(markdown, styleMap);
        return AddRichContent(blocks);
    }

    public IDocumentBuilder ReplaceWithMarkdown(string targetText, string markdown, StyleMapping? styleMap = null)
    {
        var parser = new Markdown.MarkdownParser();
        var blocks = parser.Parse(markdown, styleMap);
        return ReplaceWithRichContent(targetText, blocks);
    }

    public List<VariableInfo> DetectVariables()
    {
        var detector = new Variables.DocxVariableDetector();
        return detector.Scan(_document);
    }

    public IDocumentBuilder MergeVariables(Dictionary<string, string> data)
    {
        var replacer = new Variables.DocxVariableReplacer();
        replacer.Replace(_document, data);
        return this;
    }

    public void MergeBatch(List<Dictionary<string, string>> records, string outputPattern, string? templatePath = null)
    {
        // Use provided template path or try to get from document
        var originalPath = templatePath ?? GetDocumentPath();
        
        if (string.IsNullOrEmpty(originalPath))
        {
            throw new InvalidOperationException("Document path not available. Please provide templatePath parameter.");
        }
        
        for (int i = 0; i < records.Count; i++)
        {
            // Create a copy of the original document
            var outputPath = outputPattern.Replace("{index}", i.ToString());
            foreach (var kvp in records[i])
            {
                outputPath = outputPath.Replace($"{{{kvp.Key}}}", kvp.Value);
            }
            
            File.Copy(originalPath, outputPath, true);
            
            using var doc = WordprocessingDocument.Open(outputPath, true);
            var replacer = new Variables.DocxVariableReplacer();
            replacer.Replace(doc, records[i]);
            doc.Save();
        }
    }

    private string? GetDocumentPath() => _filePath;

    public void Dispose()
    {
        _document.Dispose();
        _documentStream?.Dispose();
    }
}
