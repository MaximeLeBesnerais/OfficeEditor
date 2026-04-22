using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

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
    void Save(string? path = null);
}

public class DocumentBuilder : IDocumentBuilder
{
    private readonly WordprocessingDocument _document;
    private readonly Body _body;
    private readonly bool _isNewDocument;
    private readonly Dictionary<string, Style> _cachedStyles;

    private DocumentBuilder(WordprocessingDocument document, bool isNew)
    {
        _document = document;
        _isNewDocument = isNew;
        _body = document.MainDocumentPart!.Document.Body!;
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
        
        return new DocumentBuilder(document, true);
    }

    public static IDocumentBuilder Open(string path)
    {
        var document = WordprocessingDocument.Open(path, true);
        return new DocumentBuilder(document, false);
    }

    public IDocumentBuilder AddParagraph(string text, string? style = null)
    {
        var paragraph = new Paragraph();
        var run = new Run(new Text(text));
        paragraph.Append(run);

        if (!string.IsNullOrEmpty(style))
        {
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
        // Verify style exists in document
        if (!_cachedStyles.ContainsKey(styleId))
        {
            throw new InvalidOperationException($"Style '{styleId}' not found in document.");
        }

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
    }

    private Paragraph CreateParagraph(string text, string? style)
    {
        var paragraph = new Paragraph();
        var run = new Run(new Text(text));
        paragraph.Append(run);

        if (!string.IsNullOrEmpty(style))
        {
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

    public void Dispose()
    {
        _document.Dispose();
    }
}
