using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxEditor.Core.Builders;

public interface IDocumentBuilder
{
    IDocumentBuilder AddParagraph(string text, string? style = null);
    IDocumentBuilder InsertAfter(string targetText, string text, string? style = null);
    IDocumentBuilder ReplaceText(string find, string replace);
    IDocumentBuilder ReplaceParagraph(string targetText, string newText, string? style = null);
    IDocumentBuilder DeleteParagraph(string targetText);
    void Save(string? path = null);
}

public class DocumentBuilder : IDocumentBuilder
{
    private readonly WordprocessingDocument _document;
    private readonly Body _body;
    private bool _isNewDocument;

    private DocumentBuilder(WordprocessingDocument document, bool isNew)
    {
        _document = document;
        _isNewDocument = isNew;
        _body = document.MainDocumentPart!.Document.Body!;
    }

    public static IDocumentBuilder Create(string path)
    {
        var document = WordprocessingDocument.Create(path, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new Document();
        var body = new Body();
        mainPart.Document.Append(body);
        
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

        var newParagraph = new Paragraph();
        var run = new Run(new Text(text));
        newParagraph.Append(run);

        if (!string.IsNullOrEmpty(style))
        {
            newParagraph.ParagraphProperties = new ParagraphProperties(
                new ParagraphStyleId { Val = style }
            );
        }

        _body.InsertAfter(newParagraph, targetParagraph);
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

        targetParagraph.RemoveAllChildren<Run>();
        var run = new Run(new Text(newText));
        targetParagraph.Append(run);

        if (!string.IsNullOrEmpty(style))
        {
            targetParagraph.ParagraphProperties ??= new ParagraphProperties();
            targetParagraph.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId { Val = style };
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

    public void Save(string? path = null)
    {
        if (!string.IsNullOrEmpty(path) && _isNewDocument)
        {
            // For new documents, we already created at the specified path
            _document.Save();
        }
        else
        {
            _document.Save();
        }
    }

    private Paragraph? FindParagraphByText(string text)
    {
        return _body.Elements<Paragraph>()
            .FirstOrDefault(p => p.InnerText.Contains(text));
    }

    public void Dispose()
    {
        _document.Dispose();
    }
}
