namespace DocxEditor.Core.Exceptions;

public class DocxEditorException : Exception
{
    public DocxEditorException(string message) : base(message) { }
    public DocxEditorException(string message, Exception innerException) : base(message, innerException) { }
}
