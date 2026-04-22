namespace OfficeEditor.Core.Exceptions;

public class OfficeEditorException : Exception
{
    public OfficeEditorException(string message) : base(message) { }
    public OfficeEditorException(string message, Exception innerException) : base(message, innerException) { }
}
