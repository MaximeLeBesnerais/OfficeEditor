using OfficeEditor.Core.Exceptions;

namespace XlsxEditor.Core.Exceptions;

public class XlsxException : OfficeEditorException
{
    public XlsxException(string message) : base(message) { }
    public XlsxException(string message, Exception innerException) : base(message, innerException) { }
}
