using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace DocxEditor.Tests;

/// <summary>
/// Asserts that a produced OOXML package passes the SDK's <see cref="OpenXmlValidator"/>.
/// This catches structural/schema defects — the class of bug that triggers Excel's
/// "we found a problem with some content" repair prompt — at test time instead of in Excel.
/// </summary>
public static class OpenXmlAssert
{
    private const int MaxErrorsReported = 10;

    /// <summary>Asserts the document produces zero OpenXmlValidator errors.</summary>
    public static void NoValidationErrors(SpreadsheetDocument document)
    {
        var errors = new OpenXmlValidator().Validate(document).ToList();
        Assert.True(errors.Count == 0, FormatErrors(errors));
    }

    /// <summary>Asserts the document produces zero OpenXmlValidator errors.</summary>
    public static void NoValidationErrors(WordprocessingDocument document)
    {
        var errors = new OpenXmlValidator().Validate(document).ToList();
        Assert.True(errors.Count == 0, FormatErrors(errors));
    }

    /// <summary>Opens the file read-only and asserts it produces zero OpenXmlValidator errors.</summary>
    public static void NoValidationErrors(string path)
    {
        using var document = SpreadsheetDocument.Open(path, false);
        NoValidationErrors(document);
    }

    /// <summary>Opens the DOCX file read-only and asserts it produces zero OpenXmlValidator errors.</summary>
    public static void NoDocxValidationErrors(string path)
    {
        using var document = WordprocessingDocument.Open(path, false);
        NoValidationErrors(document);
    }

    private static string FormatErrors(List<ValidationErrorInfo> errors)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"Expected no OOXML validation errors, but found {errors.Count}:");
        foreach (var error in errors.Take(MaxErrorsReported))
        {
            sb.AppendLine($"  - [{error.ErrorType}] {error.Description} (path: {error.Path?.XPath})");
        }

        if (errors.Count > MaxErrorsReported)
        {
            sb.AppendLine($"  … and {errors.Count - MaxErrorsReported} more.");
        }

        return sb.ToString();
    }
}
