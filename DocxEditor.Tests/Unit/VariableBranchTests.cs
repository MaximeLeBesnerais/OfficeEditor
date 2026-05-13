using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Variables;
using XlsxEditor.Core.Variables;
using W = DocumentFormat.OpenXml.Wordprocessing;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace DocxEditor.Tests.Unit;

public class VariableBranchTests : IDisposable
{
    private readonly List<string> _files = new();

    [Fact]
    public void DocxDetector_ShouldScanBodyHeadersFootersAndDeduplicatePerLocation()
    {
        var path = CreateDocx(
            new W.Paragraph(new W.Run(new W.Text("Hello {{ name |Guest}} and {{name|Again}} {{empty|}} {{missing"))),
            headerText: "Header {{title}} {{title}}",
            footerText: "Footer {{name|FooterDefault}}");

        using var document = WordprocessingDocument.Open(path, false);
        var variables = new DocxVariableDetector().Scan(document);

        Assert.Equal(4, variables.Count);
        Assert.Contains(variables, v => v.Name == "name" && v.Location == "body" && v.DefaultValue == "Guest");
        Assert.Contains(variables, v => v.Name == "empty" && v.Location == "body" && v.DefaultValue == "");
        Assert.Contains(variables, v => v.Name == "title" && v.Location.StartsWith("header:"));
        Assert.Contains(variables, v => v.Name == "name" && v.Location.StartsWith("footer:") && v.DefaultValue == "FooterDefault");
        Assert.DoesNotContain(variables, v => v.FullMatch == "{{missing");
    }

    [Fact]
    public void DocxDetector_WithEmptyDocumentPart_ShouldReturnEmptyList()
    {
        var path = NewTempFile("docx");
        using (var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
        }

        using var reopened = WordprocessingDocument.Open(path, false);

        Assert.Empty(new DocxVariableDetector().Scan(reopened));
    }

    [Fact]
    public void DocxReplacer_ShouldReplaceSingleRunSplitRunsDefaultsAndPreserveMissingVariables()
    {
        var split = new W.Paragraph(
            new W.Run(new W.Text("Split {{fi")),
            new W.Run(new W.Text("rst}} and {{second|Default}} and {{empty|}} and {{missing}}")));
        var path = CreateDocx(new[]
            {
            new W.Paragraph(new W.Run(new W.Text("Hi {{ name }}; repeated {{name}}"))),
            split
            },
            headerText: "Header {{header|Fallback}}",
            footerText: "Footer {{footer}}");

        using (var document = WordprocessingDocument.Open(path, true))
        {
            new DocxVariableReplacer().Replace(document, new Dictionary<string, string>
            {
                ["name"] = "Ada",
                ["first"] = "One",
                ["footer"] = "Done"
            });
        }

        using var reopened = WordprocessingDocument.Open(path, false);
        var allText = reopened.MainDocumentPart!.Document.InnerText
            + string.Concat(reopened.MainDocumentPart.HeaderParts.Select(h => h.Header.InnerText))
            + string.Concat(reopened.MainDocumentPart.FooterParts.Select(f => f.Footer.InnerText));

        Assert.Contains("Hi Ada; repeated Ada", allText);
        Assert.Contains("Split One and Default and  and {{missing}}", allText);
        Assert.Contains("Header Fallback", allText);
        Assert.Contains("Footer Done", allText);

        var textNodes = reopened.MainDocumentPart.Document.Body!.Descendants<W.Text>().ToList();
        Assert.Contains(textNodes, t => t.Text.Contains("Split One") && t.Space?.Value == SpaceProcessingModeValues.Preserve);
    }

    [Fact]
    public void DocxReplacer_WithNoVariablesOrSingleTextRun_ShouldLeaveTextUnchanged()
    {
        var path = CreateDocx(new W.Paragraph(new W.Run(new W.Text("Plain text"))));

        using (var document = WordprocessingDocument.Open(path, true))
        {
            new DocxVariableReplacer().Replace(document, new Dictionary<string, string>());
        }

        using var reopened = WordprocessingDocument.Open(path, false);
        Assert.Equal("Plain text", reopened.MainDocumentPart!.Document.Body!.InnerText);
    }

    [Fact]
    public void XlsxDetector_ShouldScanSharedAndPlainCellsWithDefaultsAndDuplicateLocations()
    {
        var path = CreateXlsx(includeSharedStringPart: true,
            ("A1", CellValues.SharedString, "0"),
            ("B1", CellValues.String, "{{name}} {{name|Other}}"),
            ("C1", null, "{{empty|}} {{broken"),
            (null, CellValues.String, "{{noRef|Default}}"));

        using var document = SpreadsheetDocument.Open(path, false);
        var variables = new XlsxVariableDetector().Scan(document);

        Assert.Equal(4, variables.Count);
        Assert.Contains(variables, v => v.Name == "shared" && v.DefaultValue == "SharedDefault" && v.Location == "sheet:Sheet1:cell:A1");
        Assert.Single(variables.Where(v => v.Name == "name" && v.Location == "sheet:Sheet1:cell:B1"));
        Assert.Contains(variables, v => v.Name == "empty" && v.DefaultValue == "" && v.Location == "sheet:Sheet1:cell:C1");
        Assert.Contains(variables, v => v.Name == "noRef" && v.Location == "sheet:Sheet1:cell:Unknown");
    }

    [Fact]
    public void XlsxDetector_WithMissingWorkbookOrSheets_ShouldReturnEmptyList()
    {
        var noWorkbookPath = NewTempFile("xlsx");
        using (SpreadsheetDocument.Create(noWorkbookPath, SpreadsheetDocumentType.Workbook))
        {
        }

        using (var document = SpreadsheetDocument.Open(noWorkbookPath, false))
        {
            Assert.Empty(new XlsxVariableDetector().Scan(document));
        }

        var noSheetsPath = NewTempFile("xlsx");
        using (var document = SpreadsheetDocument.Create(noSheetsPath, SpreadsheetDocumentType.Workbook))
        {
            document.AddWorkbookPart().Workbook = new Workbook();
        }

        using (var document = SpreadsheetDocument.Open(noSheetsPath, false))
        {
            Assert.Empty(new XlsxVariableDetector().Scan(document));
        }
    }

    [Fact]
    public void XlsxReplacer_ShouldReplaceSharedStringsPlainCellsDefaultsAndMissingValues()
    {
        var path = CreateXlsx(includeSharedStringPart: true,
            ("A1", CellValues.SharedString, "0"),
            ("B1", CellValues.String, "{{provided}} {{defaulted|Fallback}} {{empty|}} {{missing}}"),
            ("C1", null, "plain"));

        using (var document = SpreadsheetDocument.Open(path, true))
        {
            new XlsxVariableReplacer().Replace(document, new Dictionary<string, string>
            {
                ["shared"] = "SharedValue",
                ["provided"] = "Given"
            });
        }

        using var reopened = SpreadsheetDocument.Open(path, false);
        var workbookPart = reopened.WorkbookPart!;
        var cells = workbookPart.WorksheetParts.First().Worksheet!.Descendants<Cell>().ToDictionary(c => c.CellReference!.Value!);

        Assert.Equal(CellValues.SharedString, cells["A1"].DataType!.Value);
        Assert.Equal("SharedValue", ReadCell(cells["A1"], workbookPart));
        Assert.Equal(CellValues.String, cells["B1"].DataType!.Value);
        Assert.Equal("Given Fallback {{empty|}} {{missing}}", cells["B1"].CellValue!.Text);
        Assert.Equal("plain", cells["C1"].CellValue!.Text);
    }

    [Fact]
    public void XlsxReplacer_WithExistingSharedStringResult_ShouldReuseExistingIndex()
    {
        var path = CreateXlsx(includeSharedStringPart: true,
            ("A1", CellValues.SharedString, "0"),
            ("B1", CellValues.SharedString, "1"));

        using (var document = SpreadsheetDocument.Open(path, true))
        {
            new XlsxVariableReplacer().Replace(document, new Dictionary<string, string> { ["shared"] = "Existing" });
        }

        using var reopened = SpreadsheetDocument.Open(path, false);
        var a1 = reopened.WorkbookPart!.WorksheetParts.First().Worksheet!.Descendants<Cell>().First(c => c.CellReference == "A1");
        Assert.Equal("1", a1.CellValue!.Text);
    }

    public void Dispose()
    {
        foreach (var file in _files)
        {
            if (File.Exists(file))
            {
                File.Delete(file);
            }
        }
    }

    private string CreateDocx(W.Paragraph bodyParagraph, string? headerText = null, string? footerText = null)
        => CreateDocx(new[] { bodyParagraph }, headerText, footerText);

    private string CreateDocx(IEnumerable<W.Paragraph> bodyParagraphs, string? headerText = null, string? footerText = null)
    {
        var path = NewTempFile("docx");
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body(bodyParagraphs.Cast<OpenXmlElement>()));

        if (headerText != null)
        {
            var headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new W.Header(new W.Paragraph(new W.Run(new W.Text(headerText))));
            headerPart.Header.Save();
        }

        if (footerText != null)
        {
            var footerPart = mainPart.AddNewPart<FooterPart>();
            footerPart.Footer = new W.Footer(new W.Paragraph(new W.Run(new W.Text(footerText))));
            footerPart.Footer.Save();
        }

        mainPart.Document.Save();
        return path;
    }

    private string CreateXlsx(bool includeSharedStringPart, params (string? Reference, CellValues? DataType, string Value)[] cells)
    {
        var path = NewTempFile("xlsx");
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook(new Sheets());

        if (includeSharedStringPart)
        {
            var shared = workbookPart.AddNewPart<SharedStringTablePart>();
            shared.SharedStringTable = new SharedStringTable(
                new SharedStringItem(new S.Text("{{shared|SharedDefault}}")),
                new SharedStringItem(new S.Text("Existing")));
            shared.SharedStringTable.Save();
        }

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var row = new Row { RowIndex = 1 };
        foreach (var spec in cells)
        {
            var cell = new Cell { CellValue = new CellValue(spec.Value) };
            if (spec.Reference != null)
            {
                cell.CellReference = spec.Reference;
            }
            if (spec.DataType.HasValue)
            {
                cell.DataType = spec.DataType.Value;
            }
            row.Append(cell);
        }

        worksheetPart.Worksheet = new Worksheet(new SheetData(row));
        worksheetPart.Worksheet.Save();
        var relationshipId = workbookPart.GetIdOfPart(worksheetPart);
        workbookPart.Workbook.Sheets!.Append(new Sheet { Id = relationshipId, SheetId = 1, Name = "Sheet1" });
        workbookPart.Workbook.Save();
        return path;
    }

    private string NewTempFile(string extension)
    {
        var path = Path.Combine(Path.GetTempPath(), $"branch_vars_{Guid.NewGuid():N}.{extension}");
        _files.Add(path);
        return path;
    }

    private static string ReadCell(Cell cell, WorkbookPart workbookPart)
    {
        if (cell.DataType?.Value == CellValues.SharedString)
        {
            var index = int.Parse(cell.CellValue!.Text);
            return workbookPart.SharedStringTablePart!.SharedStringTable.Elements<SharedStringItem>().ElementAt(index).InnerText;
        }

        return cell.CellValue?.Text ?? string.Empty;
    }
}
