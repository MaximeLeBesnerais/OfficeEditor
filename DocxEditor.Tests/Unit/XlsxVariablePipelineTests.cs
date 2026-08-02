using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxEditor.Core.Builders;
using XlsxEditor.Core.Variables;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Exercises the variable pipeline (detector, replacer, template engine) against
/// <c>t="inlineStr"</c> cells, whose text lives in <c>&lt;is&gt;&lt;t&gt;</c> rather than
/// <c>&lt;v&gt;</c>. The pipeline must see those placeholders and write replacements back
/// into the inline string — never into a <c>t="str"</c> literal that would leave a stale
/// <c>&lt;is&gt;</c> beside the new value.
/// </summary>
public sealed class XlsxVariablePipelineTests : IDisposable
{
    private readonly List<string> _tempFiles = [];

    [Fact]
    public void Replacer_ReplacesPlaceholderInProperInlineStringCell_InPlace()
    {
        // Arrange: a well-formed inline-string cell (openpyxl/EPPlus-style)
        var path = CreateWorkbook(sd => AppendInlineStringCell(sd, "A1", "Hello {{name}}"));

        // Act
        using (var doc = SpreadsheetDocument.Open(path, true))
        {
            new XlsxVariableReplacer().Replace(doc, new Dictionary<string, string> { ["name"] = "World" });
        }

        // Assert: type preserved, text updated in <is><t>, no stale <v>, no t="str"
        using var reopen = SpreadsheetDocument.Open(path, false);
        var cell = GetFirstCell(reopen);
        Assert.Equal(CellValues.InlineString, cell.DataType!.Value);
        Assert.NotNull(cell.InlineString);
        Assert.Equal("Hello World", cell.InlineString!.InnerText);
        Assert.Null(cell.CellValue);
        OpenXmlAssert.NoValidationErrors(reopen);
    }

    [Fact]
    public void TemplateEngine_ProcessesProperInlineStringCell_InPlace()
    {
        // Arrange
        var path = CreateWorkbook(sd => AppendInlineStringCell(sd, "A1", "{{#if show}}Visible{{/if}}"));

        // Act
        using (var doc = SpreadsheetDocument.Open(path, true))
        {
            new XlsxTemplateEngine().Process(doc, new Dictionary<string, object> { ["show"] = true });
        }

        // Assert
        using var reopen = SpreadsheetDocument.Open(path, false);
        var cell = GetFirstCell(reopen);
        Assert.Equal(CellValues.InlineString, cell.DataType!.Value);
        Assert.Equal("Visible", cell.InlineString!.InnerText);
        Assert.Null(cell.CellValue);
        OpenXmlAssert.NoValidationErrors(reopen);
    }

    [Fact]
    public void Detector_FindsVariablesInProperInlineStringCell()
    {
        // Arrange
        var path = CreateWorkbook(sd => AppendInlineStringCell(sd, "A1", "Hello {{name}}"));

        // Act
        using var doc = SpreadsheetDocument.Open(path, false);
        var variables = new XlsxVariableDetector().Scan(doc);

        // Assert: the placeholder in <is><t> is discovered, not silently skipped
        var variable = Assert.Single(variables);
        Assert.Equal("name", variable.Name);
        Assert.Contains("A1", variable.Location);
    }

    [Fact]
    public void Replacer_PreservesLeadingAndTrailingWhitespace_InReplacedValue()
    {
        // Arrange: a shared-string cell whose replacement value is whitespace-padded.
        // The appended shared-string item must carry xml:space="preserve" (the contract
        // WorkbookBuilder applies) or Excel would trim the padding on display.
        var path = CreateWorkbook(
            sd => sd.Append(new Row(new Cell
            {
                CellReference = "A1",
                CellValue = new CellValue("0"),
                DataType = CellValues.SharedString
            })),
            sharedStrings: ["Hello {{name}}"]);

        // Act
        using (var doc = SpreadsheetDocument.Open(path, true))
        {
            new XlsxVariableReplacer().Replace(doc, new Dictionary<string, string> { ["name"] = "  padded  " });
        }

        // Assert: the appended <t> is space-preserved and the value round-trips verbatim
        using var reopen = SpreadsheetDocument.Open(path, false);
        var sharedStringTable = reopen.WorkbookPart!.SharedStringTablePart!.SharedStringTable!;
        var items = sharedStringTable.Elements<SharedStringItem>().ToList();
        Assert.Equal(2, items.Count);
        var appendedText = Assert.IsType<Text>(items[1].FirstChild);
        Assert.Equal(SpaceProcessingModeValues.Preserve, appendedText.Space!.Value);
        Assert.Equal("Hello   padded  ", items[1].InnerText);
    }

    [Fact]
    public void MergeVariables_ReplacesPlaceholderInInlineStringCell_ThroughPublicBuilderApi()
    {
        // Arrange: an inline-string template opened through the public builder API
        var path = CreateWorkbook(sd => AppendInlineStringCell(sd, "A1", "Hello {{name}}"));
        using (var document = SpreadsheetDocument.Open(path, true))
        {
            // The raw SDK save above flushed on dispose; re-open through the builder.
        }

        // Act
        using (var builder = WorkbookBuilder.Open(path))
        {
            builder.MergeVariables(new Dictionary<string, string> { ["name"] = "World" });
            builder.Save();
        }

        // Assert: merged through the builder, type preserved, no stale <v>
        using var reopen = SpreadsheetDocument.Open(path, false);
        var cell = GetFirstCell(reopen);
        Assert.Equal(CellValues.InlineString, cell.DataType!.Value);
        Assert.Equal("Hello World", cell.InlineString!.InnerText);
        Assert.Null(cell.CellValue);
        OpenXmlAssert.NoValidationErrors(reopen);
    }

    [Fact]
    public void DetectVariables_FindsInlineStringPlaceholders_ThroughPublicBuilderApi()
    {
        // Arrange
        var path = CreateWorkbook(sd => AppendInlineStringCell(sd, "A1", "Hello {{name}}"));

        // Act
        using var builder = WorkbookBuilder.Open(path);
        var variables = builder.DetectVariables();

        // Assert
        var variable = Assert.Single(variables);
        Assert.Equal("name", variable.Name);
    }

    public void Dispose()
    {
        foreach (var file in _tempFiles)
        {
            if (File.Exists(file))
            {
                File.Delete(file);
            }
        }
    }

    private string CreateWorkbook(Action<SheetData> configureSheetData, string[]? sharedStrings = null)
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx_variable_pipeline_{Guid.NewGuid():N}.xlsx");
        _tempFiles.Add(path);

        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook(new Sheets());

        if (sharedStrings is not null)
        {
            var sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
            sharedStringPart.SharedStringTable = new SharedStringTable(
                sharedStrings.Select(text => new SharedStringItem(new Text(text))));
        }

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        configureSheetData(sheetData);
        worksheetPart.Worksheet = new Worksheet(sheetData);

        var sheets = workbookPart.Workbook.GetFirstChild<Sheets>()!;
        sheets.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "Sheet1"
        });

        return path;
    }

    private static void AppendInlineStringCell(SheetData sheetData, string reference, string text)
    {
        var cell = new Cell
        {
            CellReference = reference,
            DataType = CellValues.InlineString,
            InlineString = new InlineString(new Text(text))
        };
        sheetData.Append(new Row(cell));
    }

    private static Cell GetFirstCell(SpreadsheetDocument document)
    {
        return document.WorkbookPart!.WorksheetParts.First()
            .Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>().First().Elements<Cell>().First();
    }
}
