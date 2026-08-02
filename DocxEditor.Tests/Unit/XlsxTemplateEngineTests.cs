using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxEditor.Core.Variables;

namespace DocxEditor.Tests.Unit;

public sealed class XlsxTemplateEngineTests : IDisposable
{
    private readonly List<string> _tempFiles = [];

    [Fact]
    public void Process_LeavesCellUnchanged_WhenCellDoesNotContainTemplateMarker()
    {
        // Arrange
        var path = CreateWorkbook(sheetData => AddInlineCell(sheetData, "Plain text", CellValues.String));

        // Act
        Process(path, new Dictionary<string, object> { ["show"] = true });

        // Assert
        using var document = SpreadsheetDocument.Open(path, false);
        var cell = GetFirstCell(document);
        Assert.Equal("Plain text", cell.CellValue!.Text);
        Assert.Equal(CellValues.String, cell.DataType!.Value);
    }

    [Fact]
    public void Process_UpdatesNonSharedStringCell_AndSetsDataTypeToString()
    {
        // Arrange
        var path = CreateWorkbook(sheetData => AddInlineCell(sheetData, "{{#if show}}Visible{{/if}}", CellValues.InlineString));

        // Act
        Process(path, new Dictionary<string, object> { ["show"] = true });

        // Assert
        using var document = SpreadsheetDocument.Open(path, false);
        var cell = GetFirstCell(document);
        Assert.Equal("Visible", cell.CellValue!.Text);
        Assert.Equal(CellValues.String, cell.DataType!.Value);
    }

    [Fact]
    public void Process_ReadsValidSharedStringIndex_AndReusesExistingReplacementString()
    {
        // Arrange
        var path = CreateWorkbook(
            sheetData => AddSharedStringCell(sheetData, 0),
            sharedStrings: ["{{#if show}}Visible{{/if}}", "Visible"]);

        // Act
        Process(path, new Dictionary<string, object> { ["show"] = true });

        // Assert
        using var document = SpreadsheetDocument.Open(path, false);
        var cell = GetFirstCell(document);
        var sharedStrings = GetSharedStringItems(document);
        Assert.Equal("1", cell.CellValue!.Text);
        Assert.Equal(CellValues.SharedString, cell.DataType!.Value);
        Assert.Equal(2, sharedStrings.Count);
    }

    [Fact]
    public void Process_AppendsReplacementSharedString_WhenReplacementDoesNotExist()
    {
        // Arrange
        var path = CreateWorkbook(
            sheetData => AddSharedStringCell(sheetData, 0),
            sharedStrings: ["{{#if show}}New value{{/if}}"]);

        // Act
        Process(path, new Dictionary<string, object> { ["show"] = true });

        // Assert
        using var document = SpreadsheetDocument.Open(path, false);
        var cell = GetFirstCell(document);
        var sharedStrings = GetSharedStringItems(document);
        Assert.Equal("1", cell.CellValue!.Text);
        Assert.Equal(["{{#if show}}New value{{/if}}", "New value"], sharedStrings.Select(item => item.InnerText).ToArray());
    }

    [Theory]
    [InlineData("not-an-index")]
    [InlineData("99")]
    public void Process_LeavesSharedStringCellUnchanged_WhenIndexIsInvalidOrOutOfRange(string index)
    {
        // Arrange
        var path = CreateWorkbook(
            sheetData => AddSharedStringCell(sheetData, index),
            sharedStrings: ["{{#if show}}Visible{{/if}}"]);

        // Act
        Process(path, new Dictionary<string, object> { ["show"] = true });

        // Assert
        using var document = SpreadsheetDocument.Open(path, false);
        Assert.Equal(index, GetFirstCell(document).CellValue!.Text);
        Assert.Single(GetSharedStringItems(document));
    }

    [Fact]
    public void Process_LeavesSharedStringCellUnchanged_WhenSharedStringTableIsMissing()
    {
        // Arrange
        var path = CreateWorkbook(sheetData => AddSharedStringCell(sheetData, 0));

        // Act
        Process(path, new Dictionary<string, object> { ["show"] = true });

        // Assert
        using var document = SpreadsheetDocument.Open(path, false);
        var cell = GetFirstCell(document);
        Assert.Equal("0", cell.CellValue!.Text);
        Assert.Equal(CellValues.SharedString, cell.DataType!.Value);
    }

    [Theory]
    [InlineData("{{#if enabled}}Yes{{/if}}", true, "Yes")]
    [InlineData("{{#if enabled}}Yes{{/if}}", false, "")]
    [InlineData("{{#ifnot enabled}}No{{/ifnot}}", true, "")]
    [InlineData("{{#ifnot enabled}}No{{/ifnot}}", false, "No")]
    [InlineData("{{#if missing}}Yes{{/if}}", true, "")]
    [InlineData("{{#if label}}Yes{{/if}}", "text", "Yes")]
    [InlineData("{{#if label}}Yes{{/if}}", "", "")]
    [InlineData("{{#if label}}Yes{{/if}}", "false", "")]
    [InlineData("{{#if label}}Yes{{/if}}", "0", "")]
    public void Process_EvaluatesConditionals_WithBoolStringAndMissingValues(string template, object value, string expected)
    {
        // Arrange
        var path = CreateWorkbook(sheetData => AddInlineCell(sheetData, template));
        var data = template.Contains("missing", StringComparison.Ordinal)
            ? new Dictionary<string, object>()
            : new Dictionary<string, object> { ["enabled"] = value, ["label"] = value };

        // Act
        Process(path, data);

        // Assert
        using var document = SpreadsheetDocument.Open(path, false);
        Assert.Equal(expected, GetFirstCell(document).CellValue!.Text);
    }

    [Theory]
    [InlineData("{{#if count > 5}}High{{/if}}", 10, "High")]
    [InlineData("{{#if count < 5}}Low{{/if}}", 3, "Low")]
    [InlineData("{{#if count >= 5}}At least{{/if}}", 5, "At least")]
    [InlineData("{{#if count <= 5}}At most{{/if}}", 5, "At most")]
    [InlineData("{{#if status == 'open'}}Open{{/if}}", "open", "Open")]
    [InlineData("{{#if status != 'closed'}}Not closed{{/if}}", "open", "Not closed")]
    [InlineData("{{#if missing > 5}}Missing{{/if}}", 10, "")]
    public void Process_EvaluatesComparisons(string template, object value, string expected)
    {
        // Arrange
        var path = CreateWorkbook(sheetData => AddInlineCell(sheetData, template));
        var data = new Dictionary<string, object> { ["count"] = value, ["status"] = value };

        // Act
        Process(path, data);

        // Assert
        using var document = SpreadsheetDocument.Open(path, false);
        Assert.Equal(expected, GetFirstCell(document).CellValue!.Text);
    }

    [Theory]
    [InlineData("{{#if count > 5}}High{{/if}}", 3, "")]
    [InlineData("{{#if count < 5}}Low{{/if}}", 8, "")]
    [InlineData("{{#if count >= 5}}At least{{/if}}", 4, "")]
    [InlineData("{{#if count <= 5}}At most{{/if}}", 6, "")]
    [InlineData("{{#if status == 'open'}}Open{{/if}}", "closed", "")]
    [InlineData("{{#if status != 'closed'}}Not closed{{/if}}", "closed", "")]
    [InlineData("{{#if status <> 'closed'}}Invalid op{{/if}}", "open", "")]
    [InlineData("{{#if count > 5}}Lexical{{/if}}", "abc", "Lexical")]
    public void Process_EvaluatesFalseComparisonAndFallbackBranches(string template, object value, string expected)
    {
        // Arrange
        var path = CreateWorkbook(sheetData => AddInlineCell(sheetData, template));
        var data = new Dictionary<string, object> { ["count"] = value, ["status"] = value };

        // Act
        Process(path, data);

        // Assert
        using var document = SpreadsheetDocument.Open(path, false);
        Assert.Equal(expected, GetFirstCell(document).CellValue!.Text);
    }

    [Fact]
    public void Process_ExpandsLoop_ForValidListAndNullValues()
    {
        // Arrange
        var path = CreateWorkbook(sheetData => AddInlineCell(sheetData, "{{#each items}}{name}:{value}{{/each}}"));
        var data = new Dictionary<string, object>
        {
            ["items"] = new List<Dictionary<string, object>>
            {
                new() { ["name"] = "First", ["value"] = 1 },
                new() { ["name"] = "Second", ["value"] = null! }
            }
        };

        // Act
        Process(path, data);

        // Assert
        using var document = SpreadsheetDocument.Open(path, false);
        Assert.Equal("First:1\nSecond:", GetFirstCell(document).CellValue!.Text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Process_RemovesLoop_WhenSourceIsMissingOrNotAList(bool includeNonListSource)
    {
        // Arrange
        var path = CreateWorkbook(sheetData => AddInlineCell(sheetData, "Before {{#each items}}{name}{{/each}} After"));
        var data = includeNonListSource
            ? new Dictionary<string, object> { ["items"] = "not a list" }
            : [];

        // Act
        Process(path, data);

        // Assert
        using var document = SpreadsheetDocument.Open(path, false);
        Assert.Equal("Before  After", GetFirstCell(document).CellValue!.Text);
    }

    [Fact]
    public void Process_ReturnsSafely_WhenWorkbookPartIsMissing()
    {
        // Arrange
        var path = CreateEmptySpreadsheetDocument();

        // Act / Assert
        using var document = SpreadsheetDocument.Open(path, true);
        var exception = Record.Exception(() => new XlsxTemplateEngine().Process(document, []));
        Assert.Null(exception);
    }

    [Fact]
    public void Process_WithNullDocument_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>(() => new XlsxTemplateEngine().Process(null!, []));
    }

    [Fact]
    public void Process_WithNullData_ThrowsArgumentNullException_EvenWithoutAnyMarkers()
    {
        // A null data dictionary must fail up front regardless of whether any cell contains a
        // '{{…}}' marker; previously it only failed (with an ANE/NRE deep in evaluation) when
        // a marker was actually present, making behavior marker-dependent.
        var path = CreateWorkbook(sheetData => AddInlineCell(sheetData, "Plain text", CellValues.String));

        using var document = SpreadsheetDocument.Open(path, true);
        Assert.Throws<ArgumentNullException>(() => new XlsxTemplateEngine().Process(document, null!));
    }

    [Fact]
    public void Process_ReturnsSafely_WhenWorkbookIsMissing()
    {
        // Arrange
        var path = CreateSpreadsheetDocument(document => document.AddWorkbookPart());

        // Act / Assert
        using var spreadsheet = SpreadsheetDocument.Open(path, true);
        var exception = Record.Exception(() => new XlsxTemplateEngine().Process(spreadsheet, []));
        Assert.Null(exception);
    }

    [Fact]
    public void Process_ReturnsSafely_WhenSheetsAreMissing()
    {
        // Arrange
        var path = CreateSpreadsheetDocument(document => document.AddWorkbookPart().Workbook = new Workbook());

        // Act / Assert
        using var spreadsheet = SpreadsheetDocument.Open(path, true);
        var exception = Record.Exception(() => new XlsxTemplateEngine().Process(spreadsheet, []));
        Assert.Null(exception);
    }

    [Fact]
    public void Process_SkipsWorksheet_WhenSheetDataIsMissing()
    {
        // Arrange
        var path = CreateWorkbook(_ => { }, includeSheetData: false);

        // Act / Assert
        using var spreadsheet = SpreadsheetDocument.Open(path, true);
        var exception = Record.Exception(() => new XlsxTemplateEngine().Process(spreadsheet, []));
        Assert.Null(exception);
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

    private string CreateWorkbook(
        Action<SheetData> configureSheetData,
        string[]? sharedStrings = null,
        bool includeSheetData = true)
    {
        return CreateSpreadsheetDocument(document =>
        {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook(new Sheets());

            if (sharedStrings is not null)
            {
                var sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
                sharedStringPart.SharedStringTable = new SharedStringTable(sharedStrings.Select(text => new SharedStringItem(new Text(text))));
            }

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            if (includeSheetData)
            {
                var sheetData = new SheetData();
                configureSheetData(sheetData);
                worksheetPart.Worksheet = new Worksheet(sheetData);
            }
            else
            {
                worksheetPart.Worksheet = new Worksheet();
            }

            var sheets = workbookPart.Workbook.GetFirstChild<Sheets>()!;
            sheets.Append(new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Sheet1"
            });
        });
    }

    private string CreateEmptySpreadsheetDocument()
    {
        return CreateSpreadsheetDocument(_ => { });
    }

    private string CreateSpreadsheetDocument(Action<SpreadsheetDocument> configure)
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx_template_engine_{Guid.NewGuid():N}.xlsx");
        _tempFiles.Add(path);

        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        configure(document);
        return path;
    }

    private static void Process(string path, Dictionary<string, object> data)
    {
        using var document = SpreadsheetDocument.Open(path, true);
        new XlsxTemplateEngine().Process(document, data);
    }

    private static void AddInlineCell(SheetData sheetData, string value, CellValues? dataType = null)
    {
        sheetData.Append(new Row(new Cell
        {
            CellValue = new CellValue(value),
            DataType = dataType ?? CellValues.String
        }));
    }

    private static void AddSharedStringCell(SheetData sheetData, int index)
    {
        AddSharedStringCell(sheetData, index.ToString());
    }

    private static void AddSharedStringCell(SheetData sheetData, string index)
    {
        sheetData.Append(new Row(new Cell
        {
            CellValue = new CellValue(index),
            DataType = CellValues.SharedString
        }));
    }

    private static Cell GetFirstCell(SpreadsheetDocument document)
    {
        return document.WorkbookPart!.WorksheetParts.First().Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>().First().Elements<Cell>().First();
    }

    private static List<SharedStringItem> GetSharedStringItems(SpreadsheetDocument document)
    {
        return document.WorkbookPart!.GetPartsOfType<SharedStringTablePart>().First().SharedStringTable!.Elements<SharedStringItem>().ToList();
    }
}
