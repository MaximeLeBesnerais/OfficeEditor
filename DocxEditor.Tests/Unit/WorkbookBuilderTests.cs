using XlsxEditor.Core.Builders;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocxEditor.Tests.Unit;

public class WorkbookBuilderTests : IDisposable
{
    private readonly string _testFilePath = Path.Combine(Path.GetTempPath(), $"test_xlsx_{Guid.NewGuid()}.xlsx");

    [Fact]
    public void Create_ShouldCreateNewWorkbook()
    {
        // Act
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.Save();
        }

        // Assert
        Assert.True(File.Exists(_testFilePath));
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        Assert.NotNull(doc.WorkbookPart);
    }

    [Fact]
    public void AddWorksheet_ShouldAddWorksheet()
    {
        // Act
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1");
            builder.Save();
        }

        // Assert
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var workbookPart = doc.WorkbookPart;
        Assert.NotNull(workbookPart);
        var workbook = workbookPart.Workbook;
        Assert.NotNull(workbook);
        var sheets = workbook.Sheets;
        Assert.NotNull(sheets);
        Assert.Single(sheets.Elements<Sheet>());
    }

    [Fact]
    public void AddCell_ShouldAddCellValue()
    {
        // Act
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var worksheet = builder.AddWorksheet("Sheet1");
            worksheet.AddCell("A1", "Hello World");
            builder.Save();
        }

        // Assert
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var worksheetPart = doc.WorkbookPart!.WorksheetParts.First();
        var sheetData = worksheetPart.Worksheet?.GetFirstChild<SheetData>();
        Assert.NotNull(sheetData);
        var cell = sheetData.Elements<Row>().First().Elements<Cell>().First();
        
        Assert.Equal("A1", cell.CellReference?.Value);
    }

    [Fact]
    public void AddHeaderRow_ShouldAddHeaders()
    {
        // Act
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var worksheet = builder.AddWorksheet("Sheet1");
            worksheet.AddHeaderRow(new List<string> { "Name", "Age", "City" });
            builder.Save();
        }

        // Assert
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var worksheetPart = doc.WorkbookPart!.WorksheetParts.First();
        var sheetData = worksheetPart.Worksheet?.GetFirstChild<SheetData>();
        Assert.NotNull(sheetData);
        var row = sheetData.Elements<Row>().First();
        var cells = row.Elements<Cell>().ToList();
        
        Assert.Equal(3, cells.Count);
    }

    [Fact]
    public void AddDataRow_ShouldAddData()
    {
        // Act
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var worksheet = builder.AddWorksheet("Sheet1");
            worksheet.AddHeaderRow(new List<string> { "Name", "Age" });
            worksheet.AddDataRow(new List<string> { "John", "30" }, 2);
            builder.Save();
        }

        // Assert
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var worksheetPart = doc.WorkbookPart!.WorksheetParts.First();
        var sheetData = worksheetPart.Worksheet?.GetFirstChild<SheetData>();
        Assert.NotNull(sheetData);
        var rows = sheetData.Elements<Row>().ToList();
        
        Assert.Equal(2, rows.Count);
    }

    [Fact]
    public void AddFormula_ShouldAddFormula()
    {
        // Act
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var worksheet = builder.AddWorksheet("Sheet1");
            worksheet.AddCell("A1", "10");
            worksheet.AddCell("A2", "20");
            worksheet.AddCell("A3", "=SUM(A1:A2)", true);
            builder.Save();
        }

        // Assert
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var worksheetPart = doc.WorkbookPart!.WorksheetParts.First();
        var sheetData = worksheetPart.Worksheet?.GetFirstChild<SheetData>();
        Assert.NotNull(sheetData);
        var row = sheetData.Elements<Row>().Last();
        var cell = row.Elements<Cell>().First();
        
        Assert.NotNull(cell.CellFormula);
        Assert.Equal("=SUM(A1:A2)", cell.CellFormula?.Text);
    }

    [Fact]
    public void DetectVariables_ShouldFindVariables()
    {
        // Arrange
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var worksheet = builder.AddWorksheet("Sheet1");
            worksheet.AddCell("A1", "Hello {{name}}");
            builder.Save();
        }

        // Act
        List<OfficeEditor.Core.Models.VariableInfo> variables;
        using (var builder = WorkbookBuilder.Open(_testFilePath))
        {
            variables = builder.DetectVariables();
        }

        // Assert
        Assert.Single(variables);
        Assert.Equal("name", variables[0].Name);
    }

    [Fact]
    public void MergeVariables_ShouldReplaceVariables()
    {
        // Arrange
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var worksheet = builder.AddWorksheet("Sheet1");
            worksheet.AddCell("A1", "Hello {{name}}");
            builder.Save();
        }

        // Act
        using (var builder = WorkbookBuilder.Open(_testFilePath))
        {
            builder.MergeVariables(new Dictionary<string, string>
            {
                ["name"] = "World"
            });
            builder.Save();
        }

        // Assert
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var worksheetPart = doc.WorkbookPart!.WorksheetParts.First();
        var sheetData = worksheetPart.Worksheet?.GetFirstChild<SheetData>();
        Assert.NotNull(sheetData);
        var cell = sheetData.Elements<Row>().First().Elements<Cell>().First();
        
        // The cell should have been updated
        Assert.NotNull(cell.CellValue);
    }

    public void Dispose()
    {
        if (File.Exists(_testFilePath))
        {
            File.Delete(_testFilePath);
        }
    }
}
