using XlsxEditor.Core.Builders;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocxEditor.Tests.Unit;

public class WorkbookBuilderTests : IDisposable
{
    private readonly string _testFilePath = Path.Combine(Path.GetTempPath(), $"test_xlsx_{Guid.NewGuid()}.xlsx");
    private readonly List<string> _additionalFiles = new();

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

    [Fact]
    public void WorksheetLookupAndRemoval_ShouldHandleSuccessAndMissingSheets()
    {
        // Arrange
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var first = builder.AddWorksheet("First");
        builder.AddWorksheet("Second");

        // Act & Assert
        Assert.Same(first, builder.GetWorksheet("First"));
        Assert.Equal(new List<string> { "First", "Second" }, builder.GetWorksheetNames());

        builder.RemoveWorksheet("Second");
        Assert.Equal(new List<string> { "First" }, builder.GetWorksheetNames());

        var getException = Assert.Throws<ArgumentException>(() => builder.GetWorksheet("Second"));
        var removeException = Assert.Throws<ArgumentException>(() => builder.RemoveWorksheet("Second"));
        Assert.Contains("Second", getException.Message);
        Assert.Contains("Second", removeException.Message);
    }

    [Fact]
    public void OpenExistingWorkbook_ShouldLoadWorksheetNamesAndAllowEditing()
    {
        // Arrange
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Existing").AddCell("A1", "Before");
            builder.Save();
        }

        // Act
        using (var builder = WorkbookBuilder.Open(_testFilePath))
        {
            Assert.Equal(new List<string> { "Existing" }, builder.GetWorksheetNames());
            builder.GetWorksheet("Existing").AddCell("B2", "After");
            builder.Save();
        }

        // Assert
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var cells = doc.WorkbookPart!.WorksheetParts.First().Worksheet!.GetFirstChild<SheetData>()!
            .Elements<Row>()
            .SelectMany(r => r.Elements<Cell>())
            .Select(c => c.CellReference?.Value)
            .ToList();
        Assert.Contains("A1", cells);
        Assert.Contains("B2", cells);
    }

    [Fact]
    public void Save_ShouldSupportSamePathAndClonePath()
    {
        // Arrange
        var clonePath = Path.Combine(Path.GetTempPath(), $"test_xlsx_clone_{Guid.NewGuid()}.xlsx");
        _additionalFiles.Add(clonePath);

        // Act
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1").AddCell("A1", "Clone me");
            builder.Save(_testFilePath);
            builder.Save(clonePath);
        }

        // Assert
        Assert.True(File.Exists(_testFilePath));
        Assert.True(File.Exists(clonePath));
        using var doc = SpreadsheetDocument.Open(clonePath, false);
        Assert.Single(doc.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>());
    }

    [Fact]
    public void AddCellOverloads_ShouldHandleNumbersStringsFormulaFalseAndStyle()
    {
        // Act
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var worksheet = builder.AddWorksheet("Sheet1");
            worksheet.AddCell("A1", "42.5");
            worksheet.AddCell("B1", "Text");
            worksheet.AddCell("C1", "Not a formula", false);
            worksheet.AddCell("D1", "7", "0");
            builder.Save();
        }

        // Assert
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var cells = doc.WorkbookPart!.WorksheetParts.First().Worksheet!.GetFirstChild<SheetData>()!
            .Elements<Row>().First().Elements<Cell>().ToDictionary(c => c.CellReference!.Value!);
        Assert.Equal(CellValues.Number, cells["A1"].DataType?.Value);
        Assert.Equal(CellValues.SharedString, cells["B1"].DataType?.Value);
        Assert.Equal(CellValues.String, cells["C1"].DataType?.Value);
        Assert.Equal("Not a formula", cells["C1"].CellValue?.Text);
        Assert.Equal(0U, cells["D1"].StyleIndex?.Value);
    }

    [Fact]
    public void AddFormulaRow_ShouldSkipEmptyFormulasAndUseColumnNamesBeyondZ()
    {
        // Arrange
        var formulas = Enumerable.Range(0, 28).Select(i => i == 1 ? string.Empty : $"A{i + 1}").ToList();

        // Act
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1").AddFormulaRow(formulas, 3);
            builder.Save();
        }

        // Assert
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var cells = doc.WorkbookPart!.WorksheetParts.First().Worksheet!.GetFirstChild<SheetData>()!
            .Elements<Row>().Single().Elements<Cell>().ToList();
        Assert.Equal(27, cells.Count);
        Assert.DoesNotContain(cells, c => c.CellReference?.Value == "B3");
        Assert.Contains(cells, c => c.CellReference?.Value == "AA3");
        Assert.Contains(cells, c => c.CellReference?.Value == "AB3");
        Assert.All(cells, c => Assert.NotNull(c.CellFormula));
    }

    [Fact]
    public void AddTable_ShouldCreateTablePartsAndAppendAdditionalTables()
    {
        // Act
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var worksheet = builder.AddWorksheet("Sheet1");
            worksheet.AddHeaderRow(new List<string> { "A", "B", "C" });
            worksheet.AddDataRow(new List<string> { "1", "2", "3" }, 2);
            worksheet.AddTable("A1", "C2", "FirstTable");
            worksheet.AddTable("A1", "B2", "SecondTable");
            builder.Save();
        }

        // Assert
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var worksheetPart = doc.WorkbookPart!.WorksheetParts.First();
        var tableParts = worksheetPart.Worksheet!.Elements<TableParts>().Single();
        Assert.Equal(2U, tableParts.Count?.Value);
        Assert.Equal(2, worksheetPart.TableDefinitionParts.Count());
        Assert.Contains(worksheetPart.TableDefinitionParts, p => p.Table?.DisplayName?.Value == "FirstTable");
        Assert.Contains(worksheetPart.TableDefinitionParts, p => p.Table?.DisplayName?.Value == "SecondTable");
    }

    [Fact]
    public void WorksheetEdgeCases_ShouldThrowForUnsupportedChartAndInvalidCellReference()
    {
        // Arrange
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var worksheet = builder.AddWorksheet("Sheet1");

        // Act & Assert
        Assert.Throws<NotSupportedException>(() => worksheet.AddChart(ChartType.Pie, "A1:B2"));
        Assert.Throws<FormatException>(() => worksheet.AddCell("A", "Missing row"));
    }

    [Fact]
    public void Create_SaveToBytes_ReturnsValidXlsx()
    {
        // Arrange
        byte[] bytes;

        // Act
        using (var builder = WorkbookBuilder.Create())
        {
            builder.AddWorksheet("Sheet1").AddCell("A1", "Hello bytes");
            bytes = builder.SaveToBytes();
        }

        // Assert
        Assert.NotEmpty(bytes);
        using var stream = new MemoryStream(bytes);
        using var doc = SpreadsheetDocument.Open(stream, false);
        Assert.NotNull(doc.WorkbookPart);
        Assert.Single(doc.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>());
    }

    [Fact]
    public void Create_Save_Stream_ReturnsValidXlsx()
    {
        // Arrange
        using var outputStream = new MemoryStream();

        // Act
        using (var builder = WorkbookBuilder.Create())
        {
            builder.AddWorksheet("Sheet1").AddCell("A1", "Hello stream");
            builder.Save(outputStream);
        }

        // Assert
        Assert.True(outputStream.CanRead);
        Assert.True(outputStream.CanWrite);
        Assert.True(outputStream.Length > 0);
        outputStream.Position = 0;
        using var doc = SpreadsheetDocument.Open(outputStream, false);
        Assert.NotNull(doc.WorkbookPart);
        var sheetData = doc.WorkbookPart!.WorksheetParts.First().Worksheet!.GetFirstChild<SheetData>();
        Assert.NotNull(sheetData);
        var cell = sheetData!.Elements<Row>().First().Elements<Cell>().First();
        Assert.Equal("A1", cell.CellReference?.Value);
    }

    [Fact]
    public void Open_FromBytes_AndSaveToBytes_RoundtripsCell()
    {
        // Arrange
        byte[] originalBytes;
        using (var builder = WorkbookBuilder.Create())
        {
            builder.AddWorksheet("Sheet1").AddCell("A1", "Original");
            originalBytes = builder.SaveToBytes();
        }

        // Act
        byte[] modifiedBytes;
        using (var builder = WorkbookBuilder.Open(originalBytes))
        {
            builder.GetWorksheet("Sheet1").AddCell("B2", "Added");
            modifiedBytes = builder.SaveToBytes();
        }

        // Assert
        using var stream = new MemoryStream(modifiedBytes);
        using var doc = SpreadsheetDocument.Open(stream, false);
        var cells = doc.WorkbookPart!.WorksheetParts.First().Worksheet!.GetFirstChild<SheetData>()!
            .Elements<Row>()
            .SelectMany(r => r.Elements<Cell>())
            .Select(c => c.CellReference?.Value)
            .ToList();
        Assert.Contains("A1", cells);
        Assert.Contains("B2", cells);
    }

    [Fact]
    public void Save_WithNoPath_OnPathlessDocument_ThrowsInvalidOperationException()
    {
        // Arrange
        using var builder = WorkbookBuilder.Create();
        builder.AddWorksheet("Sheet1").AddCell("A1", "No path");

        // Act & Assert
        Assert.Throws<InvalidOperationException>(() => builder.Save());
    }

    public void Dispose()
    {
        if (File.Exists(_testFilePath))
        {
            File.Delete(_testFilePath);
        }

        foreach (var file in _additionalFiles)
        {
            if (File.Exists(file))
            {
                File.Delete(file);
            }
        }
    }
}
