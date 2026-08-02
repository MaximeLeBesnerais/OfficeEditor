using XlsxEditor.Core.Builders;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxEditor.Core.Exceptions;

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
        OpenXmlAssert.NoValidationErrors(_testFilePath);
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
        // SpreadsheetML stores formula text WITHOUT the leading '='; storing it
        // triggers Excel's repair prompt. The builder must strip it.
        Assert.Equal("SUM(A1:A2)", cell.CellFormula?.Text);
        OpenXmlAssert.NoValidationErrors(_testFilePath);
    }

    [Fact]
    public void GetCellFormula_ShouldReturnDisplaySyntax_WithLeadingEquals()
    {
        // Act
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1").AddCell("A3", "=SUM(A1:A2)", true);
            builder.Save();
        }

        // Assert: stored text has no '=', but the read API re-prepends it so
        // callers always see Excel display syntax.
        using var reader = WorkbookBuilder.Open(_testFilePath);
        var ws = reader.GetWorksheet("Sheet1");
        Assert.Equal("=SUM(A1:A2)", ws.GetCellFormula("A3"));
        Assert.Equal("=SUM(A1:A2)", ws.GetCellInfo("A3")?.Formula);
    }

    [Fact]
    public void AddCell_FormulaWithoutLeadingEquals_ShouldBeAccepted()
    {
        // Act: callers may pass the bare stored form too
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1").AddCell("A1", "SUM(B1:B2)", true);
            builder.Save();
        }

        // Assert
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var cell = doc.WorkbookPart!.WorksheetParts.First().Worksheet!
            .GetFirstChild<SheetData>()!.Elements<Row>().First().Elements<Cell>().First();
        Assert.Equal("SUM(B1:B2)", cell.CellFormula?.Text);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Theory]
    [InlineData("42.5", true)]
    [InlineData("-7", true)]
    [InlineData("1e3", true)]
    [InlineData("NaN", false)]
    [InlineData("Infinity", false)]
    [InlineData("-Infinity", false)]
    [InlineData("1,000", false)]
    public void AddCell_NumericDetection_ShouldBeInvariantAndRejectNonFinite(string value, bool isNumber)
    {
        // Act
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1").AddCell("A1", value);
            builder.Save();
        }

        // Assert
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var cell = doc.WorkbookPart!.WorksheetParts.First().Worksheet!
            .GetFirstChild<SheetData>()!.Elements<Row>().Single().Elements<Cell>().Single();
        Assert.Equal(isNumber ? CellValues.Number : CellValues.SharedString, cell.DataType?.Value);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void AddHeaderRow_CalledRepeatedly_ShouldReuseOneHeaderStyle()
    {
        // Act: three header rows must not append three fonts + three cell formats
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var sheet = builder.AddWorksheet("Sheet1");
            sheet.AddHeaderRow(new List<string> { "A" }, 1);
            sheet.AddHeaderRow(new List<string> { "B" }, 5);
            sheet.AddHeaderRow(new List<string> { "C" }, 9);
            builder.Save();
        }

        // Assert: exactly one default + one bold header format
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var stylesheet = doc.WorkbookPart!.WorkbookStylesPart!.Stylesheet;
        Assert.NotNull(stylesheet);
        Assert.Equal(2U, stylesheet.CellFormats!.Count?.Value);
        Assert.Equal(2U, stylesheet.Fonts!.Count?.Value);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void SharedStrings_ShouldDeduplicateAndPreserveWhitespace()
    {
        // Act: same string twice → one entry; padded string keeps its spaces
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var sheet = builder.AddWorksheet("Sheet1");
            sheet.AddCell("A1", "repeat");
            sheet.AddCell("A2", "repeat");
            sheet.AddCell("A3", "  padded  ");
            builder.Save();
        }

        // Assert
        using (var doc = SpreadsheetDocument.Open(_testFilePath, false))
        {
            var table = doc.WorkbookPart!.SharedStringTablePart!.SharedStringTable;
            Assert.NotNull(table);
            var items = table.Elements<SharedStringItem>().ToList();
            Assert.Equal(2, items.Count);
            Assert.Equal("repeat", items[0].InnerText);
            var paddedText = Assert.IsType<Text>(items[1].FirstChild);
            Assert.Equal(SpaceProcessingModeValues.Preserve, paddedText.Space?.Value);
            OpenXmlAssert.NoValidationErrors(doc);
        }

        // Round-trip: whitespace survives reload
        using var reader = WorkbookBuilder.Open(_testFilePath);
        Assert.Equal("  padded  ", reader.GetWorksheet("Sheet1").GetCellValue("A3"));
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
    public void MergeVariables_NullDictionary_ThrowsArgumentNullException()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        builder.AddWorksheet("Sheet1").AddCell("A1", "Hello {{name}}");

        var ex = Assert.Throws<ArgumentNullException>(() => builder.MergeVariables(null!));
        Assert.Equal("data", ex.ParamName);
    }

    [Fact]
    public void MergeVariables_NullValue_ThrowsArgumentNullException_AndLeavesWorkbookUnmodified()
    {
        // Arrange
        using (var creator = WorkbookBuilder.Create(_testFilePath))
        {
            creator.AddWorksheet("Sheet1").AddCell("A1", "Hello {{name}}");
            creator.Save();
        }

        // Act & Assert: a null replacement value must fail loudly, never be written as
        // an empty string.
        using (var editor = WorkbookBuilder.Open(_testFilePath))
        {
            var ex = Assert.Throws<ArgumentNullException>(() =>
                editor.MergeVariables(new Dictionary<string, string> { ["name"] = null! }));
            Assert.Contains("name", ex.Message);
        }

        // Atomic: the placeholder survives because the merge was rejected up front.
        using var reader = WorkbookBuilder.Open(_testFilePath);
        Assert.Equal("Hello {{name}}", reader.GetWorksheet("Sheet1").GetCellValue("A1"));
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
            // Header row on row 2 creates the stylesheet so styleId "0" is valid.
            worksheet.AddHeaderRow(new List<string> { "H" }, 2);
            worksheet.AddCell("D1", "7", "0");
            builder.Save();
        }

        // Assert
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var cells = doc.WorkbookPart!.WorksheetParts.First().Worksheet!.GetFirstChild<SheetData>()!
            .Elements<Row>().First().Elements<Cell>().ToDictionary(c => c.CellReference!.Value!);
        Assert.Equal(CellValues.Number, cells["A1"].DataType?.Value);
        Assert.Equal(CellValues.SharedString, cells["B1"].DataType?.Value);
        // Literal strings go through the shared string table — t="str" is
        // reserved for formula string results, not literal values.
        Assert.Equal(CellValues.SharedString, cells["C1"].DataType?.Value);
        var sharedStringTable = doc.WorkbookPart!.SharedStringTablePart?.SharedStringTable;
        Assert.NotNull(sharedStringTable);
        var sharedStrings = sharedStringTable.Elements<SharedStringItem>().ToList();
        var c1Text = sharedStrings[int.Parse(cells["C1"].CellValue!.Text)].InnerText;
        Assert.Equal("Not a formula", c1Text);
        Assert.Equal(0U, cells["D1"].StyleIndex?.Value);
        OpenXmlAssert.NoValidationErrors(doc);
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

    // ─── Null validation: builder row/list/merge APIs ─────────────

    [Fact]
    public void AddHeaderRow_NullList_ThrowsArgumentNullException()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var worksheet = builder.AddWorksheet("Sheet1");

        var ex = Assert.Throws<ArgumentNullException>(() => worksheet.AddHeaderRow(null!));
        Assert.Equal("values", ex.ParamName);
    }

    [Fact]
    public void AddDataRow_NullList_ThrowsArgumentNullException()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var worksheet = builder.AddWorksheet("Sheet1");

        var ex = Assert.Throws<ArgumentNullException>(() => worksheet.AddDataRow(null!, 2));
        Assert.Equal("values", ex.ParamName);
    }

    [Fact]
    public void AddFormulaRow_NullList_ThrowsArgumentNullException()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var worksheet = builder.AddWorksheet("Sheet1");

        var ex = Assert.Throws<ArgumentNullException>(() => worksheet.AddFormulaRow(null!, 2));
        Assert.Equal("formulas", ex.ParamName);
    }

    [Fact]
    public void AddHeaderRow_NullElement_ThrowsBeforeAnyWorksheetMutation()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var worksheet = builder.AddWorksheet("Sheet1");

        var ex = Assert.Throws<ArgumentNullException>(
            () => worksheet.AddHeaderRow(new List<string> { "ok", null!, "later" }));
        Assert.Equal("values[1]", ex.ParamName);

        // Atomic: no row, header style or cell was written before the failure.
        Assert.Equal((0, 0, 0, 0), worksheet.GetDimensions());
        Assert.Empty(worksheet.GetRows());
    }

    [Fact]
    public void AddDataRow_NullElement_ThrowsBeforeAnyWorksheetMutation()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var worksheet = builder.AddWorksheet("Sheet1");

        var ex = Assert.Throws<ArgumentNullException>(
            () => worksheet.AddDataRow(new List<string> { "ok", null! }, 2));
        Assert.Equal("values[1]", ex.ParamName);
        Assert.Equal((0, 0, 0, 0), worksheet.GetDimensions());
        Assert.Empty(worksheet.GetRows());
    }

    [Fact]
    public void AddFormulaRow_NullElement_ThrowsBeforeAnyWorksheetMutation()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var worksheet = builder.AddWorksheet("Sheet1");

        var ex = Assert.Throws<ArgumentNullException>(
            () => worksheet.AddFormulaRow(new List<string> { "=SUM(A1:A2)", null! }, 2));
        Assert.Equal("formulas[1]", ex.ParamName);
        Assert.Equal((0, 0, 0, 0), worksheet.GetDimensions());
        Assert.Empty(worksheet.GetRows());
    }

    [Fact]
    public void RowMethods_EmptyLists_RemainSupported()
    {
        // Empty collections are legitimate (e.g. an empty data source): they must not
        // throw and they write no cells.
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var sheet = builder.AddWorksheet("Sheet1");
            sheet.AddHeaderRow(new List<string>(), 1);
            sheet.AddDataRow(new List<string>(), 2);
            sheet.AddFormulaRow(new List<string>(), 3);
            builder.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        var ws = reader.GetWorksheet("Sheet1");
        Assert.Empty(ws.GetRows().SelectMany(r => r.Cells));
    }

    [Fact]
    public void AddTable_ShouldCreateTablePartsAndAppendAdditionalTables()
    {
        // Act: two NON-overlapping tables on one sheet (overlaps are rejected — see below)
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var worksheet = builder.AddWorksheet("Sheet1");
            worksheet.AddHeaderRow(new List<string> { "A", "B", "C" });
            worksheet.AddDataRow(new List<string> { "1", "2", "3" }, 2);
            worksheet.AddTable("A1", "C2", "FirstTable");
            worksheet.AddTable("E1", "F2", "SecondTable");
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
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void AddTable_ShouldDeriveColumnNamesFromHeaderRow()
    {
        // Act
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var worksheet = builder.AddWorksheet("Sheet1");
            worksheet.AddHeaderRow(new List<string> { "Product", "Qty", "Price" });
            worksheet.AddDataRow(new List<string> { "Widget", "3", "9.99" }, 2);
            worksheet.AddTable("A1", "C2", "Sales");
            builder.Save();
        }

        // Assert: Excel requires TableColumn names to match the header cell text
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var table = doc.WorkbookPart!.WorksheetParts.First()
            .TableDefinitionParts.Single().Table!;
        var names = table.GetFirstChild<TableColumns>()!.Elements<TableColumn>()
            .Select(c => c.Name?.Value).ToList();
        Assert.Equal(new[] { "Product", "Qty", "Price" }, names);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void AddTable_ShouldFallbackToGeneratedNames_ForEmptyHeaderCells()
    {
        // Act: table over cells with no header values
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var worksheet = builder.AddWorksheet("Sheet1");
            worksheet.AddTable("A1", "B2", "EmptyHeaders");
            builder.Save();
        }

        // Assert: column names must be non-empty — generated fallback
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var table = doc.WorkbookPart!.WorksheetParts.First()
            .TableDefinitionParts.Single().Table!;
        var names = table.GetFirstChild<TableColumns>()!.Elements<TableColumn>()
            .Select(c => c.Name?.Value).ToList();
        Assert.Equal(new[] { "Column1", "Column2" }, names);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void AddTable_ShouldMakeDuplicateHeaderNamesUnique()
    {
        // Act
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var worksheet = builder.AddWorksheet("Sheet1");
            worksheet.AddHeaderRow(new List<string> { "Value", "Value" });
            worksheet.AddDataRow(new List<string> { "1", "2" }, 2);
            worksheet.AddTable("A1", "B2", "Duplicates");
            builder.Save();
        }

        // Assert: Excel requires unique column names within a table
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var table = doc.WorkbookPart!.WorksheetParts.First()
            .TableDefinitionParts.Single().Table!;
        var names = table.GetFirstChild<TableColumns>()!.Elements<TableColumn>()
            .Select(c => c.Name?.Value).ToList();
        Assert.Equal(new[] { "Value", "Value2" }, names);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void AddTable_ReversedRange_ShouldThrow()
    {
        // Arrange: a reversed range used to cast a negative column count to uint
        // (≈4 billion columns) — a corrupt table part.
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var worksheet = builder.AddWorksheet("Sheet1");

        // Act & Assert
        var ex = Assert.Throws<XlsxException>(() => worksheet.AddTable("C2", "A1", "Backwards"));
        Assert.Contains("reversed", ex.Message);
    }

    [Fact]
    public void AddTable_OverlappingRange_ShouldThrow()
    {
        // Arrange
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var worksheet = builder.AddWorksheet("Sheet1");
        worksheet.AddHeaderRow(new List<string> { "A", "B", "C" });
        worksheet.AddDataRow(new List<string> { "1", "2", "3" }, 2);
        worksheet.AddTable("A1", "C2", "FirstTable");

        // Act & Assert: Excel rejects worksheets with overlapping tables
        var ex = Assert.Throws<XlsxException>(() => worksheet.AddTable("B1", "D2", "SecondTable"));
        Assert.Contains("overlap", ex.Message.ToLowerInvariant());
        Assert.Contains("FirstTable", ex.Message);
    }

    [Theory]
    [InlineData("My Table")]
    [InlineData("Table!")]
    [InlineData("1Table")]
    [InlineData("")]
    public void AddTable_InvalidDisplayName_ShouldThrow(string tableName)
    {
        // Arrange
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var worksheet = builder.AddWorksheet("Sheet1");

        // Act & Assert
        var ex = Assert.Throws<XlsxException>(() => worksheet.AddTable("A1", "B2", tableName));
        Assert.Contains("table name", ex.Message.ToLowerInvariant());
    }

    [Theory]
    // A1-style references (within Excel's real grid)
    [InlineData("A1")]
    [InlineData("AA1")]
    [InlineData("AB12")]
    [InlineData("XFD1048576")]
    [InlineData("A1048576")]
    [InlineData("R1")]  // A1 column R row 1 (also a row-only R1C1 reference)
    [InlineData("C1")]
    // R1C1-style references
    [InlineData("R1C1")]
    [InlineData("R2C3")]
    [InlineData("r1c1")]
    [InlineData("R1048576C16384")]
    public void AddTable_CellReferenceName_ShouldThrow(string tableName)
    {
        // Excel interprets names that look like A1/R1C1 references as references, not
        // names, so they must be rejected before any table part is created.
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var worksheet = builder.AddWorksheet("Sheet1");

        var ex = Assert.Throws<XlsxException>(() => worksheet.AddTable("A1", "B2", tableName));
        Assert.Contains("reference", ex.Message.ToLowerInvariant());
    }

    [Theory]
    // Ordinary names and names containing digits stay accepted.
    [InlineData("Sales2024")]
    [InlineData("Table1")]
    [InlineData("Q1_2023")]
    [InlineData("DataTable")]
    [InlineData("data.table")]
    [InlineData("_private")]
    [InlineData("R")]
    [InlineData("C")]
    [InlineData("R2D2")]
    [InlineData("XFE1")]     // column beyond XFD — not a real cell reference
    [InlineData("A1048577")] // row beyond 1,048,576 — not a real cell reference
    [InlineData("R1C16385")] // R1C1 syntax but column beyond XFD
    public void AddTable_OrdinaryNamesAndNamesWithDigits_ShouldBeAccepted(string tableName)
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var worksheet = builder.AddWorksheet("Sheet1");
            worksheet.AddCell("A1", "Product");
            worksheet.AddCell("A2", "Widget");
            worksheet.AddTable("A1", "A2", tableName);
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        Assert.Contains(doc.WorkbookPart!.WorksheetParts.First().TableDefinitionParts,
            p => p.Table?.DisplayName?.Value == tableName);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void AddTable_DuplicateDisplayName_ShouldThrow_AcrossWorksheets()
    {
        // Arrange: table names must be unique workbook-wide (case-insensitive)
        using var builder = WorkbookBuilder.Create(_testFilePath);
        builder.AddWorksheet("Sheet1").AddTable("A1", "B2", "SalesTable");

        // Act & Assert
        var other = builder.AddWorksheet("Sheet2");
        var ex = Assert.Throws<XlsxException>(() => other.AddTable("A1", "B2", "SALESTABLE"));
        Assert.Contains("SALESTABLE", ex.Message);
        Assert.Contains("unique", ex.Message);
    }

    [Fact]
    public void WorksheetEdgeCases_ShouldThrowForUnsupportedChartAndInvalidCellReference()
    {
        // Arrange
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var worksheet = builder.AddWorksheet("Sheet1");

        // Act & Assert
        Assert.Throws<XlsxException>(() => worksheet.AddChart(ChartType.Pie, "A1:B2"));
        // Malformed references must surface as XlsxException naming the ref,
        // not a raw FormatException from int.Parse deep in the builder.
        var ex = Assert.Throws<XlsxException>(() => worksheet.AddCell("A", "Missing row"));
        Assert.Contains("'A'", ex.Message);
    }

    [Theory]
    [InlineData("A")]
    [InlineData("1A")]
    [InlineData("A0")]
    [InlineData("")]
    [InlineData("A 1")]
    [InlineData("A1B")]
    [InlineData("AAAA1")]
    public void AddCell_MalformedReference_ShouldThrowXlsxException(string badReference)
    {
        // Arrange
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var worksheet = builder.AddWorksheet("Sheet1");

        // Act & Assert
        var ex = Assert.Throws<XlsxException>(() => worksheet.AddCell(badReference, "value"));
        Assert.Contains("Invalid cell reference", ex.Message);
    }

    [Theory]
    [InlineData("XFD1")]
    [InlineData("XFD1048576")]
    [InlineData("A1048576")]
    public void AddCell_MaxExcelBounds_ShouldSucceed(string reference)
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var worksheet = builder.AddWorksheet("Sheet1");

        worksheet.AddCell(reference, "edge");

        Assert.True(worksheet.CellExists(reference));
        Assert.Equal("edge", worksheet.GetCellValue(reference));
    }

    [Theory]
    [InlineData("XFE1")]
    [InlineData("XFD1048577")]
    [InlineData("A1048577")]
    [InlineData("XFE1048577")]
    public void AddCell_BeyondExcelBounds_ShouldThrowXlsxException(string reference)
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var worksheet = builder.AddWorksheet("Sheet1");

        var ex = Assert.Throws<XlsxException>(() => worksheet.AddCell(reference, "value"));
        Assert.Contains(reference, ex.Message);
    }

    [Theory]
    [InlineData("A99999999999999999999")]
    [InlineData("XFD99999999999999999999999999")]
    public void AddCell_HugeRowString_ShouldThrowXlsxException_NotOverflow(string reference)
    {
        // A row far beyond int range must fail predictably as XlsxException,
        // never as an OverflowException from int.Parse deep in the parser.
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var worksheet = builder.AddWorksheet("Sheet1");

        var ex = Assert.Throws<XlsxException>(() => worksheet.AddCell(reference, "value"));
        Assert.Contains("row", ex.Message.ToLowerInvariant());
    }

    [Fact]
    public void AddCell_XFD1048576_LastValidCell_RoundTripsThroughSaveAndOpen()
    {
        // XFD1048576 is the last real cell Excel supports. It must be writable AND
        // readable back through the public builder API after a save/reload round-trip.
        byte[] bytes;
        using (var builder = WorkbookBuilder.Create())
        {
            builder.AddWorksheet("Sheet1").AddCell("XFD1048576", "edge");
            bytes = builder.SaveToBytes();
        }

        using var reader = WorkbookBuilder.Open(bytes);
        var ws = reader.GetWorksheet("Sheet1");
        Assert.Equal("edge", ws.GetCellValue("XFD1048576"));
        Assert.True(ws.CellExists("XFD1048576"));
        Assert.Equal("XFD1048576", ws.GetCellInfo("XFD1048576")!.Reference);
        Assert.Equal("XFD1048576", ws.GetRange("XFD1048576", "XFD1048576").Single().Reference);

        var row = ws.GetRow(1_048_576);
        Assert.NotNull(row);
        Assert.Equal("edge", row!.Cells.Single().Value);
    }

    [Fact]
    public void AddTable_AtXfdBoundary_Works()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var worksheet = builder.AddWorksheet("Sheet1");
            worksheet.AddCell("XFD1", "H");
            worksheet.AddDataRow(new List<string> { "v" }, 2);
            worksheet.AddTable("XFD1", "XFD2", "EdgeTable");
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var table = doc.WorkbookPart!.WorksheetParts.First()
            .TableDefinitionParts.Single().Table!;
        Assert.Equal("XFD1:XFD2", table.Reference?.Value);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void AddHeaderRow_SpansFullAXfdWidth()
    {
        // A header row covering the entire A-XFD column range is the strongest
        // coordinate test through a public builder operation: it forces
        // GetCellReference(16383, row) to resolve to XFD on the write path.
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var worksheet = builder.AddWorksheet("Sheet1");
            var headers = Enumerable.Range(0, 16384).Select(i => $"H{i}").ToList();
            worksheet.AddHeaderRow(headers, 1);
            builder.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        var ws = reader.GetWorksheet("Sheet1");
        Assert.Equal("H0", ws.GetCellValue("A1"));
        Assert.Equal("H16383", ws.GetCellValue("XFD1"));
        Assert.Equal(16384, ws.GetRow(1)!.Cells.Count);
        Assert.Throws<XlsxException>(() => ws.AddCell("XFE1", "value"));
    }

    [Theory]
    [InlineData("XFE1")]
    [InlineData("XFE1048576")]
    public void PublicOperations_ColumnBeyondXFD_ThrowXlsxException(string reference)
    {
        // XFE is the first invalid column. Every public operation that parses a cell
        // reference must reject it uniformly — not just the write path.
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var worksheet = builder.AddWorksheet("Sheet1");

        Assert.Throws<XlsxException>(() => worksheet.AddCell(reference, "value"));
        Assert.Throws<XlsxException>(() => worksheet.GetCellValue(reference));
        Assert.Throws<XlsxException>(() => worksheet.CellExists(reference));
        Assert.Throws<XlsxException>(() => worksheet.GetCellInfo(reference));
        Assert.Throws<XlsxException>(() => worksheet.DeleteCell(reference));
        Assert.Throws<XlsxException>(() => worksheet.GetRange(reference, reference));
        Assert.Throws<XlsxException>(() => worksheet.ClearRange(reference, reference));
        Assert.Throws<XlsxException>(() => worksheet.AddTable(reference, reference, "T"));
    }

    [Fact]
    public void GetCellReference_MaxExcelIndexes_ShouldReturnXFD1048576()
    {
        var reference = XlsxEditor.Core.Builders.WorksheetBuilder.GetCellReference(16383, 1048576);
        Assert.Equal("XFD1048576", reference);
    }

    [Fact]
    public void GetCellReference_ColumnBeyondMax_ShouldThrowXlsxException()
    {
        var ex = Assert.Throws<XlsxException>(
            () => XlsxEditor.Core.Builders.WorksheetBuilder.GetCellReference(16384, 1));
        Assert.Contains("XFD", ex.Message);
    }

    [Fact]
    public void GetCellReference_RowBeyondMax_ShouldThrowXlsxException()
    {
        var ex = Assert.Throws<XlsxException>(
            () => XlsxEditor.Core.Builders.WorksheetBuilder.GetCellReference(0, 1048577));
        Assert.Contains("1,048,576", ex.Message);
    }

    [Fact]
    public void GetRange_ReversedRange_ShouldThrowXlsxException()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var worksheet = builder.AddWorksheet("Sheet1");

        var ex = Assert.Throws<XlsxException>(() => worksheet.GetRange("C3", "A1"));
        Assert.Contains("reversed", ex.Message.ToLowerInvariant());
    }

    [Fact]
    public void GetRange_SameRowReversedColumns_ShouldThrowXlsxException()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var worksheet = builder.AddWorksheet("Sheet1");

        Assert.Throws<XlsxException>(() => worksheet.GetRange("B1", "A1"));
    }

    [Fact]
    public void GetRange_SameColumnReversedRows_ShouldThrowXlsxException()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var worksheet = builder.AddWorksheet("Sheet1");

        Assert.Throws<XlsxException>(() => worksheet.GetRange("A3", "A1"));
    }

    [Fact]
    public void GetRange_MalformedReference_ShouldThrowXlsxException()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var worksheet = builder.AddWorksheet("Sheet1");

        var ex = Assert.Throws<XlsxException>(() => worksheet.GetRange("1A", "B2"));
        Assert.Contains("Invalid cell reference", ex.Message);
    }

    [Fact]
    public void ClearRange_ReversedRange_ShouldThrowXlsxException()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var worksheet = builder.AddWorksheet("Sheet1");

        var ex = Assert.Throws<XlsxException>(() => worksheet.ClearRange("B2", "A1"));
        Assert.Contains("reversed", ex.Message.ToLowerInvariant());
    }

    [Fact]
    public void AddCell_LowercaseReference_ShouldNormalizeToSameCell()
    {
        // Act: 'a1' and 'A1' are the same cell in Excel — they must not duplicate
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var worksheet = builder.AddWorksheet("Sheet1");
            worksheet.AddCell("a1", "lower");
            worksheet.AddCell("A1", "upper");
            builder.Save();
        }

        // Assert: one cell, canonical uppercase reference, last write wins
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var cells = doc.WorkbookPart!.WorksheetParts.First().Worksheet!
            .GetFirstChild<SheetData>()!.Elements<Row>().Single().Elements<Cell>().ToList();
        Assert.Single(cells);
        Assert.Equal("A1", cells[0].CellReference?.Value);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void CellLookup_ShouldBeCaseInsensitive()
    {
        // Arrange
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var worksheet = builder.AddWorksheet("Sheet1");
        worksheet.AddCell("b2", "Present");

        // Act & Assert
        Assert.True(worksheet.CellExists("B2"));
        Assert.Equal("Present", worksheet.GetCellValue("B2"));
        Assert.Equal("Present", worksheet.GetCellValue("b2"));
    }

    [Fact]
    public void AddWorksheet_ShouldThrowForDuplicateName_DifferentCase()
    {
        // Arrange: Excel treats sheet names case-insensitively
        using var builder = WorkbookBuilder.Create(_testFilePath);
        builder.AddWorksheet("Sales");

        // Act & Assert
        var ex = Assert.Throws<XlsxException>(() => builder.AddWorksheet("SALES"));
        Assert.Contains("SALES", ex.Message);
        Assert.Contains("already exists", ex.Message);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-3)]
    public void RowOperations_ShouldThrowForRowIndexBelowOne(int rowIndex)
    {
        // Arrange
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var worksheet = builder.AddWorksheet("Sheet1");

        // Act & Assert
        Assert.Throws<XlsxException>(() => worksheet.AddDataRow(new List<string> { "x" }, rowIndex));
        Assert.Throws<XlsxException>(() => worksheet.AddHeaderRow(new List<string> { "x" }, rowIndex));
        Assert.Throws<XlsxException>(() => worksheet.GetRow(rowIndex));
        Assert.Throws<XlsxException>(() => worksheet.DeleteRow(rowIndex));
    }

    [Fact]
    public void RowOperations_MaxValidRow_EmptySheet_ReturnsNullOrNoOp()
    {
        // 1,048,576 is Excel's last valid row: reads of an absent row stay null,
        // deletes of an absent row stay no-ops.
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var worksheet = builder.AddWorksheet("Sheet1");

        Assert.Null(worksheet.GetRow(1_048_576));
        Assert.Same(worksheet, worksheet.DeleteRow(1_048_576));
    }

    [Fact]
    public void RowOperations_FirstInvalidRow_ThrowsXlsxException()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var worksheet = builder.AddWorksheet("Sheet1");

        var getEx = Assert.Throws<XlsxException>(() => worksheet.GetRow(1_048_577));
        Assert.Contains("1,048,576", getEx.Message);

        var deleteEx = Assert.Throws<XlsxException>(() => worksheet.DeleteRow(1_048_577));
        Assert.Contains("1,048,576", deleteEx.Message);
    }

    [Fact]
    public void RowOperations_FirstInvalidRow_Writes_ThrowXlsxException()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var worksheet = builder.AddWorksheet("Sheet1");

        // An empty row/list exercises GetOrCreateRow's own row bound: with no cells the
        // per-cell GetCellReference bound can never fire, so the row check must.
        Assert.Throws<XlsxException>(() => worksheet.AddDataRow(new List<string>(), 1_048_577));
        Assert.Throws<XlsxException>(() => worksheet.AddHeaderRow(new List<string>(), 1_048_577));
        Assert.Throws<XlsxException>(() => worksheet.AddFormulaRow(new List<string>(), 1_048_577));
    }

    [Fact]
    public void RowOperations_MaxValidRow_HoldsData_RoundTrips()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1").AddDataRow(new List<string> { "edge" }, 1_048_576);
            builder.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        var ws = reader.GetWorksheet("Sheet1");
        Assert.Equal("edge", ws.GetCellValue("A1048576"));
        Assert.Equal(1_048_576, ws.GetRow(1_048_576)!.RowIndex);
        Assert.Throws<XlsxException>(() => ws.GetRow(1_048_577));
        Assert.Throws<XlsxException>(() => ws.DeleteRow(1_048_577));
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
        OpenXmlAssert.NoValidationErrors(doc);
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

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void Save_WithEmptyOrWhitespacePath_ThrowsArgumentExceptionAndPreservesSource(string path)
    {
        // Gap fixed: an empty destination was a silent flush of the source document while
        // whitespace materialized a junk file named after the whitespace; both are now rejected
        // as argument errors up front, matching the Create/Open path contract.
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1").AddCell("A1", "Keep");
            var ex = Assert.Throws<ArgumentException>(() => builder.Save(path));
            Assert.Equal("path", ex.ParamName);
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        Assert.Equal("Keep", reader.GetWorksheet("Sheet1").GetCellValue("A1"));
    }

    [Fact]
    public void Save_WithWhitespacePath_CreatesNoJunkFile()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1").AddCell("A1", "Keep");
            Assert.Throws<ArgumentException>(() => builder.Save("   "));
        }

        Assert.False(File.Exists("   "), "A rejected whitespace path must not create a file.");
    }

    [Fact]
    public void Save_AfterRejectedPath_ValidSaveStillWorks()
    {
        // A rejected explicit destination must leave the builder usable: the source flush
        // via the parameterless Save() still works afterwards.
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1").AddCell("A1", "Keep");
            Assert.Throws<ArgumentException>(() => builder.Save("   "));
            builder.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        Assert.Equal("Keep", reader.GetWorksheet("Sheet1").GetCellValue("A1"));
    }

    // ─── Phase 0 tests ────────────────────────────────────────────

    [Fact]
    public void GetOrCreateCell_ShouldInsertCellsInSortedOrder_WhenWrittenOutOfOrder()
    {
        // Act: write B10 first, then A2 — this forces an out-of-order insertion
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var sheet = builder.AddWorksheet("Sheet1");
            sheet.AddCell("B10", "later column");
            sheet.AddCell("A2", "earlier column");
            builder.Save();
        }

        // Assert: row order must be 2, 10; cell order must be A, B
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var sheetData = doc.WorkbookPart!.WorksheetParts.First().Worksheet!.GetFirstChild<SheetData>()!;
        var rows = sheetData.Elements<Row>().ToList();
        Assert.Equal(2, rows.Count);
        Assert.Equal(2U, rows[0].RowIndex!.Value);
        Assert.Equal(10U, rows[1].RowIndex!.Value);

        var cellsRow2 = rows[0].Elements<Cell>().ToList();
        Assert.Single(cellsRow2);
        Assert.Equal("A2", cellsRow2[0].CellReference!.Value);

        var cellsRow10 = rows[1].Elements<Cell>().ToList();
        Assert.Single(cellsRow10);
        Assert.Equal("B10", cellsRow10[0].CellReference!.Value);
    }

    [Fact]
    public void GetOrCreateCell_ShouldInsertCellsInColumnOrder_WhenColumnsWrittenOutOfOrder()
    {
        // Act: write D1, then B1, then A1
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var sheet = builder.AddWorksheet("Sheet1");
            sheet.AddCell("D1", "column D");
            sheet.AddCell("B1", "column B");
            sheet.AddCell("A1", "column A");
            builder.Save();
        }

        // Assert: cells must be ordered A, B, D
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var sheetData = doc.WorkbookPart!.WorksheetParts.First().Worksheet!.GetFirstChild<SheetData>()!;
        var row = sheetData.Elements<Row>().Single();
        var cells = row.Elements<Cell>().Select(c => c.CellReference!.Value!).ToList();
        Assert.Equal(new[] { "A1", "B1", "D1" }, cells);
    }

    [Fact]
    public void AddCell_Formula_ShouldNotSetDataType()
    {
        // Act
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1")
                .AddCell("A1", "=IF(1>0, \"yes\", \"no\")", true);
            builder.Save();
        }

        // Assert: formula cell should have no DataType so Excel infers it
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var cell = doc.WorkbookPart!.WorksheetParts.First().Worksheet!
            .GetFirstChild<SheetData>()!.Elements<Row>().First().Elements<Cell>().First();
        Assert.NotNull(cell.CellFormula);
        Assert.Null(cell.DataType);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void AddCell_StyleId_ShouldThrowForNonIntegerStyleId()
    {
        // Arrange
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var sheet = builder.AddWorksheet("Sheet1");

        // Act & Assert
        var ex = Assert.Throws<XlsxException>(() => sheet.AddCell("A1", "value", "not-a-number"));
        Assert.Contains("not-a-number", ex.Message);
        Assert.Contains("Phase 3", ex.Message);
    }

    [Fact]
    public void AddCell_StyleId_ShouldApplyExistingStyleIndex()
    {
        // Arrange & Act: AddHeaderRow creates the stylesheet (2 cell formats:
        // 0 = default, 1 = bold header), so styleId "1" is a valid reference.
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var sheet = builder.AddWorksheet("Sheet1");
            sheet.AddHeaderRow(new List<string> { "H" });
            sheet.AddCell("A2", "42", "1");
            builder.Save();
        }

        // Assert
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var cell = doc.WorkbookPart!.WorksheetParts.First().Worksheet!
            .GetFirstChild<SheetData>()!.Elements<Row>()
            .Single(r => r.RowIndex?.Value == 2U).Elements<Cell>().First();
        Assert.Equal(1U, cell.StyleIndex?.Value);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void AddCell_StyleId_ShouldThrow_WhenWorkbookHasNoStylesheet()
    {
        // Arrange: a fresh workbook has no WorkbookStylesPart; writing s="5"
        // would point at a nonexistent cellXfs entry (Excel repair prompt).
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var sheet = builder.AddWorksheet("Sheet1");

        // Act & Assert
        var ex = Assert.Throws<XlsxException>(() => sheet.AddCell("A1", "42", "5"));
        Assert.Contains("stylesheet", ex.Message);
    }

    [Fact]
    public void AddCell_StyleId_ShouldThrow_WhenIndexOutOfRange()
    {
        // Arrange: header style gives us 2 cell formats (valid ids 0 and 1)
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var sheet = builder.AddWorksheet("Sheet1");
        sheet.AddHeaderRow(new List<string> { "H" });

        // Act & Assert
        var ex = Assert.Throws<XlsxException>(() => sheet.AddCell("A2", "42", "7"));
        Assert.Contains("7", ex.Message);
        Assert.Contains("0–1", ex.Message);
    }

    [Fact]
    public void AddWorksheet_ShouldThrowForDuplicateName()
    {
        // Arrange
        using var builder = WorkbookBuilder.Create(_testFilePath);
        builder.AddWorksheet("Sales");

        // Act & Assert
        var ex = Assert.Throws<XlsxException>(() => builder.AddWorksheet("Sales"));
        Assert.Contains("Sales", ex.Message);
        Assert.Contains("already exists", ex.Message);
    }

    [Fact]
    public void AddWorksheet_ShouldThrowForNameExceeding31Chars()
    {
        // Arrange
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var longName = new string('A', 32);

        // Act & Assert
        var ex = Assert.Throws<XlsxException>(() => builder.AddWorksheet(longName));
        Assert.Contains("31", ex.Message);
    }

    [Theory]
    [InlineData("Sheet:1")]
    [InlineData("Path/Sheet")]
    [InlineData("Why?Yes")]
    [InlineData("Star*Sheet")]
    [InlineData("Bracket[Sheet")]
    [InlineData("Bracket]Sheet")]
    public void AddWorksheet_ShouldThrowForInvalidCharacters(string name)
    {
        // Arrange
        using var builder = WorkbookBuilder.Create(_testFilePath);

        // Act & Assert
        var ex = Assert.Throws<XlsxException>(() => builder.AddWorksheet(name));
        Assert.Contains("illegal", ex.Message.ToLowerInvariant());
    }

    [Fact]
    public void AddWorksheet_ShouldThrowForEmptyOrWhitespaceName()
    {
        // Arrange
        using var builder = WorkbookBuilder.Create(_testFilePath);

        // Act & Assert
        Assert.Throws<XlsxException>(() => builder.AddWorksheet(""));
        Assert.Throws<XlsxException>(() => builder.AddWorksheet("   "));
    }

    [Theory]
    [InlineData("'Sheet")]
    [InlineData("Sheet'")]
    public void AddWorksheet_ShouldThrowForLeadingOrTrailingApostrophe(string name)
    {
        // Excel rejects sheet names that begin or end with an apostrophe (a leading
        // apostrophe is its R1C1-escape character), so they must fail loudly rather
        // than produce a workbook Excel refuses to open cleanly.
        using var builder = WorkbookBuilder.Create(_testFilePath);

        var ex = Assert.Throws<XlsxException>(() => builder.AddWorksheet(name));
        Assert.Contains("apostrophe", ex.Message);
    }

    [Fact]
    public void AddWorksheet_ShouldAllowApostropheInsideName()
    {
        // Positive control: apostrophes are legal anywhere except the first/last char.
        using var builder = WorkbookBuilder.Create(_testFilePath);

        var ws = builder.AddWorksheet("Sheet's");

        Assert.Equal("Sheet's", builder.GetWorksheetNames().Single());
    }

    [Theory]
    [InlineData(0, "A")]
    [InlineData(25, "Z")]
    [InlineData(26, "AA")]
    [InlineData(51, "AZ")]
    [InlineData(52, "BA")]
    [InlineData(675, "YZ")]
    [InlineData(676, "ZA")]
    [InlineData(701, "ZZ")]
    [InlineData(702, "AAA")]
    [InlineData(728, "ABA")]
    [InlineData(1378, "BAA")]
    [InlineData(1404, "BBA")]
    [InlineData(16383, "XFD")]
    public void GetColumnName_ShouldMatchExcelNaming(int zeroBasedIndex, string expectedName)
    {
        // Act
        var name = XlsxEditor.Core.Builders.WorksheetBuilder.GetColumnName(zeroBasedIndex);

        // Assert
        Assert.Equal(expectedName, name);
    }

    [Fact]
    public void GetColumnName_ShouldBeMonotonicAndValid_ForAllExcelColumns()
    {
        var seen = new HashSet<string>();
        for (int i = 0; i <= 16383; i++)
        {
            var name = XlsxEditor.Core.Builders.WorksheetBuilder.GetColumnName(i);
            Assert.NotNull(name);
            Assert.True(name.Length >= 1 && name.Length <= 3, $"Column {i} gave '{name}'");
            Assert.All(name, c => Assert.True(c >= 'A' && c <= 'Z'));
            Assert.True(seen.Add(name), $"Duplicate column name '{name}' at index {i}");
        }
        // Verify every name was unique across the full range
        Assert.Equal(16384, seen.Count);
        Assert.Contains("A", seen);
        Assert.Contains("XFD", seen);
    }

    [Fact]
    public void AddTable_ShouldAssignDistinctIdsAcrossWorksheets()
    {
        // Act
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var sheet1 = builder.AddWorksheet("Sheet1");
            sheet1.AddHeaderRow(new List<string> { "A", "B" });
            sheet1.AddDataRow(new List<string> { "1", "2" }, 2);
            sheet1.AddTable("A1", "B2", "TableOne");

            var sheet2 = builder.AddWorksheet("Sheet2");
            sheet2.AddHeaderRow(new List<string> { "X", "Y" });
            sheet2.AddDataRow(new List<string> { "3", "4" }, 2);
            sheet2.AddTable("A1", "B2", "TableTwo");

            builder.Save();
        }

        // Assert: table IDs must be distinct (workbook-wide unique)
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var tableParts = doc.WorkbookPart!.WorksheetParts
            .SelectMany(wp => wp.TableDefinitionParts)
            .ToList();
        Assert.Equal(2, tableParts.Count);
        var ids = tableParts.Select(p => p.Table!.Id!.Value).ToList();
        Assert.NotEqual(ids[0], ids[1]);
    }

    [Fact]
    public void AddChart_ShouldThrowXlsxExceptionWithPhase4Reference()
    {
        // Arrange
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var sheet = builder.AddWorksheet("Sheet1");

        // Act & Assert
        var ex = Assert.Throws<XlsxException>(() => sheet.AddChart(ChartType.Column, "A1:D10"));
        Assert.Contains("Phase 4", ex.Message);
    }

    // ─── Phase 1 tests: Read API ──────────────────────────────────

    [Fact]
    public void GetCellValue_ShouldReturnNumericString_ForNumberCell()
    {
        using (var creator = WorkbookBuilder.Create(_testFilePath))
        {
            creator.AddWorksheet("Sheet1").AddCell("A1", "42.5");
            creator.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        Assert.Equal("42.5", reader.GetWorksheet("Sheet1").GetCellValue("A1"));
    }

    [Fact]
    public void GetCellValue_ShouldResolveSharedString_ForStringCell()
    {
        using (var creator = WorkbookBuilder.Create(_testFilePath))
        {
            creator.AddWorksheet("Sheet1").AddCell("A1", "Hello World");
            creator.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        Assert.Equal("Hello World", reader.GetWorksheet("Sheet1").GetCellValue("A1"));
    }

    [Fact]
    public void GetCellValue_ShouldReturnCachedValue_ForFormulaCell()
    {
        using (var creator = WorkbookBuilder.Create(_testFilePath))
        {
            creator.AddWorksheet("Sheet1")
                .AddCell("A1", "10")
                .AddCell("A2", "=A1+5", true);
            creator.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        var ws = reader.GetWorksheet("Sheet1");
        Assert.NotNull(ws.GetCellFormula("A2"));
        Assert.Equal("=A1+5", ws.GetCellFormula("A2"));
    }

    [Fact]
    public void GetCellValue_ShouldReturnNull_ForMissingCell()
    {
        using (var creator = WorkbookBuilder.Create(_testFilePath))
        {
            creator.AddWorksheet("Sheet1").AddCell("A1", "There");
            creator.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        Assert.Null(reader.GetWorksheet("Sheet1").GetCellValue("Z99"));
        Assert.Null(reader.GetWorksheet("Sheet1").GetCellFormula("Z99"));
        Assert.False(reader.GetWorksheet("Sheet1").CellExists("Z99"));
    }

    [Fact]
    public void CellExists_ShouldReturnTrue_ForExistingCell()
    {
        using (var creator = WorkbookBuilder.Create(_testFilePath))
        {
            creator.AddWorksheet("Sheet1").AddCell("B2", "Present");
            creator.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        var ws = reader.GetWorksheet("Sheet1");
        Assert.True(ws.CellExists("B2"));
        Assert.False(ws.CellExists("C3"));
    }

    [Fact]
    public void GetCellInfo_ShouldReturnNull_ForMissingCell()
    {
        using (var creator = WorkbookBuilder.Create(_testFilePath))
        {
            creator.AddWorksheet("Sheet1");
            creator.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        Assert.Null(reader.GetWorksheet("Sheet1").GetCellInfo("J10"));
    }

    [Fact]
    public void GetRange_ShouldReturnAllCells_IncludingEmpty()
    {
        using (var creator = WorkbookBuilder.Create(_testFilePath))
        {
            var sht = creator.AddWorksheet("Sheet1");
            sht.AddCell("A1", "TopLeft");
            sht.AddCell("C1", "TopRight");
            sht.AddCell("A3", "BottomLeft");
            creator.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        var range = reader.GetWorksheet("Sheet1").GetRange("A1", "C3");

        Assert.Equal(9, range.Count);
        Assert.Contains(range, c => c.Reference == "A1" && c.Value == "TopLeft");
        Assert.Contains(range, c => c.Reference == "C1" && c.Value == "TopRight");
        Assert.Contains(range, c => c.Reference == "A3" && c.Value == "BottomLeft");
        Assert.Contains(range, c => c.Reference == "B1" && c.Value == null);
    }

    [Fact]
    public void GetRows_ShouldReturnAllRowsWithCells()
    {
        using (var creator = WorkbookBuilder.Create(_testFilePath))
        {
            creator.AddWorksheet("Sheet1")
                .AddCell("A1", "R1C1")
                .AddCell("B1", "R1C2")
                .AddCell("A2", "R2C1");
            creator.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        var rows = reader.GetWorksheet("Sheet1").GetRows();

        Assert.Equal(2, rows.Count);
        Assert.Equal(1, rows[0].RowIndex);
        Assert.Equal(2, rows[1].RowIndex);
        Assert.Equal(2, rows[0].Cells.Count);
        Assert.Single(rows[1].Cells);
    }

    [Fact]
    public void GetRow_ShouldReturnSpecificRow()
    {
        using (var creator = WorkbookBuilder.Create(_testFilePath))
        {
            var sht = creator.AddWorksheet("Sheet1");
            for (int i = 1; i <= 5; i++)
            {
                sht.AddCell($"A{i}", $"Row{i}");
            }
            creator.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        var row3 = reader.GetWorksheet("Sheet1").GetRow(3);
        var rowMissing = reader.GetWorksheet("Sheet1").GetRow(99);

        Assert.NotNull(row3);
        Assert.Equal(3, row3!.RowIndex);
        Assert.Single(row3.Cells);
        Assert.Null(rowMissing);
    }

    [Fact]
    public void GetDimensions_ShouldComputeUsedRange()
    {
        using (var creator = WorkbookBuilder.Create(_testFilePath))
        {
            var sht = creator.AddWorksheet("Sheet1");
            sht.AddCell("B2", "1");
            sht.AddCell("D5", "2");
            creator.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        var (firstRow, lastRow, firstCol, lastCol) = reader.GetWorksheet("Sheet1").GetDimensions();

        Assert.Equal(2, firstRow);
        Assert.Equal(5, lastRow);
        Assert.Equal(2, firstCol);
        Assert.Equal(4, lastCol);
    }

    [Fact]
    public void GetDimensions_ShouldReturnZeros_ForEmptySheet()
    {
        using (var creator = WorkbookBuilder.Create(_testFilePath))
        {
            creator.AddWorksheet("Sheet1");
            creator.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        var dims = reader.GetWorksheet("Sheet1").GetDimensions();
        Assert.Equal((0, 0, 0, 0), dims);
    }

    // ─── Phase 1 tests: Edit ──────────────────────────────────────

    [Fact]
    public void RoundTrip_BuildReopenReadback_ShouldMatchWrittenValues()
    {
        using (var creator = WorkbookBuilder.Create(_testFilePath))
        {
            var sht = creator.AddWorksheet("Data");
            sht.AddCell("A1", "42");
            sht.AddCell("B1", "Hello");
            sht.AddCell("C1", "=SUM(A1:A1)", true);
            creator.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        var ws = reader.GetWorksheet("Data");
        Assert.Equal("42", ws.GetCellValue("A1"));
        Assert.Equal("Hello", ws.GetCellValue("B1"));
        Assert.Equal("=SUM(A1:A1)", ws.GetCellFormula("C1"));
    }

    [Fact]
    public void OverwriteCell_ShouldReplaceExistingValue()
    {
        using (var creator = WorkbookBuilder.Create(_testFilePath))
        {
            creator.AddWorksheet("Sheet1").AddCell("A1", "Initial");
            creator.Save();
        }

        using (var editor = WorkbookBuilder.Open(_testFilePath))
        {
            editor.GetWorksheet("Sheet1").AddCell("A1", "Updated");
            editor.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        Assert.Equal("Updated", reader.GetWorksheet("Sheet1").GetCellValue("A1"));
    }

    [Fact]
    public void OverwriteFormulaCell_WithValue_ShouldClearFormula()
    {
        // Act: write a formula, then overwrite the same cell with a literal value
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var sheet = builder.AddWorksheet("Sheet1");
            sheet.AddCell("A1", "=SUM(B1:B2)", true);
            sheet.AddCell("A1", "42");
            builder.Save();
        }

        // Assert: no stale <f> survives; the cell is a plain number
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var cell = doc.WorkbookPart!.WorksheetParts.First().Worksheet!
            .GetFirstChild<SheetData>()!.Elements<Row>().Single().Elements<Cell>().Single();
        Assert.Null(cell.CellFormula);
        Assert.Equal(CellValues.Number, cell.DataType?.Value);
        Assert.Equal("42", cell.CellValue?.Text);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void OverwriteValueCell_WithFormula_ShouldClearValueAndDataType()
    {
        // Act: write a value, then overwrite the same cell with a formula
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var sheet = builder.AddWorksheet("Sheet1");
            sheet.AddCell("A1", "99");
            sheet.AddCell("A1", "=SUM(B1:B2)", true);
            builder.Save();
        }

        // Assert: the stale cached value and data type are gone — Excel
        // recalculates on open instead of trusting a value we know is wrong.
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var cell = doc.WorkbookPart!.WorksheetParts.First().Worksheet!
            .GetFirstChild<SheetData>()!.Elements<Row>().Single().Elements<Cell>().Single();
        Assert.NotNull(cell.CellFormula);
        Assert.Null(cell.CellValue);
        Assert.Null(cell.DataType);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void OverwriteFormulaCell_WithValue_RoundTrip_ShouldNotReturnStaleState()
    {
        // Arrange: a saved formula cell
        using (var creator = WorkbookBuilder.Create(_testFilePath))
        {
            creator.AddWorksheet("Sheet1").AddCell("A1", "=1+1", true);
            creator.Save();
        }

        // Act: reopen and overwrite with a value
        using (var editor = WorkbookBuilder.Open(_testFilePath))
        {
            editor.GetWorksheet("Sheet1").AddCell("A1", "done");
            editor.Save();
        }

        // Assert
        using var reader = WorkbookBuilder.Open(_testFilePath);
        var ws = reader.GetWorksheet("Sheet1");
        Assert.Null(ws.GetCellFormula("A1"));
        Assert.Equal("done", ws.GetCellValue("A1"));
    }

    [Fact]
    public void GetCellValue_FormulaCellWithoutCachedValue_ShouldReturnNull()
    {
        // Arrange: freshly written formula cells have no cached value
        using (var creator = WorkbookBuilder.Create(_testFilePath))
        {
            creator.AddWorksheet("Sheet1").AddCell("A1", "=1+1", true);
            creator.Save();
        }

        // Assert: GetCellValue must not fabricate or return stale content
        using var reader = WorkbookBuilder.Open(_testFilePath);
        Assert.Null(reader.GetWorksheet("Sheet1").GetCellValue("A1"));
    }

    [Fact]
    public void DeleteCell_ShouldRemoveCell()
    {
        using (var creator = WorkbookBuilder.Create(_testFilePath))
        {
            var sht = creator.AddWorksheet("Sheet1");
            sht.AddCell("A1", "one");
            sht.AddCell("B1", "two");
            sht.AddCell("C1", "three");
            creator.Save();
        }

        using (var editor = WorkbookBuilder.Open(_testFilePath))
        {
            var sht = editor.GetWorksheet("Sheet1");
            sht.DeleteCell("B1");
            editor.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        var ws = reader.GetWorksheet("Sheet1");
        Assert.True(ws.CellExists("A1"));
        Assert.False(ws.CellExists("B1"));
        Assert.True(ws.CellExists("C1"));
    }

    [Fact]
    public void DeleteRow_ShouldRemoveRowAndCells()
    {
        using (var creator = WorkbookBuilder.Create(_testFilePath))
        {
            var sht = creator.AddWorksheet("Sheet1");
            sht.AddCell("A1", "Row1");
            sht.AddCell("A2", "Row2");
            sht.AddCell("A3", "Row3");
            creator.Save();
        }

        using (var editor = WorkbookBuilder.Open(_testFilePath))
        {
            editor.GetWorksheet("Sheet1").DeleteRow(2);
            editor.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        var ws = reader.GetWorksheet("Sheet1");
        Assert.True(ws.CellExists("A1"));
        Assert.False(ws.CellExists("A2"));
        Assert.True(ws.CellExists("A3"));
        var rows = ws.GetRows();
        Assert.Equal(2, rows.Count);
        Assert.DoesNotContain(rows, r => r.RowIndex == 2);
    }

    [Fact]
    public void ClearRange_ShouldRemoveAllCellsInRange()
    {
        using (var creator = WorkbookBuilder.Create(_testFilePath))
        {
            var sht = creator.AddWorksheet("Sheet1");
            sht.AddCell("A1", "1");
            sht.AddCell("B1", "2");
            sht.AddCell("A2", "3");
            sht.AddCell("B2", "4");
            sht.AddCell("C3", "keep");
            creator.Save();
        }

        using (var editor = WorkbookBuilder.Open(_testFilePath))
        {
            editor.GetWorksheet("Sheet1").ClearRange("A1", "B2");
            editor.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        var ws = reader.GetWorksheet("Sheet1");
        Assert.False(ws.CellExists("A1"));
        Assert.False(ws.CellExists("B1"));
        Assert.False(ws.CellExists("A2"));
        Assert.False(ws.CellExists("B2"));
        Assert.True(ws.CellExists("C3"));
        Assert.Equal("keep", ws.GetCellValue("C3"));
    }

    [Fact]
    public void OpenExisting_ThenEditOneCell_UntouchedCellsPreserved()
    {
        using (var creator = WorkbookBuilder.Create(_testFilePath))
        {
            var sht = creator.AddWorksheet("Sheet1");
            sht.AddCell("A1", "Untouched");
            sht.AddCell("B1", "original");
            sht.AddCell("C1", "42");
            creator.Save();
        }

        using (var editor = WorkbookBuilder.Open(_testFilePath))
        {
            editor.GetWorksheet("Sheet1").AddCell("B1", "modified");
            editor.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        var ws = reader.GetWorksheet("Sheet1");
        Assert.Equal("Untouched", ws.GetCellValue("A1"));
        Assert.Equal("modified", ws.GetCellValue("B1"));
        Assert.Equal("42", ws.GetCellValue("C1"));
    }

    // ─── Explicit column widths and row heights ──────────────────

    [Fact]
    public void SetColumnWidth_ShouldWriteSingleColDefinition()
    {
        // Act
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var sheet = builder.AddWorksheet("Sheet1");
            Assert.Same(sheet, sheet.SetColumnWidth("B", 20.5));
            builder.Save();
        }

        // Assert: <col min="2" max="2" width="20.5" customWidth="1"/>
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var col = Assert.Single(doc.WorkbookPart!.WorksheetParts.First().Worksheet!
            .GetFirstChild<Columns>()!.Elements<Column>());
        Assert.Equal(2u, col.Min!.Value);
        Assert.Equal(2u, col.Max!.Value);
        Assert.Equal(20.5, col.Width!.Value);
        Assert.True(col.CustomWidth!.Value);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void SetColumnWidth_MultipleColumnsOutOfOrder_ShouldBeSortedWithoutOverlap()
    {
        // Act: write C, then A, then B — forced out-of-order insertion
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var sheet = builder.AddWorksheet("Sheet1");
            sheet.SetColumnWidth("C", 30);
            sheet.SetColumnWidth("A", 10);
            sheet.SetColumnWidth("B", 20);
            builder.Save();
        }

        // Assert: one sorted, non-overlapping definition per column
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var cols = doc.WorkbookPart!.WorksheetParts.First().Worksheet!
            .GetFirstChild<Columns>()!.Elements<Column>().ToList();
        Assert.Equal(3, cols.Count);
        Assert.Equal(new[] { 1u, 2u, 3u }, cols.Select(c => c.Min!.Value));
        Assert.Equal(new[] { 1u, 2u, 3u }, cols.Select(c => c.Max!.Value));
        Assert.Equal(new[] { 10.0, 20.0, 30.0 }, cols.Select(c => c.Width!.Value));
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void SetColumnWidth_RepeatedCall_UpdatesInPlace_WithoutDuplicates()
    {
        // Act: the same column is set three times, interleaved with another column
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var sheet = builder.AddWorksheet("Sheet1");
            sheet.SetColumnWidth("A", 10);
            sheet.SetColumnWidth("A", 30);
            sheet.SetColumnWidth("B", 15);
            sheet.SetColumnWidth("A", 40);
            builder.Save();
        }

        // Assert: A appears exactly once, last width wins
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var cols = doc.WorkbookPart!.WorksheetParts.First().Worksheet!
            .GetFirstChild<Columns>()!.Elements<Column>().ToList();
        Assert.Equal(2, cols.Count);
        var a = cols.Single(c => c.Min!.Value == 1);
        Assert.Equal(1u, a.Max!.Value);
        Assert.Equal(40.0, a.Width!.Value);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void SetColumnWidth_InsertBetweenExistingColumns_ShouldPreserveOthers()
    {
        // Act: A and C exist before B is inserted between them
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var sheet = builder.AddWorksheet("Sheet1");
            sheet.SetColumnWidth("A", 10);
            sheet.SetColumnWidth("C", 30);
            sheet.SetColumnWidth("B", 20);
            builder.Save();
        }

        // Assert: all three survive with their own widths
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var cols = doc.WorkbookPart!.WorksheetParts.First().Worksheet!
            .GetFirstChild<Columns>()!.Elements<Column>()
            .ToDictionary(c => c.Min!.Value, c => c.Width!.Value);
        Assert.Equal(3, cols.Count);
        Assert.Equal(new[] { 1u, 2u, 3u }, cols.Keys.OrderBy(k => k));
        Assert.Equal(10.0, cols[1]);
        Assert.Equal(20.0, cols[2]);
        Assert.Equal(30.0, cols[3]);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void SetColumnWidth_ShouldPreserveExistingCells()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var sheet = builder.AddWorksheet("Sheet1");
            sheet.AddCell("A1", "value");
            sheet.SetColumnWidth("A", 25);
            builder.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        var ws = reader.GetWorksheet("Sheet1");
        Assert.Equal("value", ws.GetCellValue("A1"));
        Assert.Equal(25.0, ws.GetColumnWidth("A"));
    }

    [Fact]
    public void SetColumnWidth_ColumnInsideExistingRange_ShouldSplitPreservingOthers()
    {
        // A workbook authored with a multi-column definition (as Excel writes when a range
        // of columns shares one width). Setting one column must split the range instead of
        // emitting an overlapping <col> that would corrupt the column model.
        var bytes = CreateWorkbookWithColumnRange();
        byte[] saved;
        using (var builder = WorkbookBuilder.Open(bytes))
        {
            builder.GetWorksheet("Sheet1").SetColumnWidth("B", 20);
            saved = builder.SaveToBytes();
        }

        // Assert: 1..1=10, 2..2=20, 3..5=10, 7..7=22 — widths preserved around the split
        using var reader = WorkbookBuilder.Open(saved);
        var ws = reader.GetWorksheet("Sheet1");
        Assert.Equal(10.0, ws.GetColumnWidth("A"));
        Assert.Equal(20.0, ws.GetColumnWidth("B"));
        Assert.Equal(10.0, ws.GetColumnWidth("D"));
        Assert.Equal(22.0, ws.GetColumnWidth("G"));
        Assert.Null(ws.GetColumnWidth("Z"));

        using var doc = SpreadsheetDocument.Open(new MemoryStream(saved), false);
        var cols = doc.WorkbookPart!.WorksheetParts.First().Worksheet!
            .GetFirstChild<Columns>()!.Elements<Column>().ToList();
        Assert.Equal(4, cols.Count);
        Assert.Equal(new[] { 1u, 2u, 3u, 7u }, cols.Select(c => c.Min!.Value));
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Theory]
    [InlineData(0.0)]
    [InlineData(-5.0)]
    [InlineData(256.0)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    public void SetColumnWidth_InvalidWidth_ThrowsXlsxException_AndLeavesWorksheetUnmodified(double width)
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var sheet = builder.AddWorksheet("Sheet1");

        // Act & Assert: validation runs before any mutation, so no <col> is written
        var ex = Assert.Throws<XlsxException>(() => sheet.SetColumnWidth("A", width));
        Assert.Contains("width", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Null(sheet.GetColumnWidth("A"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    [InlineData("A1")]
    [InlineData("AAAA")]
    [InlineData("XFE")]
    [InlineData("1")]
    public void SetColumnWidth_InvalidColumn_ThrowsXlsxException(string? column)
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var sheet = builder.AddWorksheet("Sheet1");

        var ex = Assert.Throws<XlsxException>(() => sheet.SetColumnWidth(column!, 20));
        Assert.Contains("column", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Null(sheet.GetColumnWidth("A"));
    }

    [Theory]
    [InlineData(255.0)]
    [InlineData(1.0)]
    [InlineData(0.5)]
    public void SetColumnWidth_ValidBoundaryWidths_ShouldBeAccepted(double width)
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var sheet = builder.AddWorksheet("Sheet1");

        sheet.SetColumnWidth("A", width);

        Assert.Equal(width, sheet.GetColumnWidth("A"));
    }

    [Fact]
    public void SetColumnWidth_LowercaseAndMaxColumn_ShouldNormalizeAndSucceed()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var sheet = builder.AddWorksheet("Sheet1");

        sheet.SetColumnWidth("a", 12.5);
        sheet.SetColumnWidth("XFD", 255.0);

        Assert.Equal(12.5, sheet.GetColumnWidth("A"));
        Assert.Equal(12.5, sheet.GetColumnWidth("a"));
        Assert.Equal(255.0, sheet.GetColumnWidth("XFD"));
    }

    [Fact]
    public void SetRowHeight_ShouldWriteHtAndCustomHeight()
    {
        // Act
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var sheet = builder.AddWorksheet("Sheet1");
            Assert.Same(sheet, sheet.SetRowHeight(2, 30));
            builder.Save();
        }

        // Assert: <row r="2" ht="30" customHeight="1"/>
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var row = Assert.Single(doc.WorkbookPart!.WorksheetParts.First().Worksheet!
            .GetFirstChild<SheetData>()!.Elements<Row>());
        Assert.Equal(2u, row.RowIndex!.Value);
        Assert.Equal(30.0, row.Height!.Value);
        Assert.True(row.CustomHeight!.Value);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void SetRowHeight_RepeatedCall_UpdatesInPlace()
    {
        // Act: row 1 is set twice, interleaved with another row
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var sheet = builder.AddWorksheet("Sheet1");
            sheet.SetRowHeight(1, 20);
            sheet.SetRowHeight(2, 25);
            sheet.SetRowHeight(1, 40);
            builder.Save();
        }

        // Assert: one row element per index, last height wins
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var rows = doc.WorkbookPart!.WorksheetParts.First().Worksheet!
            .GetFirstChild<SheetData>()!.Elements<Row>().ToList();
        Assert.Equal(2, rows.Count);
        Assert.Equal(40.0, rows.Single(r => r.RowIndex!.Value == 1).Height!.Value);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void SetRowHeight_ShouldPreserveExistingCellsAndStyles()
    {
        // Act: set heights on rows that already carry header/data cells
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var sheet = builder.AddWorksheet("Sheet1");
            sheet.AddHeaderRow(new List<string> { "Name", "City" });
            sheet.AddDataRow(new List<string> { "Ada", "Paris" }, 2);
            sheet.SetRowHeight(1, 30);
            sheet.SetRowHeight(2, 45);
            builder.Save();
        }

        // Assert: cells and header styles survive the height write
        using (var doc = SpreadsheetDocument.Open(_testFilePath, false))
        {
            var rows = doc.WorkbookPart!.WorksheetParts.First().Worksheet!
                .GetFirstChild<SheetData>()!.Elements<Row>().ToList();
            var header = rows.Single(r => r.RowIndex!.Value == 1);
            Assert.Equal(30.0, header.Height!.Value);
            var headerCells = header.Elements<Cell>().ToList();
            Assert.Equal(2, headerCells.Count);
            Assert.All(headerCells, c => Assert.Equal(1u, c.StyleIndex!.Value));

            var data = rows.Single(r => r.RowIndex!.Value == 2);
            Assert.Equal(45.0, data.Height!.Value);
            Assert.Equal(new[] { "A2", "B2" }, data.Elements<Cell>()
                .Select(c => c.CellReference!.Value));
            OpenXmlAssert.NoValidationErrors(doc);
        }

        // Values still resolve through the public API after the height writes
        using var reader = WorkbookBuilder.Open(_testFilePath);
        var ws = reader.GetWorksheet("Sheet1");
        Assert.Equal("Ada", ws.GetCellValue("A2"));
        Assert.Equal("Paris", ws.GetCellValue("B2"));
    }

    [Fact]
    public void SetRowHeight_ThenAddCell_ShouldCoexist()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var sheet = builder.AddWorksheet("Sheet1");
            sheet.SetRowHeight(3, 40);
            sheet.AddCell("A3", "value");
            builder.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        var ws = reader.GetWorksheet("Sheet1");
        Assert.Equal(40.0, ws.GetRowHeight(3));
        Assert.Equal("value", ws.GetCellValue("A3"));
    }

    [Theory]
    [InlineData(0.0)]
    [InlineData(-1.0)]
    [InlineData(409.6)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    public void SetRowHeight_InvalidHeight_ThrowsXlsxException_AndLeavesWorksheetUnmodified(double height)
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var sheet = builder.AddWorksheet("Sheet1");

        var ex = Assert.Throws<XlsxException>(() => sheet.SetRowHeight(1, height));
        Assert.Contains("height", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Null(sheet.GetRowHeight(1));
        Assert.Empty(sheet.GetRows());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(1_048_577)]
    public void SetRowHeight_InvalidRowIndex_ThrowsXlsxException(int rowIndex)
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var sheet = builder.AddWorksheet("Sheet1");

        var ex = Assert.Throws<XlsxException>(() => sheet.SetRowHeight(rowIndex, 30));
        Assert.Contains("row", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Empty(sheet.GetRows());
    }

    [Theory]
    [InlineData(1, 409.5)]
    [InlineData(1_048_576, 1.0)]
    public void SetRowHeight_BoundaryValues_ShouldSucceed(int rowIndex, double height)
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var sheet = builder.AddWorksheet("Sheet1");

        sheet.SetRowHeight(rowIndex, height);

        Assert.Equal(height, sheet.GetRowHeight(rowIndex));
    }

    [Fact]
    public void ColumnWidthsAndRowHeights_PersistAcrossSaveAndReopen()
    {
        // Arrange & Act: build a workbook with widths, heights and cells
        byte[] bytes;
        using (var builder = WorkbookBuilder.Create())
        {
            var sheet = builder.AddWorksheet("Sheet1");
            sheet.SetColumnWidth("A", 20);
            sheet.SetColumnWidth("C", 15.5);
            sheet.SetRowHeight(1, 30);
            sheet.SetRowHeight(2, 45);
            sheet.AddCell("A1", "keep");
            bytes = builder.SaveToBytes();
        }

        // Assert: everything reads back through the public API after a save/reopen
        using (var reader = WorkbookBuilder.Open(bytes))
        {
            var ws = reader.GetWorksheet("Sheet1");
            Assert.Equal(20.0, ws.GetColumnWidth("A"));
            Assert.Null(ws.GetColumnWidth("B"));
            Assert.Equal(15.5, ws.GetColumnWidth("C"));
            Assert.Equal(30.0, ws.GetRowHeight(1));
            Assert.Equal(45.0, ws.GetRowHeight(2));
            Assert.Null(ws.GetRowHeight(3));
            Assert.Equal("keep", ws.GetCellValue("A1"));
        }

        using var doc = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void ColumnWidthsAndRowHeights_ReopenedWorkbook_UpdateInPlaceWithoutDuplicates()
    {
        // Arrange: a saved workbook with widths and heights
        byte[] bytes;
        using (var builder = WorkbookBuilder.Create())
        {
            var sheet = builder.AddWorksheet("Sheet1");
            sheet.SetColumnWidth("A", 10);
            sheet.SetColumnWidth("B", 20);
            sheet.SetRowHeight(1, 25);
            bytes = builder.SaveToBytes();
        }

        // Act: reopen, overwrite every value, add a new column
        using (var builder = WorkbookBuilder.Open(bytes))
        {
            var sheet = builder.GetWorksheet("Sheet1");
            sheet.SetColumnWidth("A", 50);
            sheet.SetColumnWidth("B", 60);
            sheet.SetColumnWidth("C", 70);
            sheet.SetRowHeight(1, 55);
            bytes = builder.SaveToBytes();
        }

        // Assert: no duplicate <col> or <row> definitions appeared
        using (var doc = SpreadsheetDocument.Open(new MemoryStream(bytes), false))
        {
            var cols = doc.WorkbookPart!.WorksheetParts.First().Worksheet!
                .GetFirstChild<Columns>()!.Elements<Column>().ToList();
            Assert.Equal(3, cols.Count);
            var rows = doc.WorkbookPart!.WorksheetParts.First().Worksheet!
                .GetFirstChild<SheetData>()!.Elements<Row>().ToList();
            Assert.Single(rows);
            Assert.Equal(55.0, rows[0].Height!.Value);
            OpenXmlAssert.NoValidationErrors(doc);
        }

        using var reader = WorkbookBuilder.Open(bytes);
        var ws = reader.GetWorksheet("Sheet1");
        Assert.Equal(50.0, ws.GetColumnWidth("A"));
        Assert.Equal(60.0, ws.GetColumnWidth("B"));
        Assert.Equal(70.0, ws.GetColumnWidth("C"));
        Assert.Equal(55.0, ws.GetRowHeight(1));
    }

    /// <summary>
    /// Builds a workbook whose first sheet carries Excel-style multi-column width
    /// definitions (1..5 at width 10, column 7 at width 22).
    /// </summary>
    private static byte[] CreateWorkbookWithColumnRange()
    {
        using var ms = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(
                new Columns(
                    new Column { Min = 1u, Max = 5u, Width = 10, CustomWidth = true },
                    new Column { Min = 7u, Max = 7u, Width = 22, CustomWidth = true }),
                new SheetData());
            workbookPart.Workbook = new Workbook(new Sheets(new Sheet
            {
                Name = "Sheet1",
                SheetId = 1u,
                Id = workbookPart.GetIdOfPart(worksheetPart)
            }));
            doc.Save();
        }
        return ms.ToArray();
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
