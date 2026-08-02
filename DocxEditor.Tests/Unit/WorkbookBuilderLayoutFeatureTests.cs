using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxEditor.Core.Builders;
using XlsxEditor.Core.Exceptions;

namespace DocxEditor.Tests.Unit;

public class WorkbookBuilderLayoutFeatureTests : IDisposable
{
    private readonly string _testFilePath = Path.Combine(Path.GetTempPath(), $"test_xlsx_features_{Guid.NewGuid()}.xlsx");

    // ─── Freeze panes ──────────────────────────────────────────────

    [Fact]
    public void FreezePanes_TopRow_WritesFrozenPaneInSchemaOrder()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var sheet = builder.AddWorksheet("Sheet1");
            Assert.Same(sheet, sheet.FreezePanes(1, 0));
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var worksheet = doc.WorkbookPart!.WorksheetParts.First().Worksheet!;
        var sheetView = worksheet.GetFirstChild<SheetViews>()!.Elements<SheetView>().Single();
        var pane = sheetView.Pane;
        Assert.NotNull(pane);
        Assert.Equal(1.0, pane!.VerticalSplit?.Value);
        Assert.Equal(0.0, pane.HorizontalSplit?.Value);
        Assert.Equal("A2", pane.TopLeftCell?.Value);
        Assert.Equal(PaneValues.BottomLeft, pane.ActivePane?.Value);
        Assert.Equal(PaneStateValues.Frozen, pane.State?.Value);

        var selection = Assert.Single(sheetView.Elements<Selection>());
        Assert.Equal(PaneValues.BottomLeft, selection.Pane?.Value);

        // sheetViews must precede sheetData in the CT_Worksheet sequence
        var childNames = worksheet.ChildElements.Select(c => c.LocalName).ToList();
        Assert.True(childNames.IndexOf("sheetViews") < childNames.IndexOf("sheetData"), $"got {string.Join(",", childNames)}");
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void FreezePanes_FirstColumn_WritesTopRightPane()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1").FreezePanes(0, 1);
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var pane = doc.WorkbookPart!.WorksheetParts.First().Worksheet!
            .GetFirstChild<SheetViews>()!.Elements<SheetView>().Single().Pane!;
        Assert.Equal(1.0, pane.HorizontalSplit?.Value);
        Assert.Equal(0.0, pane.VerticalSplit?.Value);
        Assert.Equal("B1", pane.TopLeftCell?.Value);
        Assert.Equal(PaneValues.TopRight, pane.ActivePane?.Value);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void FreezePanes_BothRowsAndColumns_WritesBottomRightPane()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1").FreezePanes(2, 3);
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var pane = doc.WorkbookPart!.WorksheetParts.First().Worksheet!
            .GetFirstChild<SheetViews>()!.Elements<SheetView>().Single().Pane!;
        Assert.Equal(2.0, pane.VerticalSplit?.Value);
        Assert.Equal(3.0, pane.HorizontalSplit?.Value);
        Assert.Equal("D3", pane.TopLeftCell?.Value);
        Assert.Equal(PaneValues.BottomRight, pane.ActivePane?.Value);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Theory]
    [InlineData(-1, 0)]
    [InlineData(0, -1)]
    public void FreezePanes_NegativeDimension_Throws(int rows, int cols)
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var sheet = builder.AddWorksheet("Sheet1");
        var ex = Assert.Throws<XlsxException>(() => sheet.FreezePanes(rows, cols));
        Assert.Contains("non-negative", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Null(sheet.GetFreezePanes());
    }

    [Fact]
    public void FreezePanes_BothZero_Throws()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var sheet = builder.AddWorksheet("Sheet1");
        var ex = Assert.Throws<XlsxException>(() => sheet.FreezePanes(0, 0));
        Assert.Contains("at least one", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Null(sheet.GetFreezePanes());
    }

    [Theory]
    [InlineData(1_048_577, 0)]
    [InlineData(0, 16385)]
    public void FreezePanes_OutOfBounds_Throws(int rows, int cols)
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var sheet = builder.AddWorksheet("Sheet1");
        var ex = Assert.Throws<XlsxException>(() => sheet.FreezePanes(rows, cols));
        Assert.Contains("out of range", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Null(sheet.GetFreezePanes());
    }

    [Fact]
    public void FreezePanes_RepeatedCall_UpdatesInPlace_WithoutDuplicates()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var sheet = builder.AddWorksheet("Sheet1");
            sheet.FreezePanes(1, 0);
            sheet.FreezePanes(3, 2);
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var sheetViews = doc.WorkbookPart!.WorksheetParts.First().Worksheet!.GetFirstChild<SheetViews>()!;
        Assert.Single(sheetViews.Elements<SheetView>());
        var pane = sheetViews.Elements<SheetView>().Single().Pane!;
        Assert.Equal(3.0, pane.VerticalSplit?.Value);
        Assert.Equal(2.0, pane.HorizontalSplit?.Value);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void FreezePanes_RoundTrip_ThroughReopen()
    {
        byte[] bytes;
        using (var builder = WorkbookBuilder.Create())
        {
            builder.AddWorksheet("Sheet1").FreezePanes(2, 1);
            bytes = builder.SaveToBytes();
        }

        using var reader = WorkbookBuilder.Open(bytes);
        Assert.Equal((2, 1), reader.GetWorksheet("Sheet1").GetFreezePanes());
    }

    [Fact]
    public void GetFreezePanes_ReturnsNull_WhenNoFrozenPane()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var sheet = builder.AddWorksheet("Sheet1");
        Assert.Null(sheet.GetFreezePanes());
    }

    // ─── Standalone autofilter ─────────────────────────────────────

    [Fact]
    public void SetAutoFilter_WritesAutoFilterInSchemaOrder()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var sheet = builder.AddWorksheet("Sheet1");
            Assert.Same(sheet, sheet.SetAutoFilter("A1:D10"));
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var worksheet = doc.WorkbookPart!.WorksheetParts.First().Worksheet!;
        var autoFilter = worksheet.GetFirstChild<AutoFilter>();
        Assert.NotNull(autoFilter);
        Assert.Equal("A1:D10", autoFilter!.Reference?.Value);

        // autoFilter must follow sheetData and precede mergeCells/tableParts
        var childNames = worksheet.ChildElements.Select(c => c.LocalName).ToList();
        Assert.True(childNames.IndexOf("sheetData") < childNames.IndexOf("autoFilter"), $"got {string.Join(",", childNames)}");
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void SetAutoFilter_LowercaseRange_NormalizesToUpperCase()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1").SetAutoFilter("a1:d10");
            builder.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        Assert.Equal("A1:D10", reader.GetWorksheet("Sheet1").GetAutoFilterRange());
    }

    [Fact]
    public void SetAutoFilter_ReversedRange_Throws()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var sheet = builder.AddWorksheet("Sheet1");
        var ex = Assert.Throws<XlsxException>(() => sheet.SetAutoFilter("D10:A1"));
        Assert.Contains("reversed", ex.Message);
        Assert.Null(sheet.GetAutoFilterRange());
    }

    [Fact]
    public void SetAutoFilter_MalformedRange_Throws()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var sheet = builder.AddWorksheet("Sheet1");
        Assert.Throws<XlsxException>(() => sheet.SetAutoFilter("A1"));
        Assert.Null(sheet.GetAutoFilterRange());
    }

    [Fact]
    public void SetAutoFilter_RepeatedCall_ReplacesInPlace()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var sheet = builder.AddWorksheet("Sheet1");
            sheet.SetAutoFilter("A1:C5");
            sheet.SetAutoFilter("B2:D20");
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var worksheet = doc.WorkbookPart!.WorksheetParts.First().Worksheet!;
        Assert.Single(worksheet.Elements<AutoFilter>());
        Assert.Equal("B2:D20", worksheet.GetFirstChild<AutoFilter>()!.Reference?.Value);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void RemoveAutoFilter_RemovesStandaloneFilter()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var sheet = builder.AddWorksheet("Sheet1");
            sheet.SetAutoFilter("A1:C5");
            Assert.Same(sheet, sheet.RemoveAutoFilter());
            builder.Save();
        }

        using (var doc = SpreadsheetDocument.Open(_testFilePath, false))
        {
            Assert.Null(doc.WorkbookPart!.WorksheetParts.First().Worksheet!.GetFirstChild<AutoFilter>());
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        Assert.Null(reader.GetWorksheet("Sheet1").GetAutoFilterRange());
    }

    // ─── Freeze + autofilter + merge + table coexist ───────────────

    [Fact]
    public void LayoutFeatures_CoexistInSchemaOrder_AndValidate()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var sheet = builder.AddWorksheet("Sheet1");
            sheet.AddHeaderRow(new List<string> { "A", "B", "C" });
            sheet.AddDataRow(new List<string> { "1", "2", "3" }, 2);
            sheet.FreezePanes(1, 0);
            sheet.SetAutoFilter("A1:C2");
            sheet.MergeCells("A3:C3");
            sheet.AddTable("A1", "C2", "SalesTable");
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var worksheet = doc.WorkbookPart!.WorksheetParts.First().Worksheet!;
        var childNames = worksheet.ChildElements.Select(c => c.LocalName).ToList();
        var order = childNames.ToDictionary(n => n, n => childNames.IndexOf(n));
        Assert.True(order["sheetViews"] < order["sheetData"], $"got {string.Join(",", childNames)}");
        Assert.True(order["sheetData"] < order["autoFilter"], $"got {string.Join(",", childNames)}");
        Assert.True(order["autoFilter"] < order["mergeCells"], $"got {string.Join(",", childNames)}");
        Assert.True(order["mergeCells"] < order["tableParts"], $"got {string.Join(",", childNames)}");
        OpenXmlAssert.NoValidationErrors(doc);
    }

    // ─── Core metadata ─────────────────────────────────────────────

    [Fact]
    public void SetCoreProperties_WritesAllFields_RoundTrips()
    {
        var created = new DateTime(2024, 1, 15, 10, 30, 0, DateTimeKind.Utc);
        var modified = new DateTime(2024, 6, 1, 8, 0, 0, DateTimeKind.Utc);

        byte[] bytes;
        using (var builder = WorkbookBuilder.Create())
        {
            var properties = new WorkbookCoreProperties
            {
                Title = "Northwind Sales",
                Subject = "Quarterly report",
                Creator = "Ada Lovelace",
                Keywords = "sales, quarterly, northwind",
                Description = "Q1 sales figures and variance.",
                Category = "Reports",
                LastModifiedBy = "Grace Hopper",
                Created = created,
                Modified = modified
            };
            Assert.Same(builder, builder.SetCoreProperties(properties));
            bytes = builder.SaveToBytes();
        }

        using (var reader = WorkbookBuilder.Open(bytes))
        {
            var props = reader.GetCoreProperties();
            Assert.Equal("Northwind Sales", props.Title);
            Assert.Equal("Quarterly report", props.Subject);
            Assert.Equal("Ada Lovelace", props.Creator);
            Assert.Equal("sales, quarterly, northwind", props.Keywords);
            Assert.Equal("Q1 sales figures and variance.", props.Description);
            Assert.Equal("Reports", props.Category);
            Assert.Equal("Grace Hopper", props.LastModifiedBy);
            // PackageProperties stores datetimes in UTC; the read-back may carry a local
            // offset, so compare the instants rather than the DateTime Kind.
            Assert.Equal(created, props.Created!.Value.ToUniversalTime());
            Assert.Equal(modified, props.Modified!.Value.ToUniversalTime());
        }
    }

    [Fact]
    public void SetCoreProperties_NullValues_ClearFields()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.SetCoreProperties(new WorkbookCoreProperties { Title = "Set", Creator = "Someone" });
            builder.Save();
        }

        using (var builder = WorkbookBuilder.Open(_testFilePath))
        {
            builder.SetCoreProperties(new WorkbookCoreProperties());
            builder.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        var props = reader.GetCoreProperties();
        Assert.Null(props.Title);
        Assert.Null(props.Creator);
    }

    [Fact]
    public void SetCoreProperties_NullRecord_Throws()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        Assert.Throws<ArgumentNullException>(() => builder.SetCoreProperties(null!));
    }

    [Fact]
    public void GetCoreProperties_EmptyWorkbook_ReturnsNulls()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var props = builder.GetCoreProperties();
        Assert.Null(props.Title);
        Assert.Null(props.Creator);
    }

    // ─── Calculation properties ────────────────────────────────────

    [Fact]
    public void SetCalculationProperties_Default_ForcesFullCalcOnLoad()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1").AddFormula("A1", "=1+1", null);
            Assert.Same(builder, builder.SetCalculationProperties(new WorkbookCalculationProperties()));
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var workbook = doc.WorkbookPart!.Workbook!;
        var calcPr = workbook.CalculationProperties;
        Assert.NotNull(calcPr);
        Assert.True(calcPr!.FullCalculationOnLoad?.Value);

        // calcPr must follow sheets in the CT_Workbook sequence
        var childNames = workbook.ChildElements.Select(c => c.LocalName).ToList();
        Assert.True(childNames.IndexOf("sheets") < childNames.IndexOf("calcPr"), $"got {string.Join(",", childNames)}");
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void SetCalculationProperties_CustomValues_RoundTrip()
    {
        byte[] bytes;
        using (var builder = WorkbookBuilder.Create())
        {
            builder.SetCalculationProperties(new WorkbookCalculationProperties
            {
                FullCalcOnLoad = false,
                ForceFullCalc = true,
                CalcOnSave = true,
                CalculationId = 191029
            });
            bytes = builder.SaveToBytes();
        }

        using (var reader = WorkbookBuilder.Open(bytes))
        {
            var props = reader.GetCalculationProperties();
            Assert.False(props.FullCalcOnLoad);
            Assert.True(props.ForceFullCalc);
            Assert.True(props.CalcOnSave);
            Assert.Equal(191029U, props.CalculationId);
        }
    }

    [Fact]
    public void SetCalculationProperties_RepeatedCall_UpdatesInPlace()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1");
            builder.SetCalculationProperties(new WorkbookCalculationProperties { CalculationId = 1 });
            builder.SetCalculationProperties(new WorkbookCalculationProperties { CalculationId = 2 });
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var workbook = doc.WorkbookPart!.Workbook!;
        Assert.Single(workbook.ChildElements, c => c is CalculationProperties);
        Assert.Equal(2U, workbook.CalculationProperties!.CalculationId?.Value);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void GetCalculationProperties_NoneSet_ReturnsFalseFullCalcOnLoad()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var props = builder.GetCalculationProperties();
        Assert.False(props.FullCalcOnLoad);
        Assert.Null(props.CalculationId);
    }

    [Fact]
    public void FormulaPlusFullCalcOnLoad_RecalculationContract()
    {
        // The full product contract: a formula cell with no cached value plus full calc on
        // load means Excel computes the value on open instead of displaying nothing.
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var sheet = builder.AddWorksheet("Sheet1");
            sheet.AddCellNumber("A1", 10);
            sheet.AddCellNumber("B1", 20);
            sheet.AddFormula("C1", "=SUM(A1:B1)", null);
            builder.SetCalculationProperties(new WorkbookCalculationProperties());
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var cell = GetCell(doc, "C1");
        Assert.NotNull(cell.CellFormula);
        Assert.Null(cell.CellValue);
        Assert.True(doc.WorkbookPart!.Workbook!.CalculationProperties!.FullCalculationOnLoad?.Value);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    private static Cell GetCell(SpreadsheetDocument doc, string reference)
    {
        return doc.WorkbookPart!.WorksheetParts.First().Worksheet!.GetFirstChild<SheetData>()!
            .Elements<Row>()
            .SelectMany(r => r.Elements<Cell>())
            .Single(c => c.CellReference!.Value == reference);
    }

    public void Dispose()
    {
        if (File.Exists(_testFilePath))
        {
            File.Delete(_testFilePath);
        }
    }
}
