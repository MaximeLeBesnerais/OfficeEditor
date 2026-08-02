using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxEditor.Core.Builders;
using XlsxEditor.Core.Exceptions;

namespace DocxEditor.Tests.Unit;

public class WorkbookBuilderTypedCellTests : IDisposable
{
    private readonly string _testFilePath = Path.Combine(Path.GetTempPath(), $"test_xlsx_typed_{Guid.NewGuid()}.xlsx");

    // ─── Typed string ──────────────────────────────────────────────

    [Fact]
    public void AddCellString_WritesSharedString_AndClearsStaleFormula()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var sheet = builder.AddWorksheet("Sheet1");
            sheet.AddFormula("A1", "=1+1", null);
            sheet.AddCellString("A1", "typed");
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var cell = GetCell(doc, "A1");
        Assert.Equal(CellValues.SharedString, cell.DataType?.Value);
        Assert.Null(cell.CellFormula);
        Assert.NotNull(cell.CellValue);
        var sharedStrings = doc.WorkbookPart!.SharedStringTablePart!.SharedStringTable!
            .Elements<SharedStringItem>().ToList();
        Assert.Equal("typed", sharedStrings[int.Parse(cell.CellValue!.Text)].InnerText);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void AddCellString_NullValue_Throws()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var sheet = builder.AddWorksheet("Sheet1");
        var ex = Assert.Throws<ArgumentNullException>(() => sheet.AddCellString("A1", null!));
        Assert.Equal("value", ex.ParamName);
        Assert.False(sheet.CellExists("A1"));
    }

    [Fact]
    public void AddCellString_WithNamedStyle_AppliesStyleIndex()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.DefineStyle(new CellStyleSpec { Name = "Mono", Font = new CellFontSpec { Italic = true } });
            builder.AddWorksheet("Sheet1").AddCellString("B2", "styled", "Mono");
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var cell = GetCell(doc, "B2");
        Assert.Equal(1U, cell.StyleIndex?.Value);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void AddCellString_WithoutStyle_LeavesStyleUnset()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1").AddCellString("A1", "plain");
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        Assert.Null(GetCell(doc, "A1").StyleIndex);
    }

    // ─── Typed number ──────────────────────────────────────────────

    [Fact]
    public void AddCellNumber_WritesNumberCell_WithExactValue()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1").AddCellNumber("A1", 42.5);
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var cell = GetCell(doc, "A1");
        Assert.Equal(CellValues.Number, cell.DataType?.Value);
        Assert.Equal("42.5", cell.CellValue?.Text);
        Assert.Null(cell.CellFormula);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    public void AddCellNumber_RejectsNonFinite(double value)
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var sheet = builder.AddWorksheet("Sheet1");
        var ex = Assert.Throws<XlsxException>(() => sheet.AddCellNumber("A1", value));
        Assert.Contains("finite", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.False(sheet.CellExists("A1"));
    }

    [Fact]
    public void AddCellNumber_WithNumberFormat_AppliesCustomNumFmt()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1").AddCellNumber("A1", 1234.5, "0.000");
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var cell = GetCell(doc, "A1");
        Assert.NotNull(cell.StyleIndex);
        var cellFormat = doc.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.CellFormats!
            .Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
        Assert.Equal(164U, cellFormat.NumberFormatId?.Value);
        Assert.True(cellFormat.ApplyNumberFormat?.Value);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void AddCellNumber_OverFormulaCell_ClearsFormulaAndDataType()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var sheet = builder.AddWorksheet("Sheet1");
            sheet.AddFormula("A1", "=SUM(B1:B2)", null);
            sheet.AddCellNumber("A1", 99);
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var cell = GetCell(doc, "A1");
        Assert.Null(cell.CellFormula);
        Assert.Equal(CellValues.Number, cell.DataType?.Value);
        Assert.Equal("99", cell.CellValue?.Text);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    // ─── Typed boolean ─────────────────────────────────────────────

    [Theory]
    [InlineData(true, "1")]
    [InlineData(false, "0")]
    public void AddCellBoolean_WritesBooleanCell(bool value, string expected)
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1").AddCellBoolean("A1", value);
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var cell = GetCell(doc, "A1");
        Assert.Equal(CellValues.Boolean, cell.DataType?.Value);
        Assert.Equal(expected, cell.CellValue?.Text);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void AddCellBoolean_RoundTrip_ReturnsRawValue()
    {
        byte[] bytes;
        using (var builder = WorkbookBuilder.Create())
        {
            builder.AddWorksheet("Sheet1").AddCellBoolean("A1", true);
            bytes = builder.SaveToBytes();
        }

        using var reader = WorkbookBuilder.Open(bytes);
        Assert.Equal("1", reader.GetWorksheet("Sheet1").GetCellValue("A1"));
    }

    // ─── ISO date ──────────────────────────────────────────────────

    [Fact]
    public void AddCellDate_WritesSerialWithDefaultDateFormat()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1").AddCellDate("A1", "2024-01-15");
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var cell = GetCell(doc, "A1");
        Assert.Equal(CellValues.Number, cell.DataType?.Value);
        // 2024-01-15 is day 45306 in Excel's 1900 date system (epoch 1899-12-30)
        Assert.Equal("45306", cell.CellValue?.Text);
        Assert.NotNull(cell.StyleIndex);

        var cellFormat = doc.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.CellFormats!
            .Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
        Assert.Equal(164U, cellFormat.NumberFormatId?.Value);
        var numFmt = doc.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.NumberingFormats!
            .Elements<NumberingFormat>().Single();
        Assert.Equal("yyyy-mm-dd", numFmt.FormatCode?.Value);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void AddCellDate_CustomNumberFormat_OverridesDefault()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1").AddCellDate("A1", "2024-01-15", "dd/mm/yyyy");
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var numFmt = doc.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.NumberingFormats!
            .Elements<NumberingFormat>().Single();
        Assert.Equal("dd/mm/yyyy", numFmt.FormatCode?.Value);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Theory]
    [InlineData("2024/01/15")]
    [InlineData("15-01-2024")]
    [InlineData("2024-01-15T10:00:00")]
    [InlineData("not-a-date")]
    [InlineData("")]
    public void AddCellDate_InvalidIso_Throws(string isoDate)
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var sheet = builder.AddWorksheet("Sheet1");
        var ex = Assert.Throws<XlsxException>(() => sheet.AddCellDate("A1", isoDate));
        Assert.Contains("ISO date", ex.Message);
        Assert.False(sheet.CellExists("A1"));
    }

    [Fact]
    public void AddCellDate_BeforeExcelEpoch_Throws()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var sheet = builder.AddWorksheet("Sheet1");
        var ex = Assert.Throws<XlsxException>(() => sheet.AddCellDate("A1", "1899-12-31"));
        Assert.Contains("1900", ex.Message);
        Assert.False(sheet.CellExists("A1"));
    }

    [Fact]
    public void AddCellDate_RoundTrip_SerialSurvivesReopen()
    {
        byte[] bytes;
        using (var builder = WorkbookBuilder.Create())
        {
            builder.AddWorksheet("Sheet1").AddCellDate("A1", "2024-01-15");
            bytes = builder.SaveToBytes();
        }

        using var reader = WorkbookBuilder.Open(bytes);
        Assert.Equal("45306", reader.GetWorksheet("Sheet1").GetCellValue("A1"));
    }

    // ─── ISO datetime ──────────────────────────────────────────────

    [Fact]
    public void AddCellDateTime_WritesSerialWithFraction_AndDatetimeFormat()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1").AddCellDateTime("A1", "2024-01-15T14:30:00");
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var cell = GetCell(doc, "A1");
        Assert.Equal(CellValues.Number, cell.DataType?.Value);
        var serial = double.Parse(cell.CellValue!.Text, CultureInfo.InvariantCulture);
        // 45306 full days + 14.5/24 ≈ 0.6041666667
        Assert.Equal(45306.6041666667, serial, 9);

        var numFmt = doc.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.NumberingFormats!
            .Elements<NumberingFormat>().Single();
        Assert.Equal("yyyy-mm-dd h:mm:ss", numFmt.FormatCode?.Value);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void AddCellDateTime_FractionalSeconds_Accepted()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1").AddCellDateTime("A1", "2024-01-15T14:30:00.500");
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var cell = GetCell(doc, "A1");
        var serial = double.Parse(cell.CellValue!.Text, CultureInfo.InvariantCulture);
        Assert.Equal(45306.6041, serial, 3);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Theory]
    [InlineData("2024-01-15 14:30:00")] // space separator is not ISO
    [InlineData("2024-01-15T14:30")]    // missing seconds
    [InlineData("garbage")]
    public void AddCellDateTime_InvalidIso_Throws(string isoDateTime)
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var sheet = builder.AddWorksheet("Sheet1");
        var ex = Assert.Throws<XlsxException>(() => sheet.AddCellDateTime("A1", isoDateTime));
        Assert.Contains("ISO datetime", ex.Message);
        Assert.False(sheet.CellExists("A1"));
    }

    // ─── Formula cells ─────────────────────────────────────────────

    [Fact]
    public void AddFormula_StripsLeadingEquals_AndClearsStaleValue()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var sheet = builder.AddWorksheet("Sheet1");
            sheet.AddCellNumber("A1", 10);
            sheet.AddFormula("A1", "=SUM(B1:B2)", null);
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var cell = GetCell(doc, "A1");
        Assert.Equal("SUM(B1:B2)", cell.CellFormula?.Text);
        Assert.Null(cell.CellValue);
        Assert.Null(cell.DataType);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void AddFormula_WithStyleAndNumberFormat_AppliesBoth()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.DefineStyle(new CellStyleSpec { Name = "Bold", Font = new CellFontSpec { Bold = true } });
            builder.AddWorksheet("Sheet1").AddFormula("A1", "=1+1", "Bold", "0.00");
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var cell = GetCell(doc, "A1");
        Assert.NotNull(cell.StyleIndex);
        var cellFormat = doc.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.CellFormats!
            .Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
        Assert.Equal(1U, cellFormat.FontId?.Value);       // the Bold style's font
        Assert.Equal(2U, cellFormat.NumberFormatId?.Value); // built-in "0.00"
        Assert.True(cellFormat.ApplyFont?.Value);
        Assert.True(cellFormat.ApplyNumberFormat?.Value);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void AddFormula_WithoutLeadingEquals_Accepted()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1").AddFormula("A1", "A1+B1", null);
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        Assert.Equal("A1+B1", GetCell(doc, "A1").CellFormula?.Text);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Theory]
    [InlineData("=")]
    [InlineData("   ")]
    public void AddFormula_EmptyFormula_Throws(string formula)
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var sheet = builder.AddWorksheet("Sheet1");
        var ex = Assert.Throws<XlsxException>(() => sheet.AddFormula("A1", formula, null));
        Assert.Contains("empty formula", ex.Message);
        Assert.False(sheet.CellExists("A1"));
    }

    [Fact]
    public void AddFormula_RoundTrip_ReturnsDisplaySyntax()
    {
        byte[] bytes;
        using (var builder = WorkbookBuilder.Create())
        {
            builder.AddWorksheet("Sheet1").AddFormula("A1", "=SUM(A2:A3)", null);
            bytes = builder.SaveToBytes();
        }

        using var reader = WorkbookBuilder.Open(bytes);
        Assert.Equal("=SUM(A2:A3)", reader.GetWorksheet("Sheet1").GetCellFormula("A1"));
    }

    [Fact]
    public void AddFormula_UnknownStyle_ThrowsBeforeCellWrite()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var sheet = builder.AddWorksheet("Sheet1");
        var ex = Assert.Throws<XlsxException>(() => sheet.AddFormula("A1", "=1+1", "Nope"));
        Assert.Contains("Nope", ex.Message);
        Assert.False(sheet.CellExists("A1"));
    }

    // ─── Mixed typed writes validate together ──────────────────────

    [Fact]
    public void MixedTypedCells_AllValidateAndRoundTrip()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            var sheet = builder.AddWorksheet("Sheet1");
            sheet.AddCellString("A1", "Text");
            sheet.AddCellNumber("B1", 3.14);
            sheet.AddCellBoolean("C1", true);
            sheet.AddCellDate("D1", "2024-02-29"); // leap day
            sheet.AddCellDateTime("E1", "2024-06-01T12:00:00");
            sheet.AddFormula("F1", "=SUM(B1:B1)", null);
            builder.Save();
        }

        using (var doc = SpreadsheetDocument.Open(_testFilePath, false))
        {
            OpenXmlAssert.NoValidationErrors(doc);
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        var ws = reader.GetWorksheet("Sheet1");
        Assert.Equal("Text", ws.GetCellValue("A1"));
        Assert.Equal("3.14", ws.GetCellValue("B1"));
        Assert.Equal("1", ws.GetCellValue("C1"));
        Assert.Equal("45351", ws.GetCellValue("D1")); // 2024-02-29
        Assert.NotNull(ws.GetCellFormula("F1"));
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
