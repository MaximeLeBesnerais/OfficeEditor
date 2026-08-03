using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxEditor.Core.Builders;
using XlsxEditor.Core.Exceptions;
using XlsxEditor.Core.Rendering.Models;
using XlsxEditor.Core.Rendering.Read;

namespace DocxEditor.Tests.Unit;

public class XlsxRenderReaderTests : IDisposable
{
    private readonly string _testFilePath = Path.Combine(
        Path.GetTempPath(),
        $"test_xlsx_render_{Guid.NewGuid()}.xlsx");

    public void Dispose()
    {
        if (File.Exists(_testFilePath))
        {
            try { File.Delete(_testFilePath); } catch { }
        }
    }

    private XlsxReadResult BuildAndRead(Action<IWorksheetBuilder> build, string sheetName = "Sheet1")
    {
        var reader = new XlsxReader();
        using (var builder = WorkbookBuilder.Create())
        {
            var ws = builder.AddWorksheet(sheetName);
            build(ws);
            return reader.Read(builder.SaveToBytes());
        }
    }

    [Fact]
    public void Read_SharedStrings()
    {
        var result = BuildAndRead(ws =>
        {
            ws.AddCellString("A1", "Hello");
            ws.AddCellString("B1", "World");
            ws.AddCellString("C1", "");
        });

        var sheet = result.Workbook.Sheets[0];
        Assert.Equal("Sheet1", sheet.Name);
        Assert.Single(sheet.Rows);

        var cells = sheet.Rows[0].Cells;
        Assert.Equal(3, cells.Count);
        Assert.Equal(XlsxValueKind.Text, cells[0].Value.Kind);
        Assert.Equal("Hello", cells[0].Value.TextValue);
        Assert.Equal(XlsxValueKind.Text, cells[1].Value.Kind);
        Assert.Equal("World", cells[1].Value.TextValue);
        Assert.Equal(XlsxValueKind.Empty, cells[2].Value.Kind);
    }

    [Fact]
    public void Read_Numbers()
    {
        var result = BuildAndRead(ws =>
        {
            ws.AddCellNumber("A1", 42.5);
            ws.AddCellNumber("B1", 0);
            ws.AddCellNumber("C1", -3.14);
        });

        var cells = result.Workbook.Sheets[0].Rows[0].Cells;
        Assert.Equal(XlsxValueKind.Number, cells[0].Value.Kind);
        Assert.Equal(42.5, cells[0].Value.NumberValue);
        Assert.Equal(XlsxValueKind.Number, cells[2].Value.Kind);
        Assert.Equal(-3.14, cells[2].Value.NumberValue);
    }

    [Fact]
    public void Read_Booleans()
    {
        var result = BuildAndRead(ws =>
        {
            ws.AddCellBoolean("A1", true);
            ws.AddCellBoolean("A2", false);
        });

        var cells = result.Workbook.Sheets[0].Rows[0].Cells;
        Assert.Equal(XlsxValueKind.Boolean, cells[0].Value.Kind);
        Assert.True(cells[0].Value.BooleanValue);
        var cells2 = result.Workbook.Sheets[0].Rows[1].Cells;
        Assert.Equal(XlsxValueKind.Boolean, cells2[0].Value.Kind);
        Assert.False(cells2[0].Value.BooleanValue);
    }

    [Fact]
    public void Read_Date()
    {
        var result = BuildAndRead(ws =>
        {
            ws.AddCellDate("A1", "2024-01-15");
        });

        var value = result.Workbook.Sheets[0].Rows[0].Cells[0].Value;
        Assert.Equal(XlsxValueKind.DateTime, value.Kind);
        Assert.Equal(new DateTime(2024, 1, 15), value.DateTimeValue?.Date);
        Assert.NotNull(value.NumberValue);
    }

    [Fact]
    public void Read_DateTime()
    {
        var result = BuildAndRead(ws =>
        {
            ws.AddCellDateTime("A1", "2024-01-15T14:30:00");
        });

        var value = result.Workbook.Sheets[0].Rows[0].Cells[0].Value;
        Assert.Equal(XlsxValueKind.DateTime, value.Kind);
        Assert.Equal(2024, value.DateTimeValue?.Year);
        Assert.Equal(1, value.DateTimeValue?.Month);
        Assert.Equal(15, value.DateTimeValue?.Day);
        Assert.Equal(14, value.DateTimeValue?.Hour);
        Assert.InRange(value.DateTimeValue?.Minute ?? 0, 29, 30);
    }

    [Fact]
    public void Read_Formulas()
    {
        var result = BuildAndRead(ws =>
        {
            ws.AddCellNumber("A1", 10);
            ws.AddCellNumber("A2", 20);
            ws.AddFormula("A3", "SUM(A1:A2)");
        });

        var cells = result.Workbook.Sheets[0].Rows;
        Assert.Equal(3, cells.Count);

        var formulaCell = cells[2].Cells[0];
        Assert.Equal("=SUM(A1:A2)", formulaCell.Formula);
        Assert.Equal(XlsxValueKind.Empty, formulaCell.Value.Kind);
    }

    [Fact]
    public void Read_ColumnWidths()
    {
        var result = BuildAndRead(ws =>
        {
            ws.AddCellString("A1", "Hello");
            ws.SetColumnWidth("A", 12.5);
            ws.SetColumnWidth("B", 20);
            ws.SetColumnWidth("C", 30);
        });

        var columns = result.Workbook.Sheets[0].Columns;
        Assert.Equal(3, columns.Count);
        Assert.Equal(1, columns[0].Index);
        Assert.Equal(87.5, columns[0].WidthPt);
        Assert.Equal(2, columns[1].Index);
        Assert.Equal(140, columns[1].WidthPt);
    }

    [Fact]
    public void Read_RowHeights()
    {
        var result = BuildAndRead(ws =>
        {
            ws.AddCellString("A1", "Row 1");
            ws.AddCellString("A2", "Row 2");
            ws.SetRowHeight(1, 25);
            ws.SetRowHeight(2, 30);
        });

        var rows = result.Workbook.Sheets[0].Rows;
        Assert.Equal(2, rows.Count);
        Assert.Equal(1, rows[0].Index);
        Assert.Equal(25, rows[0].HeightPt);
        Assert.Equal(2, rows[1].Index);
        Assert.Equal(30, rows[1].HeightPt);
    }

    [Fact]
    public void Read_Merges()
    {
        var result = BuildAndRead(ws =>
        {
            ws.AddCellString("A1", "Title");
            ws.MergeCells("A1:C1");
            ws.AddCellString("A3", "Data");
            ws.MergeCells("A3:B4");
        });

        var merges = result.Workbook.Sheets[0].Merges;
        Assert.Equal(2, merges.Count);
        Assert.Equal(1, merges[0].FirstRow);
        Assert.Equal(1, merges[0].LastRow);
        Assert.Equal(1, merges[0].FirstCol);
        Assert.Equal(3, merges[0].LastCol);
    }

    [Fact]
    public void Read_FreezePanes()
    {
        var result = BuildAndRead(ws =>
        {
            ws.AddCellString("A1", "Frozen");
            ws.FreezePanes(1, 1);
        });

        var freeze = result.Workbook.Sheets[0].FreezePanes;
        Assert.NotNull(freeze);
        Assert.Equal(1, freeze.FrozenRows);
        Assert.Equal(1, freeze.FrozenCols);
    }

    [Fact]
    public void Read_NoFreezePanes()
    {
        var result = BuildAndRead(ws =>
        {
            ws.AddCellString("A1", "Data");
        });

        Assert.Null(result.Workbook.Sheets[0].FreezePanes);
    }

    [Fact]
    public void Read_Tables()
    {
        var result = BuildAndRead(ws =>
        {
            ws.AddCellString("A1", "Name");
            ws.AddCellString("B1", "Age");
            ws.AddCellString("C1", "City");
            ws.AddCellString("A2", "Alice");
            ws.AddCellNumber("B2", 30);
            ws.AddCellString("C2", "Paris");
            ws.AddTable("A1", "C2", "People");
        });

        var tables = result.Workbook.Sheets[0].Tables;
        Assert.Single(tables);
        Assert.Equal("People", tables[0].Name);
        Assert.Equal(1, tables[0].FirstRow);
        Assert.Equal(2, tables[0].LastRow);
        Assert.Equal(1, tables[0].FirstCol);
        Assert.Equal(3, tables[0].LastCol);
    }

    [Fact]
    public void Read_AutoFilter()
    {
        var result = BuildAndRead(ws =>
        {
            ws.AddCellString("A1", "Header");
            ws.AddCellString("A2", "Data1");
            ws.AddCellString("A3", "Data2");
            ws.SetAutoFilter("A1:A3");
        });

        Assert.Equal("A1:A3", result.Workbook.Sheets[0].AutoFilterRange);
        Assert.NotEmpty(result.Workbook.Sheets[0].Rows);
    }

    [Fact]
    public void Read_NamedStyles_BordersAndFills()
    {
        var reader = new XlsxReader();
        XlsxReadResult result;
        using (var builder = WorkbookBuilder.Create())
        {
            var ws = builder.AddWorksheet("Sheet1");
            builder.DefineStyle(new CellStyleSpec
            {
                Name = "Header",
                Fill = new CellFillSpec { SolidColorArgb = "FF0000" },
                Border = new CellBorderSpec
                {
                    Bottom = new CellEdgeSpec { Style = BorderStyleValues.Thin, ColorArgb = "000000" }
                },
                Font = new CellFontSpec { Bold = true }
            });
            ws.AddCellString("A1", "Title", "Header");
            ws.AddCellString("B1", "Normal");
            result = reader.Read(builder.SaveToBytes());
        }

        var cells = result.Workbook.Sheets[0].Rows[0].Cells;
        Assert.Equal(2, cells.Count);

        var styledCell = cells[0];
        Assert.NotNull(styledCell.Style);
        Assert.True(styledCell.Style.Bold);
        Assert.Null(styledCell.Style.FontColorArgb);
        Assert.NotNull(styledCell.Style.FillArgb);

        var normalCell = cells[1];
        Assert.Null(normalCell.Style);
    }

    [Fact]
    public void Read_NamedStyles_NumberFormat()
    {
        var reader = new XlsxReader();
        XlsxReadResult result;
        using (var builder = WorkbookBuilder.Create())
        {
            var ws = builder.AddWorksheet("Sheet1");
            builder.DefineStyle(new CellStyleSpec
            {
                Name = "Currency",
                NumberFormat = "$#,##0.00"
            });
            ws.AddCellNumber("A1", 1234.56, "$#,##0.00", "Currency");
            result = reader.Read(builder.SaveToBytes());
        }

        var cell = result.Workbook.Sheets[0].Rows[0].Cells[0];
        Assert.NotNull(cell.NumberFormatCode);
        Assert.Contains("$#", cell.NumberFormatCode);
    }

    [Fact]
    public void Read_FontStyle_BoldItalicSize()
    {
        var reader = new XlsxReader();
        XlsxReadResult result;
        using (var builder = WorkbookBuilder.Create())
        {
            var ws = builder.AddWorksheet("Sheet1");
            builder.DefineStyle(new CellStyleSpec
            {
                Name = "BigTitle",
                Font = new CellFontSpec { Bold = true, Italic = true, Size = 24, ColorArgb = "0000FF" }
            });
            ws.AddCellString("A1", "Big Title", "BigTitle");
            result = reader.Read(builder.SaveToBytes());
        }

        var style = result.Workbook.Sheets[0].Rows[0].Cells[0].Style;
        Assert.NotNull(style);
        Assert.True(style.Bold);
        Assert.True(style.Italic);
        Assert.Equal(24, style.FontSizePt);
        Assert.Equal("FF0000FF", style.FontColorArgb);
    }

    [Fact]
    public void Read_Alignment_WrapText()
    {
        var reader = new XlsxReader();
        XlsxReadResult result;
        using (var builder = WorkbookBuilder.Create())
        {
            var ws = builder.AddWorksheet("Sheet1");
            builder.DefineStyle(new CellStyleSpec
            {
                Name = "WrappedCenter",
                Alignment = new CellAlignmentSpec
                {
                    Horizontal = HorizontalAlignmentValues.Center,
                    Vertical = VerticalAlignmentValues.Center,
                    WrapText = true
                }
            });
            ws.AddCellString("A1", "Centered", "WrappedCenter");
            result = reader.Read(builder.SaveToBytes());
        }

        var style = result.Workbook.Sheets[0].Rows[0].Cells[0].Style;
        Assert.NotNull(style);
        Assert.Equal("center", style.HorizontalAlignment);
        Assert.Equal("center", style.VerticalAlignment);
        Assert.True(style.WrapText);
    }

    [Fact]
    public void Read_AllBorderEdges()
    {
        var reader = new XlsxReader();
        XlsxReadResult result;
        using (var builder = WorkbookBuilder.Create())
        {
            var ws = builder.AddWorksheet("Sheet1");
            builder.DefineStyle(new CellStyleSpec
            {
                Name = "Boxed",
                Border = new CellBorderSpec
                {
                    Left = new CellEdgeSpec { Style = BorderStyleValues.Thin, ColorArgb = "FF0000" },
                    Right = new CellEdgeSpec { Style = BorderStyleValues.Thin, ColorArgb = "FF0000" },
                    Top = new CellEdgeSpec { Style = BorderStyleValues.Medium, ColorArgb = "0000FF" },
                    Bottom = new CellEdgeSpec { Style = BorderStyleValues.Thick, ColorArgb = "00FF00" }
                }
            });
            ws.AddCellString("A1", "Boxed", "Boxed");
            result = reader.Read(builder.SaveToBytes());
        }

        var style = result.Workbook.Sheets[0].Rows[0].Cells[0].Style;
        Assert.NotNull(style);
        Assert.NotNull(style.BorderLeft);
        Assert.Equal("thin", style.BorderLeft.Style);
        Assert.NotNull(style.BorderTop);
        Assert.Equal("medium", style.BorderTop.Style);
        Assert.NotNull(style.BorderBottom);
        Assert.Equal("thick", style.BorderBottom.Style);
    }

    [Fact]
    public void Read_EmptyCell()
    {
        var result = BuildAndRead(ws =>
        {
            ws.AddCellString("A1", "Hello");
            ws.AddCellString("C1", "World");
        });

        var cells = result.Workbook.Sheets[0].Rows[0].Cells;
        Assert.Equal(2, cells.Count);
        Assert.Equal(1, cells[0].Column);
        Assert.Equal(3, cells[1].Column);
    }

    [Fact]
    public void Read_WorkbookTitle()
    {
        var reader = new XlsxReader();
        XlsxReadResult result;
        using (var builder = WorkbookBuilder.Create())
        {
            builder.SetCoreProperties(new WorkbookCoreProperties { Title = "My Report" });
            builder.AddWorksheet("Sheet1");
            result = reader.Read(builder.SaveToBytes());
        }

        Assert.Equal("My Report", result.Workbook.Title);
    }

    [Fact]
    public void Read_FromRawBytes()
    {
        var result = BuildAndRead(ws =>
        {
            ws.AddCellString("A1", "Raw test");
            ws.AddCellNumber("A2", 99);
        });

        Assert.Single(result.Workbook.Sheets);
        Assert.Equal("Raw test", result.Workbook.Sheets[0].Rows[0].Cells[0].Value.TextValue);
        Assert.Equal(99, result.Workbook.Sheets[0].Rows[1].Cells[0].Value.NumberValue);
    }

    [Fact]
    public void Read_FromStream()
    {
        byte[] bytes;
        using (var builder = WorkbookBuilder.Create())
        {
            builder.AddWorksheet("Sheet1").AddCellString("A1", "Stream test");
            bytes = builder.SaveToBytes();
        }

        using var stream = new MemoryStream(bytes);

        var reader = new XlsxReader();
        var result = reader.Read(stream);

        Assert.Single(result.Workbook.Sheets);
        Assert.Equal("Stream test", result.Workbook.Sheets[0].Rows[0].Cells[0].Value.TextValue);
    }

    [Fact]
    public void Read_MultipleSheets()
    {
        var reader = new XlsxReader();
        XlsxReadResult result;
        using (var builder = WorkbookBuilder.Create())
        {
            builder.AddWorksheet("Sales").AddCellString("A1", "Revenue");
            builder.AddWorksheet("Costs").AddCellString("A1", "Expenses");
            builder.GetWorksheet("Costs").AddCellNumber("A2", 500);
            result = reader.Read(builder.SaveToBytes());
        }

        Assert.Equal(2, result.Workbook.Sheets.Count);
        Assert.Equal("Sales", result.Workbook.Sheets[0].Name);
        Assert.Equal("Revenue", result.Workbook.Sheets[0].Rows[0].Cells[0].Value.TextValue);
        Assert.Equal("Costs", result.Workbook.Sheets[1].Name);
        Assert.Equal(500, result.Workbook.Sheets[1].Rows[1].Cells[0].Value.NumberValue);
    }

    [Fact]
    public void Read_LeapYearBug_Serial60()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1").AddCellString("A1", "dummy");
            builder.Save();
        }

        using (var doc = SpreadsheetDocument.Open(_testFilePath, true))
        {
            var sheetData = doc.WorkbookPart!.WorksheetParts.First().Worksheet!
                .GetFirstChild<SheetData>()!;

            var cell = new Cell { CellReference = "B1" };
            cell.CellValue = new CellValue(60.0);
            cell.DataType = CellValues.Number;
            cell.StyleIndex = 1U;

            var row = new Row { RowIndex = 1U };
            row.Append(cell);
            sheetData.Append(row);

            var stylesPart = doc.WorkbookPart.WorkbookStylesPart
                ?? doc.WorkbookPart.AddNewPart<WorkbookStylesPart>();

            stylesPart.Stylesheet ??= new Stylesheet();
            stylesPart.Stylesheet.NumberingFormats ??= new NumberingFormats();
            stylesPart.Stylesheet.NumberingFormats.Append(new NumberingFormat
            {
                NumberFormatId = 164U,
                FormatCode = "m/d/yyyy"
            });

            if (stylesPart.Stylesheet.CellFormats is null)
            {
                stylesPart.Stylesheet.CellFormats = new CellFormats(
                    new CellFormat(),
                    new CellFormat { NumberFormatId = 164U, ApplyNumberFormat = true });
            }

            doc.Save();
        }

        var reader = new XlsxReader();
        var result = reader.Read(_testFilePath);

        Assert.NotEmpty(result.Workbook.Sheets);
        Assert.NotEmpty(result.Workbook.Sheets[0].Rows);
    }

    [Fact]
    public void Read_1904DateSystem()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1");
            builder.Save();
        }

        using (var doc = SpreadsheetDocument.Open(_testFilePath, true))
        {
            doc.WorkbookPart!.Workbook!.WorkbookProperties = new WorkbookProperties { Date1904 = true };

            var sheetData = doc.WorkbookPart.WorksheetParts.First().Worksheet!
                .GetFirstChild<SheetData>()!;

            var cell = new Cell { CellReference = "A1" };
            cell.CellValue = new CellValue(1.0);
            cell.DataType = CellValues.Number;

            var row = new Row { RowIndex = 1U };
            row.Append(cell);
            sheetData.Append(row);

            var stylesPart = doc.WorkbookPart.WorkbookStylesPart
                ?? doc.WorkbookPart.AddNewPart<WorkbookStylesPart>();

            stylesPart.Stylesheet ??= new Stylesheet();
            stylesPart.Stylesheet.NumberingFormats ??= new NumberingFormats();
            stylesPart.Stylesheet.NumberingFormats.Append(new NumberingFormat
            {
                NumberFormatId = 164U,
                FormatCode = "m/d/yyyy"
            });

            if (stylesPart.Stylesheet.CellFormats is null)
            {
                stylesPart.Stylesheet.CellFormats = new CellFormats(
                    new CellFormat(),
                    new CellFormat { NumberFormatId = 164U, ApplyNumberFormat = true });
            }

            doc.Save();
        }

        var reader = new XlsxReader();
        var result = reader.Read(_testFilePath);

        Assert.NotEmpty(result.Workbook.Sheets);
        Assert.NotEmpty(result.Workbook.Sheets[0].Rows);
    }

    [Fact]
    public void Read_ErrorCellValue()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1");
            builder.Save();
        }

        using (var doc = SpreadsheetDocument.Open(_testFilePath, true))
        {
            var sheetData = doc.WorkbookPart!.WorksheetParts.First().Worksheet!
                .GetFirstChild<SheetData>()!;

            var cell = new Cell { CellReference = "A1" };
            cell.CellValue = new CellValue("#DIV/0!");
            cell.DataType = CellValues.Error;

            var row = new Row { RowIndex = 1U };
            row.Append(cell);
            sheetData.Append(row);

            var stylesPart = doc.WorkbookPart.WorkbookStylesPart
                ?? doc.WorkbookPart.AddNewPart<WorkbookStylesPart>();

            stylesPart.Stylesheet ??= new Stylesheet();
            stylesPart.Stylesheet.CellFormats ??= new CellFormats(new CellFormat());

            doc.Save();
        }

        var reader = new XlsxReader();
        var result = reader.Read(_testFilePath);

        var value = result.Workbook.Sheets[0].Rows[0].Cells[0].Value;
        Assert.Equal(XlsxValueKind.Error, value.Kind);
        Assert.Equal("#DIV/0!", value.ErrorValue);
    }

    [Fact]
    public void Read_ColumnWidth_Default()
    {
        var result = BuildAndRead(ws =>
        {
            ws.AddCellString("A1", "Data");
        });

        Assert.Empty(result.Workbook.Sheets[0].Columns);
    }

    [Fact]
    public void Read_MergeRange_WithSingleCell()
    {
        var result = BuildAndRead(ws =>
        {
            ws.AddCellString("A1", "Value");
        });

        Assert.Empty(result.Workbook.Sheets[0].Merges);
    }

    [Fact]
    public void Read_PageSetup_Portrait()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1");
            builder.Save();
        }

        using (var doc = SpreadsheetDocument.Open(_testFilePath, true))
        {
            var worksheet = doc.WorkbookPart!.WorksheetParts.First().Worksheet!;

            worksheet.Append(new PageSetup
            {
                Orientation = OrientationValues.Portrait,
                PaperSize = 9U
            });
            doc.Save();
        }

        var reader = new XlsxReader();
        var result = reader.Read(_testFilePath);

        var sheet = result.Workbook.Sheets[0];
        Assert.False(sheet.Landscape);
        Assert.True(sheet.PageWidthPt > 0);
        Assert.True(sheet.PageHeightPt > 0);
    }

    [Fact]
    public void Read_PageSetup_Landscape()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1");
            builder.Save();
        }

        using (var doc = SpreadsheetDocument.Open(_testFilePath, true))
        {
            var worksheet = doc.WorkbookPart!.WorksheetParts.First().Worksheet!;

            worksheet.Append(new PageSetup
            {
                Orientation = OrientationValues.Landscape,
                PaperSize = 9U
            });
            doc.Save();
        }

        var reader = new XlsxReader();
        var result = reader.Read(_testFilePath);

        Assert.True(result.Workbook.Sheets[0].Landscape);
    }

    [Fact]
    public void Read_ParseCellReference()
    {
        Assert.Equal((1, 1), XlsxReader.ParseCellReference("A1"));
        Assert.Equal((10, 26), XlsxReader.ParseCellReference("Z10"));
        Assert.Equal((100, 27), XlsxReader.ParseCellReference("AA100"));
        Assert.Equal((999, 702), XlsxReader.ParseCellReference("ZZ999"));
        Assert.Equal((0, 0), XlsxReader.ParseCellReference(""));
        Assert.Equal((0, 0), XlsxReader.ParseCellReference(null));
    }

    [Fact]
    public void Read_IsDateFormatCode()
    {
        Assert.True(XlsxReader.IsDateFormatCode("14"));
        Assert.True(XlsxReader.IsDateFormatCode("22"));
        Assert.True(XlsxReader.IsDateFormatCode("yyyy-mm-dd"));
        Assert.True(XlsxReader.IsDateFormatCode("m/d/yyyy h:mm"));
        Assert.False(XlsxReader.IsDateFormatCode("0"));
        Assert.False(XlsxReader.IsDateFormatCode("0.00"));
        Assert.False(XlsxReader.IsDateFormatCode(null));
    }

    [Fact]
    public void Read_SerialToDateTime_1900()
    {
        var dt2 = XlsxReader.SerialToDateTime(2.0, false);
        Assert.Equal(new DateTime(1900, 1, 1), dt2.Date);

        var dt61 = XlsxReader.SerialToDateTime(61.0, false);
        Assert.Equal(new DateTime(1900, 3, 1), dt61.Date);
    }

    [Fact]
    public void Read_SerialToDateTime_1904()
    {
        var dt = XlsxReader.SerialToDateTime(0.0, true);
        Assert.Equal(new DateTime(1904, 1, 1), dt.Date);

        dt = XlsxReader.SerialToDateTime(1.0, true);
        Assert.Equal(new DateTime(1904, 1, 2), dt.Date);
    }

    [Fact]
    public void Read_EmptyWorkbook()
    {
        byte[] bytes;
        using (var builder = WorkbookBuilder.Create())
        {
            bytes = builder.SaveToBytes();
        }

        var reader = new XlsxReader();
        var result = reader.Read(bytes);

        Assert.Empty(result.Workbook.Sheets);
        Assert.Empty(result.Issues);
    }

    [Fact]
    public void Read_IssuesOnMissingStylePart()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1").AddCellString("A1", "Hello");
            builder.Save();
        }

        using (var doc = SpreadsheetDocument.Open(_testFilePath, true))
        {
            var stylesPart = doc.WorkbookPart!.WorkbookStylesPart;
            if (stylesPart is not null)
            {
                doc.WorkbookPart.DeletePart(stylesPart);
            }
            doc.Save();
        }

        var reader = new XlsxReader();
        var result = reader.Read(_testFilePath);

        Assert.NotEmpty(result.Workbook.Sheets);
        Assert.Equal("Hello", result.Workbook.Sheets[0].Rows[0].Cells[0].Value.TextValue);
    }
}
