using System.IO.Compression;
using XlsxEditor.Core.Builders;
using XlsxEditor.Core.Instructions;
using XlsxEditor.Core.Rendering;
using XlsxEditor.Core.Rendering.Emit;
using XlsxEditor.Core.Rendering.Models;
using XlsxEditor.Core.Rendering.Read;
using Xunit;

namespace DocxEditor.Tests.Unit;

public class XlsxRenderEmitterTests
{
    private static XlsxRenderWorkbook BuildSampleWorkbook()
    {
        var sheet = new XlsxRenderSheet
        {
            Name = "Sales",
            PageWidthPt = 842,
            PageHeightPt = 595,
            Landscape = true,
            Columns =
            [
                new XlsxRenderColumn { Index = 1, WidthPt = 90 },
                new XlsxRenderColumn { Index = 2, WidthPt = 120 },
                new XlsxRenderColumn { Index = 3, WidthPt = 80 },
                new XlsxRenderColumn { Index = 4, WidthPt = 90 }
            ],
            Rows =
            [
                new XlsxRenderRow
                {
                    Index = 1,
                    Cells =
                    [
                        new XlsxRenderCell
                        {
                            Row = 1,
                            Column = 1,
                            Value = XlsxRenderValue.Text("Region"),
                            Style = new XlsxRenderStyle { Bold = true, FillArgb = "FF1F3A5F" }
                        },
                        new XlsxRenderCell
                        {
                            Row = 1,
                            Column = 2,
                            Value = XlsxRenderValue.Text("Revenue"),
                            Style = new XlsxRenderStyle { Bold = true, FillArgb = "FF1F3A5F", HorizontalAlignment = "center" }
                        },
                        new XlsxRenderCell
                        {
                            Row = 1,
                            Column = 3,
                            Value = XlsxRenderValue.Text("Growth"),
                            Style = new XlsxRenderStyle { Bold = true, FillArgb = "FF1F3A5F", HorizontalAlignment = "center" }
                        },
                        new XlsxRenderCell
                        {
                            Row = 1,
                            Column = 4,
                            Value = XlsxRenderValue.Text("As Of"),
                            Style = new XlsxRenderStyle { Bold = true, FillArgb = "FF1F3A5F", HorizontalAlignment = "center" }
                        }
                    ]
                },
                new XlsxRenderRow
                {
                    Index = 2,
                    Cells =
                    [
                        new XlsxRenderCell { Row = 2, Column = 1, Value = XlsxRenderValue.Text("North America") },
                        new XlsxRenderCell { Row = 2, Column = 2, Value = XlsxRenderValue.Number(1250000), NumberFormatCode = "$#,##0.00" },
                        new XlsxRenderCell { Row = 2, Column = 3, Value = XlsxRenderValue.Number(0.082), NumberFormatCode = "0.0%" },
                        new XlsxRenderCell { Row = 2, Column = 4, Value = XlsxRenderValue.DateTime(new DateTime(2023, 12, 31), 45291), NumberFormatCode = "yyyy-mm-dd" }
                    ]
                },
                new XlsxRenderRow
                {
                    Index = 3,
                    Cells =
                    [
                        new XlsxRenderCell { Row = 3, Column = 1, Value = XlsxRenderValue.Text("Europe") },
                        new XlsxRenderCell { Row = 3, Column = 2, Value = XlsxRenderValue.Number(-980000), NumberFormatCode = "$#,##0.00;[Red]-$#,##0.00" },
                        new XlsxRenderCell { Row = 3, Column = 3, Value = XlsxRenderValue.Number(0.051), NumberFormatCode = "0.0%" },
                        new XlsxRenderCell { Row = 3, Column = 4, Value = XlsxRenderValue.Text("") }
                    ]
                }
            ],
            FreezePanes = new XlsxRenderFreezePanes { FrozenRows = 1 }
        };

        return new XlsxRenderWorkbook { Title = "Report", Sheets = [sheet] };
    }

    [Fact]
    public void GenerateTypstSource_EmitsTableWithColumnsAndAlignment()
    {
        var source = new XlsxToTypstConverter(BuildSampleWorkbook()).GenerateTypstSource();

        Assert.Contains("#table(", source);
        Assert.Contains("columns: (90pt, 120pt, 80pt, 90pt)", source);
        Assert.Contains("align: (", source);
        Assert.Contains("table.header(repeat: true", source);
    }

    [Fact]
    public void GenerateTypstSource_FormatsNumbersAndDates()
    {
        var source = new XlsxToTypstConverter(BuildSampleWorkbook()).GenerateTypstSource();

        Assert.Contains(@"\$1,250,000.00", source);
        Assert.Contains(@"8.2\%", source);
        Assert.Contains("2023-12-31", source);
    }

    [Fact]
    public void GenerateTypstSource_EmitsNegativeRedColor()
    {
        var source = new XlsxToTypstConverter(BuildSampleWorkbook()).GenerateTypstSource();

        // Negative value with [Red] section gets text(fill: rgb(...)).
        Assert.Contains("text(fill: rgb(\"FF0000\"))", source);
    }

    [Fact]
    public void GenerateTypstSource_EmitsFillAndBoldHeader()
    {
        var source = new XlsxToTypstConverter(BuildSampleWorkbook()).GenerateTypstSource();

        Assert.Contains("fill: rgb(\"1F3A5F\")", source);
    }

    [Fact]
    public void GenerateTypstSource_EscapesSpecialCharacters()
    {
        var sheet = new XlsxRenderSheet
        {
            Name = "A & B #1",
            Columns = [new XlsxRenderColumn { Index = 1, WidthPt = 100 }],
            Rows =
            [
                new XlsxRenderRow
                {
                    Index = 1,
                    Cells = [new XlsxRenderCell { Row = 1, Column = 1, Value = XlsxRenderValue.Text("50% of [$total]") }]
                }
            ]
        };

        var source = new XlsxToTypstConverter(new XlsxRenderWorkbook { Sheets = [sheet] }).GenerateTypstSource();

        Assert.Contains("50\\% of \\[\\$total\\]", source);
    }

    [Fact]
    public void GenerateTypstSource_EmitsMergedCellSpans()
    {
        var sheet = new XlsxRenderSheet
        {
            Name = "Merged",
            Columns =
            [
                new XlsxRenderColumn { Index = 1, WidthPt = 100 },
                new XlsxRenderColumn { Index = 2, WidthPt = 100 },
                new XlsxRenderColumn { Index = 3, WidthPt = 100 }
            ],
            Rows =
            [
                new XlsxRenderRow
                {
                    Index = 1,
                    Cells = [new XlsxRenderCell { Row = 1, Column = 1, Value = XlsxRenderValue.Text("Merged title"), Style = new XlsxRenderStyle { Bold = true } }]
                },
                new XlsxRenderRow
                {
                    Index = 2,
                    Cells =
                    [
                        new XlsxRenderCell { Row = 2, Column = 1, Value = XlsxRenderValue.Text("a") },
                        new XlsxRenderCell { Row = 2, Column = 2, Value = XlsxRenderValue.Text("b") },
                        new XlsxRenderCell { Row = 2, Column = 3, Value = XlsxRenderValue.Text("c") }
                    ]
                }
            ],
            Merges = [new XlsxRenderMerge { FirstRow = 1, LastRow = 1, FirstCol = 1, LastCol = 3 }]
        };

        var source = new XlsxToTypstConverter(new XlsxRenderWorkbook { Sheets = [sheet] }).GenerateTypstSource();

        Assert.Contains("table.cell(colspan: 3)", source);
        Assert.Contains("Merged title", source);
    }

    [Fact]
    public void RenderFile_ProducesPdfFromGeneratedWorkbook()
    {
        // Build a real .xlsx via WorkbookBuilder, then render it.
        var bytes = BuildWorkbookBytes();
        var result = XlsxRenderer.RenderBytes(bytes);

        Assert.True(result.Success, result.Compile.ErrorMessage);
        Assert.NotNull(result.Compile.Pages);
        Assert.Single(result.Compile.Pages);
        Assert.NotEmpty(result.Compile.Pages[0]);
    }

    private static byte[] BuildWorkbookBytes()
    {
        var workbook = WorkbookBuilder.Create();
        var sheet = workbook.AddWorksheet("Data");
        sheet.AddCellString("A1", "Region");
        sheet.AddCellNumber("B1", 1250000, numberFormat: "$#,##0.00");
        sheet.AddCellString("A2", "North America");
        sheet.AddCellNumber("B2", 980000, numberFormat: "$#,##0.00");
        sheet.AddCellString("A3", "Europe");
        return workbook.SaveToBytes();
    }
}
