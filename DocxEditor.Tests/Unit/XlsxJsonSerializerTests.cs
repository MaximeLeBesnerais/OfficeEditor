using System.Text.Json;
using XlsxEditor.Core.Builders;
using XlsxEditor.Core.Rendering;
using XlsxEditor.Core.Rendering.Models;
using Xunit;

namespace DocxEditor.Tests.Unit;

public class XlsxJsonSerializerTests
{
    [Fact]
    public void Serialize_ProducesInstructionJson()
    {
        var workbook = BuildWorkbook();
        var json = XlsxJsonSerializer.Serialize(workbook);

        var jsonText = JsonSerializer.Serialize(json);
        using var doc = JsonDocument.Parse(jsonText);
        var root = doc.RootElement;

        Assert.Equal("1.0", root.GetProperty("version").GetString());
        var sheets = root.GetProperty("worksheets");
        Assert.Equal(1, sheets.GetArrayLength());

        var sheet = sheets[0];
        Assert.Equal("Data", sheet.GetProperty("name").GetString());
        Assert.Equal("Region", sheet.GetProperty("headers")[0].GetString());
        Assert.Equal("1250000", sheet.GetProperty("rows")[0][1].GetString());
        Assert.True(sheet.GetProperty("columns").GetArrayLength() >= 1);
    }

    [Fact]
    public void Serialize_MergesAndFreezePanes_ArePreserved()
    {
        var workbook = BuildWorkbook();
        var json = XlsxJsonSerializer.Serialize(workbook);
        var jsonText = JsonSerializer.Serialize(json);
        using var doc = JsonDocument.Parse(jsonText);
        var sheet = doc.RootElement.GetProperty("worksheets")[0];

        Assert.True(sheet.TryGetProperty("freezePanes", out var freeze));
        Assert.Equal(2, freeze.GetProperty("row").GetInt32());
    }

    [Fact]
    public void Serialize_RoundTripsThroughGenerator()
    {
        var workbook = BuildWorkbook();
        var json = JsonSerializer.Serialize(XlsxJsonSerializer.Serialize(workbook));

        var generated = XlsxEditor.Core.Instructions.XlsxGenerator.Generate(json);
        Assert.True(generated.IsValid, string.Join("; ", generated.Validation.Errors.Select(e => e.Message)));
        Assert.NotNull(generated.Bytes);
        Assert.NotEmpty(generated.Bytes!);
    }

    private static XlsxRenderWorkbook BuildWorkbook()
    {
        var sheet = new XlsxRenderSheet
        {
            Name = "Data",
            Columns =
            [
                new XlsxRenderColumn { Index = 1, WidthPt = 90 },
                new XlsxRenderColumn { Index = 2, WidthPt = 110 }
            ],
            Rows =
            [
                new XlsxRenderRow
                {
                    Index = 1,
                    Cells =
                    [
                        new XlsxRenderCell { Row = 1, Column = 1, Value = XlsxRenderValue.Text("Region") },
                        new XlsxRenderCell { Row = 1, Column = 2, Value = XlsxRenderValue.Text("Revenue") }
                    ]
                },
                new XlsxRenderRow
                {
                    Index = 2,
                    Cells =
                    [
                        new XlsxRenderCell { Row = 2, Column = 1, Value = XlsxRenderValue.Text("North America") },
                        new XlsxRenderCell { Row = 2, Column = 2, Value = XlsxRenderValue.Number(1250000), NumberFormatCode = "$#,##0.00" }
                    ]
                },
                new XlsxRenderRow
                {
                    Index = 3,
                    Cells = [new XlsxRenderCell { Row = 3, Column = 1, Value = XlsxRenderValue.Text("Europe") }]
                }
            ],
            FreezePanes = new XlsxRenderFreezePanes { FrozenRows = 2 }
        };

        return new XlsxRenderWorkbook { Title = "Report", Sheets = [sheet] };
    }
}
