using System.Text.Json;
using XlsxEditor.Core.Builders;
using XlsxEditor.Core.Instructions;
using XlsxEditor.Core.Exceptions;

namespace DocxEditor.Tests.Unit;

public class XlsxInstructionTests : IDisposable
{
    private readonly string _testFilePath = Path.Combine(Path.GetTempPath(), $"test_xlsx_instructions_{Guid.NewGuid()}.xlsx");

    // ─── Parser ────────────────────────────────────────────────────

    [Fact]
    public void Parse_ShouldProduceInstructionSet_FromValidJson()
    {
        var json = """
        {
            "version": "1.0",
            "worksheets": [
                { "name": "Sheet1", "headers": ["A", "B"], "rows": [["1", "2"]] }
            ]
        }
        """;

        var set = XlsxInstructionParser.Parse(json);
        Assert.Equal("1.0", set.Version);
        Assert.Single(set.Worksheets);
        Assert.Equal("Sheet1", set.Worksheets[0].Name);
    }

    [Fact]
    public void Parse_ShouldThrow_ForNullJson()
    {
        Assert.Throws<XlsxException>(() => XlsxInstructionParser.Parse(null!));
    }

    [Fact]
    public void Parse_ShouldThrow_ForEmptyJson()
    {
        Assert.Throws<XlsxException>(() => XlsxInstructionParser.Parse(""));
    }

    [Fact]
    public void Parse_ShouldThrow_ForNonJsonContent()
    {
        var ex = Assert.Throws<XlsxException>(() => XlsxInstructionParser.Parse("not json"));
        Assert.Contains("Invalid JSON", ex.Message);
    }

    [Fact]
    public void Parse_ShouldThrow_ForMissingVersion()
    {
        var json = """{ "worksheets": [{"name": "S"}] }""";
        Assert.Throws<XlsxException>(() => XlsxInstructionParser.Parse(json));
    }

    [Fact]
    public void Parse_ShouldThrow_ForBadVersion()
    {
        var json = """{ "version": "2.0", "worksheets": [{"name": "S"}] }""";
        var ex = Assert.Throws<XlsxException>(() => XlsxInstructionParser.Parse(json));
        Assert.Contains("2.0", ex.Message);
    }

    [Fact]
    public void Parse_ShouldThrow_ForMissingWorksheets()
    {
        var json = """{ "version": "1.0", "worksheets": [] }""";
        Assert.Throws<XlsxException>(() => XlsxInstructionParser.Parse(json));
    }

    [Fact]
    public void Parse_ShouldDeserializeVariables()
    {
        var json = """
        {
            "version": "1.0",
            "worksheets": [{"name": "S", "rows": [["{{x}}"]]}],
            "variables": {"x": "replaced"}
        }
        """;

        var set = XlsxInstructionParser.Parse(json);
        Assert.NotNull(set.Variables);
        Assert.Equal("replaced", set.Variables!["x"]);
    }

    // ─── Validator ─────────────────────────────────────────────────

    [Fact]
    public void Validate_ShouldRejectDuplicateSheetNames()
    {
        var json = """
        {
            "version": "1.0",
            "worksheets": [
                {"name": "Sales", "rows": [["1"]]},
                {"name": "Sales", "rows": [["2"]]}
            ]
        }
        """;

        var ex = Assert.Throws<XlsxException>(() => XlsxInstructionParser.Parse(json));
        Assert.Contains("Sales", ex.Message);
        Assert.Contains("Duplicate", ex.Message);
    }

    [Fact]
    public void Validate_ShouldRejectDuplicateSheetNames_DifferentCase()
    {
        // Excel sheet names are case-insensitive: "Sales" and "SALES" collide.
        var json = """
        {
            "version": "1.0",
            "worksheets": [
                {"name": "Sales", "rows": [["1"]]},
                {"name": "SALES", "rows": [["2"]]}
            ]
        }
        """;

        var ex = Assert.Throws<XlsxException>(() => XlsxInstructionParser.Parse(json));
        Assert.Contains("Duplicate", ex.Message);
    }

    [Fact]
    public void Execute_ShouldNormalizeLowercaseCellAddresses()
    {
        var set = XlsxInstructionParser.Parse("""
        {
            "version": "1.0",
            "worksheets": [{
                "name": "Calc",
                "cells": [
                    {"address": "a1", "value": "10"},
                    {"address": "A1", "value": "20"}
                ]
            }]
        }
        """);

        // 'a1' and 'A1' are the same cell; last write wins, no duplicate cells.
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            XlsxInstructionExecutor.Execute(set, builder);
            builder.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        var ws = reader.GetWorksheet("Calc");
        Assert.Equal("20", ws.GetCellValue("A1"));
        Assert.Single(ws.GetRow(1)!.Cells);
    }

    [Fact]
    public void Validate_ShouldRejectSheetNameOver31Chars()
    {
        var longName = new string('X', 32);
        var json = $$"""{"version":"1.0","worksheets":[{"name":"{{longName}}","rows":[["1"]]}]}""";

        var ex = Assert.Throws<XlsxException>(() => XlsxInstructionParser.Parse(json));
        Assert.Contains("31", ex.Message);
    }

    [Theory]
    [InlineData("Sheet:1")]
    [InlineData("Path/Sheet")]
    [InlineData("Why?Yes")]
    [InlineData("Bracket[Sheet")]
    public void Validate_ShouldRejectIllegalSheetCharacters(string name)
    {
        var json = $$"""{"version":"1.0","worksheets":[{"name":"{{name}}","rows":[["1"]]}]}""";

        var ex = Assert.Throws<XlsxException>(() => XlsxInstructionParser.Parse(json));
        Assert.Contains("illegal", ex.Message.ToLowerInvariant());
    }

    [Fact]
    public void Validate_ShouldRejectCellWithBothValueAndFormula()
    {
        var json = """
        {
            "version": "1.0",
            "worksheets": [{
                "name": "S",
                "cells": [{"address": "A1", "value": "x", "formula": "=SUM(B1:B2)"}]
            }]
        }
        """;

        var ex = Assert.Throws<XlsxException>(() => XlsxInstructionParser.Parse(json));
        Assert.Contains("both", ex.Message.ToLowerInvariant());
    }

    [Fact]
    public void Validate_ShouldRejectCellWithNeitherValueNorFormula()
    {
        var json = """
        {
            "version": "1.0",
            "worksheets": [{
                "name": "S",
                "cells": [{"address": "A1"}]
            }]
        }
        """;

        var ex = Assert.Throws<XlsxException>(() => XlsxInstructionParser.Parse(json));
        Assert.Contains("neither", ex.Message.ToLowerInvariant());
    }

    [Fact]
    public void Validate_ShouldRejectFormulaNotStartingWithEquals()
    {
        var json = """
        {
            "version": "1.0",
            "worksheets": [{
                "name": "S",
                "cells": [{"address": "A1", "formula": "SUM(B1:B2)"}]
            }]
        }
        """;

        var ex = Assert.Throws<XlsxException>(() => XlsxInstructionParser.Parse(json));
        Assert.Contains("start with '='", ex.Message);
    }

    [Fact]
    public void Validate_ShouldRejectInvalidCellAddress()
    {
        var json = """
        {
            "version": "1.0",
            "worksheets": [{
                "name": "S",
                "cells": [{"address": "A", "value": "bad"}]
            }]
        }
        """;

        var ex = Assert.Throws<XlsxException>(() => XlsxInstructionParser.Parse(json));
        Assert.Contains("Invalid cell address", ex.Message);
    }

    [Fact]
    public void Validate_ShouldRejectUnknownCellType()
    {
        var json = """
        {
            "version": "1.0",
            "worksheets": [{
                "name": "S",
                "cells": [{"address": "A1", "value": "x", "type": "currency"}]
            }]
        }
        """;

        var ex = Assert.Throws<XlsxException>(() => XlsxInstructionParser.Parse(json));
        Assert.Contains("currency", ex.Message);
    }

    [Fact]
    public void Validate_ShouldRejectEmptyVariablesBlock()
    {
        var json = """
        {
            "version": "1.0",
            "worksheets": [{"name": "S", "rows": [["1"]]}],
            "variables": {}
        }
        """;

        var ex = Assert.Throws<XlsxException>(() => XlsxInstructionParser.Parse(json));
        Assert.Contains("empty", ex.Message.ToLowerInvariant());
    }

    // ─── Executor ──────────────────────────────────────────────────

    [Fact]
    public void Execute_ShouldBuildWorksheetFromHeadersAndRows()
    {
        var set = XlsxInstructionParser.Parse("""
        {
            "version": "1.0",
            "worksheets": [{
                "name": "Sales",
                "headers": ["Product", "Q1"],
                "rows": [["Widget", "100"]]
            }]
        }
        """);

        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            XlsxInstructionExecutor.Execute(set, builder);
            builder.Save();
        }

        using (var reader = WorkbookBuilder.Open(_testFilePath))
        {
            var ws = reader.GetWorksheet("Sales");
            Assert.Equal("Product", ws.GetCellValue("A1"));
            Assert.Equal("Q1", ws.GetCellValue("B1"));
            Assert.Equal("Widget", ws.GetCellValue("A2"));
            Assert.Equal("100", ws.GetCellValue("B2"));
        }

        OpenXmlAssert.NoValidationErrors(_testFilePath);
    }

    [Fact]
    public void Execute_ShouldAddCellValuesAndFormulas()
    {
        var set = XlsxInstructionParser.Parse("""
        {
            "version": "1.0",
            "worksheets": [{
                "name": "Calc",
                "cells": [
                    {"address": "A1", "value": "10"},
                    {"address": "A2", "value": "20"},
                    {"address": "A3", "formula": "=SUM(A1:A2)"}
                ]
            }]
        }
        """);

        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            XlsxInstructionExecutor.Execute(set, builder);
            builder.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        var ws = reader.GetWorksheet("Calc");
        Assert.Equal("10", ws.GetCellValue("A1"));
        Assert.Equal("20", ws.GetCellValue("A2"));
        Assert.Equal("=SUM(A1:A2)", ws.GetCellFormula("A3"));
    }

    [Fact]
    public void Execute_ShouldResolveVariablesInRowValues()
    {
        var set = XlsxInstructionParser.Parse("""
        {
            "version": "1.0",
            "worksheets": [{
                "name": "Template",
                "headers": ["Field", "Value"],
                "rows": [
                    ["Company", "{{name}}"]
                ]
            }],
            "variables": {"name": "Acme Corp"}
        }
        """);

        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            XlsxInstructionExecutor.Execute(set, builder);
            builder.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        var ws = reader.GetWorksheet("Template");
        Assert.Equal("Acme Corp", ws.GetCellValue("B2"));
    }

    [Fact]
    public void Execute_ShouldResolveVariablesInFormulaCells()
    {
        var set = XlsxInstructionParser.Parse("""
        {
            "version": "1.0",
            "worksheets": [{
                "name": "Calc",
                "cells": [
                    {"address": "A1", "value": "10"},
                    {"address": "B1", "formula": "=A1*{{factor}}"}
                ]
            }],
            "variables": {"factor": "2"}
        }
        """);

        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            XlsxInstructionExecutor.Execute(set, builder);
            builder.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        Assert.Equal("=A1*2", reader.GetWorksheet("Calc").GetCellFormula("B1"));
    }

    [Fact]
    public void Execute_ShouldBuildMultipleWorksheets()
    {
        var set = XlsxInstructionParser.Parse("""
        {
            "version": "1.0",
            "worksheets": [
                {"name": "First", "headers": ["A"], "rows": [["1"]]},
                {"name": "Second", "headers": ["B"], "rows": [["2"]]}
            ]
        }
        """);

        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            XlsxInstructionExecutor.Execute(set, builder);
            builder.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        Assert.Equal("1", reader.GetWorksheet("First").GetCellValue("A2"));
        Assert.Equal("2", reader.GetWorksheet("Second").GetCellValue("A2"));
    }

    [Fact]
    public void Execute_ShouldBuildFormulaRows()
    {
        var set = XlsxInstructionParser.Parse("""
        {
            "version": "1.0",
            "worksheets": [{
                "name": "Sales",
                "headers": ["A", "B", "Total"],
                "rows": [
                    ["1", "2", "=SUM(A2:B2)"]
                ]
            }]
        }
        """);

        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            XlsxInstructionExecutor.Execute(set, builder);
            builder.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        Assert.Equal("=SUM(A2:B2)", reader.GetWorksheet("Sales").GetCellFormula("C2"));
    }

    [Fact]
    public void Execute_ShouldHandleCellWithStyle()
    {
        var set = XlsxInstructionParser.Parse("""
        {
            "version": "1.0",
            "worksheets": [{
                "name": "Styled",
                "headers": ["H"],
                "cells": [
                    {"address": "A2", "value": "7", "style": "0"}
                ]
            }]
        }
        """);

        // The header row creates the stylesheet, so styleId "0" (default format) is valid.
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            XlsxInstructionExecutor.Execute(set, builder);
            builder.Save();
        }

        using (var reader = WorkbookBuilder.Open(_testFilePath))
        {
            var ws = reader.GetWorksheet("Styled");
            Assert.Equal("7", ws.GetCellValue("A2"));
        }

        OpenXmlAssert.NoValidationErrors(_testFilePath);
    }

    [Fact]
    public void Execute_ShouldSkipSheetWithNoRowsOrCells()
    {
        var set = XlsxInstructionParser.Parse("""
        {
            "version": "1.0",
            "worksheets": [{
                "name": "HeadersOnly",
                "headers": ["A", "B", "C"]
            }]
        }
        """);

        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            XlsxInstructionExecutor.Execute(set, builder);
            builder.Save();
        }

        using var reader = WorkbookBuilder.Open(_testFilePath);
        var ws = reader.GetWorksheet("HeadersOnly");
        Assert.True(ws.CellExists("A1"));
        Assert.False(ws.CellExists("A2"));
    }

    [Fact]
    public void EndToEnd_SampleJson_ShouldProduceValidWorkbook()
    {
        var json = """
        {
            "version": "1.0",
            "description": "Instructions for creating a budget workbook",
            "worksheets": [
                {
                    "name": "Revenue",
                    "headers": ["Month", "Product A", "Product B", "Product C", "Total"],
                    "rows": [
                        ["January", "{{janA}}", "{{janB}}", "{{janC}}", "=SUM(B2:D2)"],
                        ["February", "{{febA}}", "{{febB}}", "{{febC}}", "=SUM(B3:D3)"],
                        ["March", "{{marA}}", "{{marB}}", "{{marC}}", "=SUM(B4:D4)"]
                    ]
                },
                {
                    "name": "Summary",
                    "cells": [
                        {"address": "A1", "value": "Total Revenue Q1"},
                        {"address": "B1", "formula": "=SUM(Revenue!E2:E4)"}
                    ]
                }
            ],
            "variables": {
                "janA": "10000", "janB": "8000", "janC": "6000",
                "febA": "12000", "febB": "9500", "febC": "7000",
                "marA": "11000", "marB": "9000", "marC": "7500"
            }
        }
        """;

        var set = XlsxInstructionParser.Parse(json);

        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            XlsxInstructionExecutor.Execute(set, builder);
            builder.Save();
        }

        using (var reader = WorkbookBuilder.Open(_testFilePath))
        {
            var revenue = reader.GetWorksheet("Revenue");
            Assert.Equal("Month", revenue.GetCellValue("A1"));
            Assert.Equal("10000", revenue.GetCellValue("B2"));
            Assert.Equal("=SUM(B2:D2)", revenue.GetCellFormula("E2"));

            var summary = reader.GetWorksheet("Summary");
            Assert.Equal("Total Revenue Q1", summary.GetCellValue("A1"));
            Assert.Equal("=SUM(Revenue!E2:E4)", summary.GetCellFormula("B1"));
        }

        OpenXmlAssert.NoValidationErrors(_testFilePath);
    }

    public void Dispose()
    {
        if (File.Exists(_testFilePath))
        {
            File.Delete(_testFilePath);
        }
    }
}
