using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxEditor.Core.Builders;
using XlsxEditor.Core.Exceptions;
using XlsxEditor.Core.Instructions;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Focused tests for the JSON/model → XLSX generation path: the high-level
/// <see cref="XlsxGenerator"/> and the plan-first <see cref="XlsxInstructionExecutor"/>.
/// Covers legacy compatibility, named-style dedup, typed cells, formulas with styles,
/// metadata, layout (merges / freeze panes / autofilter / tables / widths / heights),
/// explicit empty-string cells, variable handling, atomic output, stream ownership, and
/// OpenXmlValidator conformance of produced packages.
/// </summary>
public class XlsxGeneratorTests
{
    private static string FixturePath(string relative) =>
        Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", relative));

    private static SpreadsheetDocument OpenBytes(byte[] bytes)
    {
        var stream = new MemoryStream(bytes);
        return SpreadsheetDocument.Open(stream, false);
    }

    private static void AssertNoValidationErrors(byte[] bytes)
    {
        using var document = OpenBytes(bytes);
        OpenXmlAssert.NoValidationErrors(document);
    }

    // ─── Generator: result shape & diagnostics ─────────────────────

    [Fact]
    public void Generate_ShouldBuildValidWorkbook_FromJson()
    {
        var result = XlsxGenerator.Generate("""
        {
            "version": "1.0",
            "worksheets": [
                { "name": "Sheet1", "headers": ["A", "B"], "rows": [["1", "2"]] }
            ]
        }
        """);

        Assert.True(result.IsValid);
        Assert.NotNull(result.Bytes);
        Assert.False(result.HasErrors);

        using var reader = WorkbookBuilder.Open(result.Bytes!);
        var ws = reader.GetWorksheet("Sheet1");
        Assert.Equal("A", ws.GetCellValue("A1"));
        Assert.Equal("2", ws.GetCellValue("B2"));
        AssertNoValidationErrors(result.Bytes!);
    }

    [Fact]
    public void Generate_ShouldReturnDiagnostics_ForInvalidInstructionSet()
    {
        var result = XlsxGenerator.Generate("""
        { "version": "2.0", "worksheets": [{"name": "S"}] }
        """);

        Assert.False(result.IsValid);
        Assert.Null(result.Bytes);
        Assert.Contains(result.Validation.Errors, d => d.Code == XlsxDiagnosticCode.UnsupportedVersion);
    }

    [Fact]
    public void Generate_ShouldReturnDiagnostics_ForUnresolvedVariable()
    {
        var result = XlsxGenerator.Generate("""
        {
            "version": "1.0",
            "worksheets": [{"name": "S", "rows": [["{{missing}}"]]}]
        }
        """);

        Assert.False(result.IsValid);
        Assert.Null(result.Bytes);
        Assert.Contains(result.Validation.Errors, d => d.Code == XlsxDiagnosticCode.UnresolvedVariable);
    }

    [Fact]
    public void Generate_ShouldRejectVariableCycle_WithNoOutput()
    {
        var result = XlsxGenerator.Generate("""
        {
            "version": "1.0",
            "variables": {"a": "{{b}}", "b": "{{a}}"},
            "worksheets": [{"name": "S", "rows": [["{{a}}"]]}]
        }
        """);

        Assert.False(result.IsValid);
        Assert.Null(result.Bytes);
        Assert.Contains(result.Validation.Errors, d => d.Code == XlsxDiagnosticCode.VariableCycle);
    }

    [Fact]
    public void GenerateBytes_ShouldThrowFirstError_ForInvalidSet()
    {
        var ex = Assert.Throws<XlsxException>(() =>
            XlsxGenerator.GenerateBytes("""{ "version": "9.9", "worksheets": [{"name": "S"}] }"""));
        Assert.Contains("version", ex.Message);
    }

    // ─── Generator: values, types, styles, formulas ────────────────

    [Fact]
    public void Generate_ShouldWriteTypedCells()
    {
        var result = XlsxGenerator.Generate("""
        {
            "version": "1.0",
            "worksheets": [{
                "name": "S",
                "cells": [
                    { "address": "A1", "value": "hello", "type": "string" },
                    { "address": "A2", "value": "42", "type": "number" },
                    { "address": "A3", "value": "true", "type": "boolean" },
                    { "address": "A4", "value": "2026-09-30", "type": "date" },
                    { "address": "A5", "value": "2026-09-30T14:30:00", "type": "datetime" }
                ]
            }]
        }
        """);

        Assert.True(result.IsValid);
        using var reader = WorkbookBuilder.Open(result.Bytes!);
        var ws = reader.GetWorksheet("S");
        Assert.Equal(CellValues.SharedString, ws.GetCellInfo("A1")!.DataType);
        Assert.Equal(CellValues.Number, ws.GetCellInfo("A2")!.DataType);
        Assert.Equal(CellValues.Boolean, ws.GetCellInfo("A3")!.DataType);
        Assert.Equal(CellValues.Number, ws.GetCellInfo("A4")!.DataType);
        Assert.Equal(CellValues.Number, ws.GetCellInfo("A5")!.DataType);
        AssertNoValidationErrors(result.Bytes!);
    }

    [Fact]
    public void Generate_ShouldHandleExplicitEmptyStringCell()
    {
        var result = XlsxGenerator.Generate("""
        {
            "version": "1.0",
            "worksheets": [{
                "name": "S",
                "cells": [{"address": "A1", "value": "", "type": "string"}]
            }]
        }
        """);

        Assert.True(result.IsValid);
        using var reader = WorkbookBuilder.Open(result.Bytes!);
        var ws = reader.GetWorksheet("S");
        Assert.True(ws.CellExists("A1"));
        Assert.Equal(string.Empty, ws.GetCellValue("A1"));
        AssertNoValidationErrors(result.Bytes!);
    }

    [Fact]
    public void Generate_ShouldResolveVariables_InValuesAndFormulas()
    {
        var result = XlsxGenerator.Generate("""
        {
            "version": "1.0",
            "variables": {"factor": "3"},
            "worksheets": [{
                "name": "S",
                "cells": [
                    { "address": "A1", "value": "{{factor}}", "type": "number" },
                    { "address": "B1", "formula": "=A1*{{factor}}" }
                ]
            }]
        }
        """);

        Assert.True(result.IsValid);
        using var reader = WorkbookBuilder.Open(result.Bytes!);
        var ws = reader.GetWorksheet("S");
        Assert.Equal("3", ws.GetCellValue("A1"));
        Assert.Equal("=A1*3", ws.GetCellFormula("B1"));
    }

    [Fact]
    public void Generate_ShouldApplyNamedStyles_AndDeduplicateStructurallyIdenticalOnes()
    {
        var result = XlsxGenerator.Generate("""
        {
            "version": "1.0",
            "styles": [
                { "name": "money", "numberFormat": "$#,##0.00" },
                { "name": "money2", "numberFormat": "$#,##0.00" }
            ],
            "worksheets": [{
                "name": "S",
                "columns": [
                    { "style": "money", "type": "number" },
                    { "style": "money2", "type": "number" }
                ],
                "rows": [["1", "2"]]
            }]
        }
        """);

        Assert.True(result.IsValid);
        AssertNoValidationErrors(result.Bytes!);

        using var document = OpenBytes(result.Bytes!);
        // Default cellXf (0) plus a single deduplicated money format: two distinct
        // named styles with identical content must share one cellXf, not duplicate it.
        var stylesheet = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
        Assert.Equal(2u, stylesheet.CellFormats!.Count!.Value);

        var sheet = document.WorkbookPart.WorksheetParts.First().Worksheet!;
        var a1 = sheet.Descendants<Cell>().Single(c => c.CellReference == "A1");
        var b1 = sheet.Descendants<Cell>().Single(c => c.CellReference == "B1");
        Assert.NotNull(a1.StyleIndex);
        Assert.Equal(a1.StyleIndex, b1.StyleIndex);
        Assert.NotEqual(0u, a1.StyleIndex!.Value);
    }

    [Fact]
    public void Generate_ShouldApplyStylesAndNumberFormats_ToFormulaCells()
    {
        var result = XlsxGenerator.Generate("""
        {
            "version": "1.0",
            "styles": [{"name": "money", "numberFormat": "$#,##0.00"}],
            "worksheets": [{
                "name": "S",
                "rows": [["1", "2"]],
                "cells": [
                    { "address": "C1", "formula": "=SUM(A1:B1)", "style": "money", "numberFormat": "$#,##0.00" }
                ]
            }]
        }
        """);

        Assert.True(result.IsValid);
        using var reader = WorkbookBuilder.Open(result.Bytes!);
        Assert.Equal("=SUM(A1:B1)", reader.GetWorksheet("S").GetCellFormula("C1"));

        using var document = OpenBytes(result.Bytes!);
        var cell = document.WorkbookPart!.WorksheetParts.First().Worksheet!
            .Descendants<Cell>().Single(c => c.CellReference == "C1");
        Assert.NotNull(cell.StyleIndex);
        Assert.NotEqual(0u, cell.StyleIndex!.Value);
        AssertNoValidationErrors(result.Bytes!);
    }

    [Fact]
    public void Generate_ShouldEmitWorkbookMetadata()
    {
        var result = XlsxGenerator.Generate("""
        {
            "version": "1.0",
            "metadata": {
                "title": "Q1 Report",
                "subject": "Quarterly",
                "author": "Acme",
                "category": "Finance",
                "keywords": "budget,revenue",
                "comments": "Draft"
            },
            "worksheets": [{"name": "S", "rows": [["1"]]}]
        }
        """);

        Assert.True(result.IsValid);
        using var document = OpenBytes(result.Bytes!);
        Assert.Equal("Q1 Report", document.PackageProperties.Title);
        Assert.Equal("Quarterly", document.PackageProperties.Subject);
        Assert.Equal("Acme", document.PackageProperties.Creator);
        Assert.Equal("Finance", document.PackageProperties.Category);
        Assert.Equal("Draft", document.PackageProperties.Description);
        AssertNoValidationErrors(result.Bytes!);
    }

    // ─── Generator: layout (merges / panes / autofilter / tables) ──

    [Fact]
    public void Generate_ShouldApplyMerges_FreezePanes_AutoFilter_Tables_Widths_Heights()
    {
        var result = XlsxGenerator.Generate("""
        {
            "version": "1.0",
            "worksheets": [{
                "name": "S",
                "headers": ["a", "b", "c"],
                "rows": [["1", "2", "3"]],
                "columns": [{"width": 20}],
                "rowHeights": [{"row": 1, "height": 30}],
                "merges": ["A3:C3"],
                "freezePanes": {"row": 2, "column": 1},
                "autoFilter": "E1:F2",
                "tables": [{"name": "DataTable", "range": "A1:C2"}]
            }]
        }
        """);

        Assert.True(result.IsValid);
        using var reader = WorkbookBuilder.Open(result.Bytes!);
        var ws = reader.GetWorksheet("S");
        Assert.Equal(new List<string> { "A3:C3" }, ws.GetMergeRanges());
        Assert.Equal((1, 0), ws.GetFreezePanes());
        Assert.Equal("E1:F2", ws.GetAutoFilterRange());
        Assert.Equal(20d, ws.GetColumnWidth("A"));
        Assert.Equal(30d, ws.GetRowHeight(1));

        using var document = OpenBytes(result.Bytes!);
        var tablePart = document.WorkbookPart!.WorksheetParts.First().TableDefinitionParts.First();
        Assert.Equal("DataTable", tablePart.Table!.DisplayName!.Value);
        Assert.Equal("A1:C2", tablePart.Table.Reference!.Value);
        AssertNoValidationErrors(result.Bytes!);
    }

    // ─── Rich-report fixture (three sheets) ────────────────────────

    [Fact]
    public void GenerateFromFile_ShouldBuildValidRichReport_FromFixture()
    {
        var fixturePath = FixturePath("examples/Xlsx/instructions/rich-report.json");
        var result = XlsxGenerator.GenerateFromFile(fixturePath);

        Assert.True(result.IsValid, string.Join("; ", result.Validation.Errors.Select(e => e.Message)));
        Assert.NotNull(result.Bytes);
        AssertNoValidationErrors(result.Bytes!);

        using var reader = WorkbookBuilder.Open(result.Bytes!);
        var summary = reader.GetWorksheet("Summary");
        Assert.Equal("Q3 2026 — Analytics Team", summary.GetCellValue("A1"));
        Assert.Equal("=SUM(B3:B6)", summary.GetCellFormula("F2"));
        Assert.Equal(new List<string> { "A1:E1" }, summary.GetMergeRanges());
        Assert.Equal((2, 0), summary.GetFreezePanes());
        Assert.Null(summary.GetAutoFilterRange());
        Assert.Equal(CellValues.Number, summary.GetCellInfo("B3")!.DataType);
        Assert.Equal(CellValues.Boolean, summary.GetCellInfo("D3")!.DataType);
        Assert.Equal(CellValues.Number, summary.GetCellInfo("E3")!.DataType);
        Assert.Equal(16d, summary.GetColumnWidth("A"));
        Assert.Equal(24d, summary.GetRowHeight(2));

        var breakdown = reader.GetWorksheet("Breakdown");
        Assert.Equal("A1:E4", breakdown.GetAutoFilterRange());
        Assert.Equal("=SUM(B2:B4)", breakdown.GetCellFormula("B5"));
        Assert.True(breakdown.CellExists("A7"));
        Assert.Equal(string.Empty, breakdown.GetCellValue("A7"));
        Assert.Equal("Legacy-styled", breakdown.GetCellValue("A8"));
        Assert.Equal(CellValues.Number, breakdown.GetCellInfo("B8")!.DataType);
        Assert.Equal(CellValues.Number, breakdown.GetCellInfo("E2")!.DataType);

        var notes = reader.GetWorksheet("Notes");
        Assert.Equal("Q3 2026 — Analytics Team", notes.GetCellValue("A1"));
        Assert.Equal("=B3*2", notes.GetCellFormula("E3"));
        Assert.Equal(CellValues.Number, notes.GetCellInfo("B3")!.DataType);
        Assert.Equal(CellValues.Boolean, notes.GetCellInfo("C3")!.DataType);

        using var document = OpenBytes(result.Bytes!);
        var tablePart = document.WorkbookPart!.WorksheetParts.First().TableDefinitionParts.First();
        Assert.Equal("SummaryTable", tablePart.Table!.DisplayName!.Value);
        Assert.Equal("A2:E6", tablePart.Table.Reference!.Value);
        Assert.Equal("Northwind Labs Quarterly Report", document.PackageProperties.Title);
        Assert.Equal("OfficeEditor Generator", document.PackageProperties.Creator);
    }

    [Fact]
    public void GenerateFromFile_ShouldValidatePackage_WhenRequested()
    {
        var fixturePath = FixturePath("examples/Xlsx/instructions/rich-report.json");
        var result = XlsxGenerator.GenerateFromFile(
            fixturePath, new XlsxGenerateOptions { ValidatePackage = true });

        Assert.True(result.IsValid, string.Join("; ", result.Validation.Errors.Select(e => e.Message)));
        Assert.NotNull(result.Bytes);
        Assert.False(result.HasWarnings);
    }

    // ─── Executor: no-partial-output semantics ─────────────────────

    [Fact]
    public void Execute_ShouldNotMutateBuilder_WhenValueConversionFails()
    {
        var set = new XlsxInstructionSet
        {
            Version = "1.0",
            Worksheets =
            [
                new WorksheetInstruction
                {
                    Name = "S",
                    Cells =
                    [
                        new CellInstruction { Address = "A1", Value = "ok", Type = "string" },
                        new CellInstruction { Address = "A2", Value = "not-a-number", Type = "number" }
                    ]
                }
            ]
        };

        using var builder = WorkbookBuilder.Create();
        var ex = Assert.Throws<XlsxException>(() => XlsxInstructionExecutor.Execute(set, builder));
        Assert.Contains("number", ex.Message);
        Assert.Empty(builder.GetWorksheetNames());
    }

    [Fact]
    public void Execute_ShouldNotMutateBuilder_WhenStyleAspectUnsupported()
    {
        var set = new XlsxInstructionSet
        {
            Version = "1.0",
            Styles = [new NamedStyle { Name = "bordered", Border = new BorderStyleInstruction() }],
            Worksheets =
            [
                new WorksheetInstruction
                {
                    Name = "S",
                    Cells = [new CellInstruction { Address = "A1", Value = "x", Style = "bordered" }]
                }
            ]
        };

        using var builder = WorkbookBuilder.Create();
        var ex = Assert.Throws<XlsxException>(() => XlsxInstructionExecutor.Execute(set, builder));
        Assert.Contains("borders", ex.Message);
        Assert.Empty(builder.GetWorksheetNames());
    }

    // ─── Outputs: file (atomic) & stream ownership ─────────────────

    [Fact]
    public void GenerateToFile_ShouldWriteAtomically_AndOverwrite()
    {
        var output = Path.Combine(Path.GetTempPath(), $"gen_{Guid.NewGuid():N}.xlsx");
        try
        {
            var result = XlsxGenerator.GenerateToFile(
                """{ "version": "1.0", "worksheets": [{"name": "S", "rows": [["1"]]}] }""", output);
            Assert.True(result.IsValid);
            Assert.True(File.Exists(output));
            OpenXmlAssert.NoValidationErrors(output);

            var overwritten = XlsxGenerator.GenerateToFile(
                """{ "version": "1.0", "worksheets": [{"name": "T", "rows": [["2"]]}] }""", output);
            Assert.True(overwritten.IsValid);
            using var reader = WorkbookBuilder.Open(output);
            Assert.Contains("T", reader.GetWorksheetNames());

            Assert.Empty(Directory.GetFiles(Path.GetDirectoryName(output)!, $".{Path.GetFileName(output)}.*.tmp"));
        }
        finally
        {
            if (File.Exists(output))
            {
                File.Delete(output);
            }
        }
    }

    [Fact]
    public void GenerateToFile_ShouldNotWrite_WhenGenerationFails()
    {
        var output = Path.Combine(Path.GetTempPath(), $"gen_{Guid.NewGuid():N}.xlsx");
        try
        {
            var result = XlsxGenerator.GenerateToFile(
                """{ "version": "9.9", "worksheets": [{"name": "S"}] }""", output);
            Assert.False(result.IsValid);
            Assert.False(File.Exists(output));
        }
        finally
        {
            if (File.Exists(output))
            {
                File.Delete(output);
            }
        }
    }

    [Fact]
    public void GenerateToFile_ShouldPreserveExistingFile_WhenGenerationFails()
    {
        var output = Path.Combine(Path.GetTempPath(), $"gen_{Guid.NewGuid():N}.xlsx");
        try
        {
            File.WriteAllText(output, "keep me");
            var result = XlsxGenerator.GenerateToFile(
                """{ "version": "9.9", "worksheets": [{"name": "S"}] }""", output);
            Assert.False(result.IsValid);
            Assert.Equal("keep me", File.ReadAllText(output));
        }
        finally
        {
            if (File.Exists(output))
            {
                File.Delete(output);
            }
        }
    }

    [Fact]
    public void GenerateToStream_ShouldLeaveStreamOpen_AndPositionedAtEnd()
    {
        using var stream = new MemoryStream();
        var result = XlsxGenerator.GenerateToStream(
            """{ "version": "1.0", "worksheets": [{"name": "S", "rows": [["1"]]}] }""", stream);

        Assert.True(result.IsValid);
        Assert.True(stream.CanRead);
        Assert.Equal(stream.Length, stream.Position);

        using var reader = WorkbookBuilder.Open(stream.ToArray());
        Assert.Equal("1", reader.GetWorksheet("S").GetCellValue("A1"));
    }
}
