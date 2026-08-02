using XlsxEditor.Core.Exceptions;
using XlsxEditor.Core.Instructions;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Covers the structured, aggregate validation API: path-qualified diagnostics with
/// code/severity/message, the non-throwing parser entry point, and the throwing adapter
/// that preserves the legacy <see cref="XlsxException"/> surface.
/// </summary>
public class XlsxDiagnosticsTests
{
    [Fact]
    public void Validate_ShouldAggregateMultipleErrors_WithPathAndCode()
    {
        var set = XlsxInstructionParser.ParseAndValidate("""
        {
            "version": "1.0",
            "worksheets": [
                {"name": "Sales", "rows": [["1"]]},
                {"name": "SALES", "cells": [{"address": "A", "value": "x"}]}
            ]
        }
        """).InstructionSet!;

        var result = XlsxValidationEngine.Validate(set);

        Assert.False(result.IsValid);
        Assert.True(result.HasErrors);
        var errors = result.Errors.ToList();
        Assert.Equal(2, errors.Count);

        var duplicate = Assert.Single(errors, d => d.Code == XlsxDiagnosticCode.SheetNameDuplicate);
        Assert.Equal(XlsxDiagnosticSeverity.Error, duplicate.Severity);
        Assert.Equal("worksheets[1].name", duplicate.Path);

        var badAddress = Assert.Single(errors, d => d.Code == XlsxDiagnosticCode.CellAddressInvalid);
        Assert.Equal("worksheets[1].cells[0].address", badAddress.Path);
        Assert.Contains("Invalid cell address 'A'", badAddress.Message);
    }

    [Fact]
    public void Validate_ShouldFlagDuplicateRowHeight_AsNonBlockingWarning()
    {
        var set = XlsxInstructionParser.Parse("""
        {
            "version": "1.0",
            "worksheets": [{
                "name": "S",
                "rows": [["1"]],
                "rowHeights": [{"row": 1, "height": 20}, {"row": 1, "height": 24}]
            }]
        }
        """);

        var result = XlsxValidationEngine.Validate(set);
        Assert.True(result.IsValid);
        Assert.True(result.HasWarnings);
        var warning = Assert.Single(result.Warnings, d => d.Code == XlsxDiagnosticCode.DuplicateRowHeight);
        Assert.Equal(XlsxDiagnosticSeverity.Warning, warning.Severity);
        Assert.Equal("worksheets[0].rowHeights[1].row", warning.Path);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(300)]
    [InlineData(double.NaN)]
    public void Validate_ShouldRejectInvalidColumnWidth(double width)
    {
        var set = new XlsxInstructionSet
        {
            Version = "1.0",
            Worksheets = [new WorksheetInstruction { Name = "S", Rows = [["1"]], Columns = [new ColumnInstruction { Width = width }] }]
        };

        var result = XlsxValidationEngine.Validate(set);
        var error = Assert.Single(result.Errors);
        Assert.Equal(XlsxDiagnosticCode.ColumnWidthInvalid, error.Code);
        Assert.Equal("worksheets[0].columns[0].width", error.Path);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(410)]
    [InlineData(double.PositiveInfinity)]
    public void Validate_ShouldRejectInvalidRowHeight(double height)
    {
        var set = new XlsxInstructionSet
        {
            Version = "1.0",
            Worksheets = [new WorksheetInstruction { Name = "S", Rows = [["1"]], RowHeights = [new RowHeightInstruction { Row = 1, Height = height }] }]
        };

        var result = XlsxValidationEngine.Validate(set);
        var error = Assert.Single(result.Errors);
        Assert.Equal(XlsxDiagnosticCode.RowHeightInvalid, error.Code);
    }

    [Fact]
    public void Validate_ShouldRejectOutOfBoundsStartRow()
    {
        var set = new XlsxInstructionSet
        {
            Version = "1.0",
            Worksheets = [new WorksheetInstruction { Name = "S", Rows = [["1"]], StartRow = 1_048_577 }]
        };

        var result = XlsxValidationEngine.Validate(set);
        var error = Assert.Single(result.Errors, d => d.Code == XlsxDiagnosticCode.SheetStartRowOutOfBounds && d.Path == "worksheets[0].startRow");
        Assert.Equal("worksheets[0].startRow", error.Path);
    }

    [Fact]
    public void Validate_ShouldRejectOutOfBoundsCellAddress_WithDistinctCode()
    {
        var set = new XlsxInstructionSet
        {
            Version = "1.0",
            Worksheets = [new WorksheetInstruction { Name = "S", Cells = [new CellInstruction { Address = "XFE1", Value = "x" }] }]
        };

        var result = XlsxValidationEngine.Validate(set);
        var error = Assert.Single(result.Errors);
        Assert.Equal(XlsxDiagnosticCode.CellAddressOutOfBounds, error.Code);
        Assert.Contains("Excel", error.Message);
    }

    [Fact]
    public void Validate_ShouldRejectUnknownCellStyle_WithPath()
    {
        var set = new XlsxInstructionSet
        {
            Version = "1.0",
            Worksheets = [new WorksheetInstruction { Name = "S", Cells = [new CellInstruction { Address = "A1", Value = "x", Style = "nope" }] }]
        };

        var result = XlsxValidationEngine.Validate(set);
        var error = Assert.Single(result.Errors);
        Assert.Equal(XlsxDiagnosticCode.CellStyleUnknown, error.Code);
        Assert.Equal("worksheets[0].cells[0].style", error.Path);
    }

    [Fact]
    public void Validate_ShouldRejectSingleCellMerge()
    {
        var set = new XlsxInstructionSet
        {
            Version = "1.0",
            Worksheets = [new WorksheetInstruction { Name = "S", Headers = ["a"], Merges = ["B2:B2"] }]
        };

        var result = XlsxValidationEngine.Validate(set);
        var error = Assert.Single(result.Errors);
        Assert.Equal(XlsxDiagnosticCode.MergeSingleCell, error.Code);
    }

    [Fact]
    public void Validate_ShouldRejectReversedMerge()
    {
        var set = new XlsxInstructionSet
        {
            Version = "1.0",
            Worksheets = [new WorksheetInstruction { Name = "S", Headers = ["a"], Merges = ["C3:A1"] }]
        };

        var result = XlsxValidationEngine.Validate(set);
        var error = Assert.Single(result.Errors);
        Assert.Equal(XlsxDiagnosticCode.MergeReversed, error.Code);
    }

    [Fact]
    public void Validate_ShouldRejectOutOfBoundsFreezePanes()
    {
        var set = new XlsxInstructionSet
        {
            Version = "1.0",
            Worksheets = [new WorksheetInstruction { Name = "S", Headers = ["a"], FreezePanes = new FreezePanesInstruction { Row = 2_000_000, Column = 1 } }]
        };

        var result = XlsxValidationEngine.Validate(set);
        var error = Assert.Single(result.Errors);
        Assert.Equal(XlsxDiagnosticCode.FreezePanesOutOfBounds, error.Code);
    }

    [Fact]
    public void Validate_ShouldRejectWorkbookWideDuplicateTableNames()
    {
        var set = new XlsxInstructionSet
        {
            Version = "1.0",
            Worksheets =
            [
                new WorksheetInstruction { Name = "A", Headers = ["h"], Rows = [["1"]], Tables = [new TableInstruction { Name = "Shared", Range = "A1:B2" }] },
                new WorksheetInstruction { Name = "B", Headers = ["h"], Rows = [["1"]], Tables = [new TableInstruction { Name = "shared", Range = "A1:B2" }] }
            ]
        };

        var result = XlsxValidationEngine.Validate(set);
        var error = Assert.Single(result.Errors, d => d.Code == XlsxDiagnosticCode.TableDuplicateName);
        Assert.Equal("worksheets[1].tables[0].name", error.Path);
    }

    [Fact]
    public void Validate_ShouldRejectEmptyWorksheet()
    {
        var set = new XlsxInstructionSet
        {
            Version = "1.0",
            Worksheets = [new WorksheetInstruction { Name = "Blank" }]
        };

        var result = XlsxValidationEngine.Validate(set);
        var error = Assert.Single(result.Errors);
        Assert.Equal(XlsxDiagnosticCode.SheetEmpty, error.Code);
    }

    // ─── Parser adapter ───────────────────────────────────────────

    [Fact]
    public void ParseAndValidate_ShouldNeverThrow_ForSemanticErrors()
    {
        var result = XlsxInstructionParser.ParseAndValidate("""
        {
            "version": "1.0",
            "worksheets": [{"name": "S", "cells": [{"address": "A", "value": "x"}]}]
        }
        """);

        Assert.False(result.IsValid);
        Assert.NotNull(result.InstructionSet);
        Assert.Contains(result.Validation.Errors, d => d.Code == XlsxDiagnosticCode.CellAddressInvalid);
    }

    [Fact]
    public void ParseAndValidate_ShouldReturnInvalidJsonDiagnostic_ForSyntaxErrors()
    {
        var result = XlsxInstructionParser.ParseAndValidate("not json");

        Assert.False(result.IsValid);
        Assert.Null(result.InstructionSet);
        var error = Assert.Single(result.Validation.Errors);
        Assert.Equal(XlsxDiagnosticCode.InvalidJson, error.Code);
        Assert.Contains("Invalid JSON", error.Message);
    }

    [Fact]
    public void ParseAndValidate_ShouldReturnInvalidJsonDiagnostic_ForUnmappedMembers()
    {
        var result = XlsxInstructionParser.ParseAndValidate("""
        {
            "version": "1.0",
            "worksheets": [{"name": "S", "rows": [["1"]], "unknownThing": 1}]
        }
        """);

        Assert.False(result.IsValid);
        var error = Assert.Single(result.Validation.Errors);
        Assert.Equal(XlsxDiagnosticCode.InvalidJson, error.Code);
        Assert.Contains("unknownThing", error.Message);
    }

    [Fact]
    public void ThrowingValidator_ShouldThrowFirstError_WithPathSuffix()
    {
        var set = new XlsxInstructionSet
        {
            Version = "1.0",
            Worksheets = [new WorksheetInstruction { Name = "S", Cells = [new CellInstruction { Address = "A", Value = "x" }] }]
        };

        var ex = Assert.Throws<XlsxException>(() => XlsxInstructionValidator.Validate(set));
        Assert.Contains("Invalid cell address", ex.Message);
        Assert.Contains("(at worksheets[0].cells[0].address)", ex.Message);
    }

    [Fact]
    public void ThrowingValidator_ShouldPass_ForWarningsOnly()
    {
        var set = new XlsxInstructionSet
        {
            Version = "1.0",
            Worksheets =
            [
                new WorksheetInstruction
                {
                    Name = "S",
                    Rows = [["1"]],
                    RowHeights = [new RowHeightInstruction { Row = 1, Height = 20 }, new RowHeightInstruction { Row = 1, Height = 24 }]
                }
            ]
        };

        // Warnings must not throw — only errors block.
        XlsxInstructionValidator.Validate(set);
    }
}
