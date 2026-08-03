using XlsxEditor.Core.Instructions;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Covers the immutable preflight/execution plan: deterministic variable resolution with
/// cycle detection, cell type resolution, address/range normalization, limits, and
/// duplicate-write warnings. The planner performs no workbook mutation and depends on no
/// builder implementation — every test here plans a set without touching a builder.
/// </summary>
public class XlsxPlannerTests
{
    [Fact]
    public void Plan_ShouldResolveVariablesDeterministically_IntoConcreteCells()
    {
        var set = XlsxInstructionParser.Parse("""
        {
            "version": "1.0",
            "variables": {"b": "y", "a": "x{{b}}"},
            "worksheets": [{
                "name": "S",
                "headers": ["K"],
                "rows": [["{{a}}"]]
            }]
        }
        """);

        var result = XlsxPlanner.Plan(set);
        Assert.True(result.IsValid);
        var plan = result.Plan!;

        Assert.Equal("xy", plan.Variables["a"]);
        Assert.Equal("y", plan.Variables["b"]);
        Assert.Equal("xy", plan.Worksheets[0].Rows[0].Cells[0].Value);
        Assert.Null(plan.Worksheets[0].Rows[0].Cells[0].Formula);
    }

    [Fact]
    public void Plan_ShouldResolveFormulaCells_WithVariables()
    {
        var set = XlsxInstructionParser.Parse("""
        {
            "version": "1.0",
            "variables": {"rate": "2"},
            "worksheets": [{
                "name": "S",
                "cells": [{"address": "B1", "formula": "=A1*{{rate}}"}]
            }]
        }
        """);

        var result = XlsxPlanner.Plan(set);
        var cell = result.Plan!.Worksheets[0].Cells[0];
        Assert.Equal("=A1*2", cell.Formula);
        Assert.Null(cell.Value);
        Assert.Equal(XlsxCellType.Auto, cell.Type);
    }

    [Fact]
    public void Plan_ShouldDetectVariableCycle_AndProduceNoPlan()
    {
        var set = XlsxInstructionParser.Parse("""
        {
            "version": "1.0",
            "variables": {"a": "{{b}}", "b": "{{a}}"},
            "worksheets": [{"name": "S", "rows": [["{{a}}"]]}]
        }
        """);

        var result = XlsxPlanner.Plan(set);
        Assert.False(result.IsValid);
        Assert.Null(result.Plan);
        var cycle = Assert.Single(result.Validation.Errors, d => d.Code == XlsxDiagnosticCode.VariableCycle);
        Assert.Contains("a", cycle.Message);
        Assert.Contains("b", cycle.Message);
    }

    [Fact]
    public void Plan_ShouldFlagUndefinedVariableReference_WithinVariable()
    {
        var set = XlsxInstructionParser.Parse("""
        {
            "version": "1.0",
            "variables": {"a": "{{missing}}"},
            "worksheets": [{"name": "S", "rows": [["{{a}}"]]}]
        }
        """);

        var result = XlsxPlanner.Plan(set);
        Assert.False(result.IsValid);
        Assert.Null(result.Plan);
        var error = Assert.Single(result.Validation.Errors, d => d.Code == XlsxDiagnosticCode.UndefinedVariableReference);
        Assert.Equal("variables.a", error.Path);
    }

    [Fact]
    public void Plan_ShouldFlagUnresolvedVariable_InCell()
    {
        var set = XlsxInstructionParser.Parse("""
        {
            "version": "1.0",
            "worksheets": [{"name": "S", "rows": [["{{missing}}"]]}]
        }
        """);

        var result = XlsxPlanner.Plan(set);
        Assert.False(result.IsValid);
        var error = Assert.Single(result.Validation.Errors, d => d.Code == XlsxDiagnosticCode.UnresolvedVariable);
        Assert.Equal("worksheets[0].rows[0][0]", error.Path);
        Assert.Contains("{{missing}}", error.Message);
    }

    // ─── Type resolution ──────────────────────────────────────────

    [Fact]
    public void Plan_ShouldInferCellTypes_FromResolvedValues()
    {
        var set = XlsxInstructionParser.Parse("""
        {
            "version": "1.0",
            "worksheets": [{
                "name": "S",
                "rows": [["42", "true", "2024-01-15", "hello"]]
            }]
        }
        """);

        var cells = XlsxPlanner.Plan(set).Plan!.Worksheets[0].Rows[0].Cells;
        Assert.Equal(XlsxCellType.Number, cells[0].Type);
        Assert.Equal(XlsxCellType.Boolean, cells[1].Type);
        Assert.Equal(XlsxCellType.Date, cells[2].Type);
        Assert.Equal(XlsxCellType.String, cells[3].Type);
    }

    [Fact]
    public void Plan_ShouldHonorExplicitCellType_OverInference()
    {
        var set = XlsxInstructionParser.Parse("""
        {
            "version": "1.0",
            "worksheets": [{
                "name": "S",
                "cells": [{"address": "A1", "value": "42", "type": "string"}]
            }]
        }
        """);

        var cell = XlsxPlanner.Plan(set).Plan!.Worksheets[0].Cells[0];
        Assert.Equal(XlsxCellType.String, cell.Type);
    }

    [Fact]
    public void Plan_ShouldApplyColumnDefaultType_ToRowCells()
    {
        var set = XlsxInstructionParser.Parse("""
        {
            "version": "1.0",
            "worksheets": [{
                "name": "S",
                "columns": [{"type": "number"}],
                "rows": [["xyz"]]
            }]
        }
        """);

        var cell = XlsxPlanner.Plan(set).Plan!.Worksheets[0].Rows[0].Cells[0];
        Assert.Equal(XlsxCellType.Number, cell.Type);
    }

    [Fact]
    public void Plan_ShouldInheritColumnStyle_OnRowCells()
    {
        var set = XlsxInstructionParser.Parse("""
        {
            "version": "1.0",
            "styles": [{"name": "num", "numberFormat": "0.00"}],
            "worksheets": [{
                "name": "S",
                "columns": [{"style": "num"}],
                "rows": [["1.5"]]
            }]
        }
        """);

        var cell = XlsxPlanner.Plan(set).Plan!.Worksheets[0].Rows[0].Cells[0];
        Assert.Equal("num", cell.StyleRef);
    }

    [Fact]
    public void Plan_ShouldTreatEmptyStringValue_AsIntentionalBlankCell()
    {
        var set = XlsxInstructionParser.Parse("""
        {
            "version": "1.0",
            "worksheets": [{
                "name": "S",
                "cells": [{"address": "A1", "value": ""}]
            }]
        }
        """);

        var cell = XlsxPlanner.Plan(set).Plan!.Worksheets[0].Cells[0];
        Assert.Equal(string.Empty, cell.Value);
        Assert.Equal(XlsxCellType.String, cell.Type);
    }

    [Fact]
    public void Plan_ShouldWarn_WhenVariablePromotesRowValueToFormula()
    {
        var set = XlsxInstructionParser.Parse("""
        {
            "version": "1.0",
            "variables": { "total": "=SUM(B2:D2)" },
            "worksheets": [{ "name": "S", "rows": [["{{total}}"]] }]
        }
        """);

        var result = XlsxPlanner.Plan(set);
        Assert.True(result.IsValid);
        var warning = Assert.Single(result.Validation.Warnings, d => d.Code == XlsxDiagnosticCode.VariableValueBecameFormula);
        Assert.Equal("worksheets[0].rows[0][0]", warning.Path);

        // Legacy behaviour preserved: the value is promoted to a formula.
        var cell = result.Plan!.Worksheets[0].Rows[0].Cells[0];
        Assert.Equal("=SUM(B2:D2)", cell.Formula);
        Assert.Null(cell.Value);
    }

    [Fact]
    public void Plan_ShouldKeepVariableValueLiteral_WhenColumnTypedString()
    {
        var set = XlsxInstructionParser.Parse("""
        {
            "version": "1.0",
            "variables": { "total": "=SUM(B2:D2)" },
            "worksheets": [{ "name": "S", "columns": [{ "type": "string" }], "rows": [["{{total}}"]] }]
        }
        """);

        var result = XlsxPlanner.Plan(set);
        Assert.True(result.IsValid);
        Assert.Single(result.Validation.Warnings, d => d.Code == XlsxDiagnosticCode.VariableValueBecameFormula);

        var cell = result.Plan!.Worksheets[0].Rows[0].Cells[0];
        Assert.Null(cell.Formula);
        Assert.Equal("=SUM(B2:D2)", cell.Value);
        Assert.Equal(XlsxCellType.String, cell.Type);
    }

    [Fact]
    public void Plan_ShouldNotWarn_WhenFormulaIsAuthored_NotPromoted()
    {
        var set = XlsxInstructionParser.Parse("""
        {
            "version": "1.0",
            "variables": { "range": "B2:D2" },
            "worksheets": [{ "name": "S", "rows": [["=SUM({{range}})"]] }]
        }
        """);

        var result = XlsxPlanner.Plan(set);
        Assert.True(result.IsValid);
        Assert.DoesNotContain(result.Validation.Warnings, d => d.Code == XlsxDiagnosticCode.VariableValueBecameFormula);
        Assert.Equal("=SUM(B2:D2)", result.Plan!.Worksheets[0].Rows[0].Cells[0].Formula);
    }

    [Theory]
    [InlineData("2024-01-15", XlsxCellType.Date)]
    [InlineData("2024-01-15T14:30:00", XlsxCellType.DateTime)]
    [InlineData("2024-01-15T00:00:00", XlsxCellType.DateTime)]
    [InlineData("2024-01-15T00:00:00.000", XlsxCellType.DateTime)]
    [InlineData("01/15/2024", XlsxCellType.String)]
    [InlineData("2024-01-15 10:30", XlsxCellType.String)]
    [InlineData("2024", XlsxCellType.Number)]
    [InlineData("not a date", XlsxCellType.String)]
    public void Plan_ShouldInferDateTypes_UsingExactWriterFormats(string value, XlsxCellType expected)
    {
        // The inferred type must be one the writer can actually produce: only strict ISO
        // date ("yyyy-MM-dd") and ISO datetime ("yyyy-MM-ddTHH:mm:ss[.fff]") values are
        // dates; a midnight datetime stays a datetime; anything else is a string.
        Assert.Equal(expected, XlsxCellTypeParser.ResolveAuto(XlsxCellType.Auto, value));
    }

    // ─── Address / layout ─────────────────────────────────────────

    [Fact]
    public void Plan_ShouldPositionHeadersAndRows_FromStartRow()
    {
        var set = XlsxInstructionParser.Parse("""
        {
            "version": "1.0",
            "worksheets": [{
                "name": "S",
                "startRow": 3,
                "headers": ["A", "B"],
                "rows": [["1", "2"], ["3", "4"]]
            }]
        }
        """);

        var ws = XlsxPlanner.Plan(set).Plan!.Worksheets[0];
        Assert.Equal(3, ws.StartRow);
        Assert.Equal("A3", ws.Headers[0].Reference);
        Assert.Equal("B3", ws.Headers[1].Reference);
        Assert.Equal(4, ws.Rows[0].RowIndex);
        Assert.Equal(5, ws.Rows[1].RowIndex);
        Assert.Equal("B4", ws.Rows[0].Cells[1].Reference);
    }

    [Fact]
    public void Plan_ShouldNormalizeCellAddresses_AndDiscreteCells()
    {
        var set = XlsxInstructionParser.Parse("""
        {
            "version": "1.0",
            "worksheets": [{
                "name": "S",
                "cells": [{"address": "aa10", "value": "x"}]
            }]
        }
        """);

        var cell = XlsxPlanner.Plan(set).Plan!.Worksheets[0].Cells[0];
        Assert.Equal("AA10", cell.Reference);
        Assert.Equal(10, cell.Row);
        Assert.Equal(27, cell.Column);
    }

    [Fact]
    public void Plan_ShouldNormalizeRanges_AndFreezePanes()
    {
        var set = XlsxInstructionParser.Parse("""
        {
            "version": "1.0",
            "worksheets": [{
                "name": "S",
                "headers": ["a", "b"],
                "rows": [["1", "2"]],
                "merges": ["a1:c3"],
                "autoFilter": "a20:d30",
                "tables": [{"name": "T", "range": "A5:B6"}],
                "freezePanes": {"cell": "A2"}
            }]
        }
        """);

        var ws = XlsxPlanner.Plan(set).Plan!.Worksheets[0];
        Assert.Equal("A1:C3", Assert.Single(ws.Merges));
        Assert.Equal("A20:D30", ws.AutoFilterRange);
        Assert.Equal("A5:B6", ws.Tables[0].Range);
        Assert.Equal(2, ws.FreezePanes!.Row);
        Assert.Equal(1, ws.FreezePanes.Column);
    }

    // ─── Limits & duplicate writes ────────────────────────────────

    [Fact]
    public void Plan_ShouldFlagCellTextOverExcelLimit()
    {
        var tooLong = new string('x', 32_768);
        var set = new XlsxInstructionSet
        {
            Version = "1.0",
            Worksheets = [new WorksheetInstruction { Name = "S", Cells = [new CellInstruction { Address = "A1", Value = tooLong }] }]
        };

        var result = XlsxPlanner.Plan(set);
        Assert.False(result.IsValid);
        Assert.Null(result.Plan);
        var error = Assert.Single(result.Validation.Errors, d => d.Code == XlsxDiagnosticCode.TextTooLong);
        Assert.Contains("32,767", error.Message);
    }

    [Fact]
    public void Plan_ShouldFlagFormulaOverExcelLimit()
    {
        var tooLong = "=" + new string('f', 9_000);
        var set = new XlsxInstructionSet
        {
            Version = "1.0",
            Worksheets = [new WorksheetInstruction { Name = "S", Cells = [new CellInstruction { Address = "A1", Formula = tooLong }] }]
        };

        var result = XlsxPlanner.Plan(set);
        Assert.False(result.IsValid);
        Assert.Null(result.Plan);
        var error = Assert.Single(result.Validation.Errors, d => d.Code == XlsxDiagnosticCode.FormulaTooLong);
        Assert.Contains("8,192", error.Message);
    }

    [Fact]
    public void Plan_ShouldWarnOnDuplicateWrite_AndKeepLastWriter()
    {
        var set = XlsxInstructionParser.Parse("""
        {
            "version": "1.0",
            "worksheets": [{
                "name": "S",
                "headers": ["h"],
                "cells": [{"address": "A1", "value": "override"}]
            }]
        }
        """);

        var result = XlsxPlanner.Plan(set);
        Assert.True(result.IsValid);
        var warning = Assert.Single(result.Validation.Warnings, d => d.Code == XlsxDiagnosticCode.DuplicateWrite);
        Assert.Equal("worksheets[0].cells[0]", warning.Path);

        var cell = Assert.Single(result.Plan!.Worksheets[0].Cells);
        Assert.Equal("override", cell.Value);
    }

    [Fact]
    public void Plan_ShouldProduceImmutableData_WithoutTouchingABuilder()
    {
        var set = XlsxInstructionParser.Parse("""
        {
            "version": "1.0",
            "worksheets": [{"name": "S", "rows": [["1"]]}]
        }
        """);

        var plan = XlsxPlanner.Plan(set).Plan!;
        Assert.IsType<XlsxPlan>(plan);
        Assert.IsAssignableFrom<IReadOnlyList<XlsxPlanWorksheet>>(plan.Worksheets);
        Assert.IsAssignableFrom<IReadOnlyList<XlsxPlanCell>>(plan.Worksheets[0].Rows[0].Cells);

        // Records with init-only setters: re-binding a worksheet must produce a new instance.
        var worksheet = plan.Worksheets[0];
        var rebound = worksheet with { Name = "Renamed" };
        Assert.Equal("S", worksheet.Name);
        Assert.Equal("Renamed", rebound.Name);
    }

    // ─── Variable expansion budget (memory-exhaustion guard) ───────

    [Fact]
    public void Plan_ShouldRejectExponentialVariableExpansion_WhenReferenced()
    {
        // A doubling chain doubles at every level: the referenced terminal variable would
        // need ~2^40 characters if fully materialized. The budget aborts expansion with a
        // structured diagnostic instead of ever building the value.
        var variables = new Dictionary<string, string> { ["a0"] = "x" };
        for (var i = 1; i <= 40; i++)
        {
            variables[$"a{i}"] = $"{{{{a{i - 1}}}}}{{{{a{i - 1}}}}}";
        }

        var set = new XlsxInstructionSet
        {
            Version = "1.0",
            Variables = variables,
            Worksheets = [new WorksheetInstruction { Name = "S", Rows = [["{{a40}}"]] }]
        };

        var result = XlsxPlanner.Plan(set, maxVariableLength: 10_000);
        Assert.False(result.IsValid);
        Assert.Null(result.Plan);
        // Several independent variables in the chain cross the budget (a14, a25, a36…);
        // each is reported once with a path-qualified diagnostic.
        var errors = result.Validation.Errors.Where(d => d.Code == XlsxDiagnosticCode.VariableExpansionTooLarge).ToList();
        Assert.NotEmpty(errors);
        Assert.All(errors, e => Assert.Contains("variables.", e.Path));
        Assert.Contains("10,000", errors[0].Message);
    }

    [Fact]
    public void Plan_ShouldNotExpandUnreferencedVariables()
    {
        // A hostile doubling chain that is defined but never referenced by the sheet must
        // stay inert: only the variable the sheet actually uses is resolved.
        var variables = new Dictionary<string, string> { ["a0"] = "x" };
        for (var i = 1; i <= 60; i++)
        {
            variables[$"a{i}"] = $"{{{{a{i - 1}}}}}{{{{a{i - 1}}}}}";
        }
        variables["used"] = "hello";

        var set = new XlsxInstructionSet
        {
            Version = "1.0",
            Variables = variables,
            Worksheets = [new WorksheetInstruction { Name = "S", Rows = [["{{used}}"]] }]
        };

        var result = XlsxPlanner.Plan(set);
        Assert.True(result.IsValid, string.Join("; ", result.Validation.Errors.Select(e => e.Message)));
        Assert.Equal("hello", result.Plan!.Worksheets[0].Rows[0].Cells[0].Value);
        // The unused chain was never expanded into the resolved set.
        Assert.False(result.Plan.Variables.ContainsKey("a60"));
    }

    [Fact]
    public void Plan_ShouldRejectCellResolution_OverLengthBudget()
    {
        // A cell that combines two budget-sized placeholders crosses the budget during
        // cell resolution; the abort is path-qualified to the offending cell.
        var set = new XlsxInstructionSet
        {
            Version = "1.0",
            Variables = new Dictionary<string, string> { ["big"] = new string('x', 6_000) },
            Worksheets = [new WorksheetInstruction { Name = "S", Rows = [["{{big}}{{big}}"]] }]
        };

        var result = XlsxPlanner.Plan(set, maxVariableLength: 10_000);
        Assert.False(result.IsValid);
        Assert.Null(result.Plan);
        var error = Assert.Single(result.Validation.Errors, d => d.Code == XlsxDiagnosticCode.VariableExpansionTooLarge);
        Assert.Equal("worksheets[0].rows[0][0]", error.Path);
    }

    // ─── Row width (XFD) limits ────────────────────────────────────

    [Fact]
    public void Plan_ShouldRejectRowWiderThanExcel_WithPathQualifiedDiagnostic()
    {
        var wideRow = Enumerable.Range(0, 16_385).Select(_ => "x").ToList();
        var set = new XlsxInstructionSet
        {
            Version = "1.0",
            Worksheets = [new WorksheetInstruction { Name = "S", Rows = [wideRow] }]
        };

        var result = XlsxPlanner.Plan(set);
        Assert.False(result.IsValid);
        Assert.Null(result.Plan);
        var error = Assert.Single(result.Validation.Errors, d => d.Code == XlsxDiagnosticCode.RowTooManyCells);
        Assert.Equal("worksheets[0].rows[0]", error.Path);
        Assert.Contains("16384 columns", error.Message);
    }

    [Fact]
    public void Plan_ShouldAcceptRowExactlyAtColumnLimit()
    {
        var maxRow = Enumerable.Range(0, 16_384).Select(_ => "x").ToList();
        var set = new XlsxInstructionSet
        {
            Version = "1.0",
            Worksheets = [new WorksheetInstruction { Name = "S", Rows = [maxRow] }]
        };

        var result = XlsxPlanner.Plan(set);
        Assert.True(result.IsValid);
        Assert.NotNull(result.Plan);
        Assert.Equal(16_384, result.Plan!.Worksheets[0].Rows[0].Cells.Count);
        Assert.Equal("XFD1", result.Plan.Worksheets[0].Rows[0].Cells[^1].Reference);
    }
}
