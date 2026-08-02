using System.Text;
using XlsxEditor.Core.Builders;

namespace XlsxEditor.Core.Instructions;

/// <summary>
/// Builds an immutable <see cref="XlsxPlan"/> from a validated instruction set. The
/// planner is pure: it performs no workbook mutation and depends on no builder
/// implementation. It resolves variables deterministically (with cycle detection),
/// expands headers/rows/cells into concrete normalized addresses, resolves cell types,
/// applies column defaults, and collects pre-execution diagnostics (unresolved
/// placeholders, out-of-limit text/formulas, duplicate writes) — all before anything is
/// written.
/// </summary>
public static class XlsxPlanner
{
    public static XlsxPlanResult Plan(XlsxInstructionSet instructions)
    {
        var validation = XlsxValidationEngine.Validate(instructions);
        if (!validation.IsValid)
        {
            return new XlsxPlanResult(null, validation);
        }

        var diagnostics = new List<XlsxDiagnostic>(validation.Diagnostics);

        var resolver = new VariableResolver(instructions.Variables);
        var variables = resolver.ResolveAll(diagnostics);
        if (diagnostics.Any(d => d.IsError))
        {
            return new XlsxPlanResult(null, new XlsxValidationResult { Diagnostics = diagnostics });
        }

        var plans = new List<XlsxPlanWorksheet>(instructions.Worksheets.Count);
        for (var s = 0; s < instructions.Worksheets.Count; s++)
        {
            plans.Add(PlanWorksheet(instructions.Worksheets[s], s, variables, diagnostics));
        }

        var plan = new XlsxPlan
        {
            Version = instructions.Version,
            Description = instructions.Description,
            Metadata = instructions.Metadata,
            Styles = instructions.Styles is { Count: > 0 } ? instructions.Styles : Array.Empty<NamedStyle>(),
            Worksheets = plans,
            Variables = variables
        };

        // A plan is only produced when it is safe to execute: any error (validation or
        // resolution, e.g. an over-limit cell) yields no plan and the full diagnostic set.
        if (diagnostics.Any(d => d.IsError))
        {
            return new XlsxPlanResult(null, new XlsxValidationResult { Diagnostics = diagnostics });
        }

        return new XlsxPlanResult(plan, new XlsxValidationResult { Diagnostics = diagnostics });
    }

    // ─── Variables ────────────────────────────────────────────────

    private sealed class VariableResolver
    {
        private readonly IReadOnlyDictionary<string, string>? _raw;
        private readonly Dictionary<string, string> _resolved = new(StringComparer.Ordinal);
        private readonly HashSet<string> _inProgress = new(StringComparer.Ordinal);

        public VariableResolver(Dictionary<string, string>? variables) => _raw = variables;

        public IReadOnlyDictionary<string, string> ResolveAll(List<XlsxDiagnostic> diagnostics)
        {
            if (_raw is null)
            {
                return new Dictionary<string, string>();
            }

            // Materialize keys once so iteration order cannot change mid-resolution.
            foreach (var key in _raw.Keys.ToList())
            {
                ResolveVariable(key, diagnostics);
            }

            return _resolved;
        }

        private void ResolveVariable(string key, List<XlsxDiagnostic> diagnostics)
        {
            if (_resolved.ContainsKey(key))
            {
                return;
            }

            if (_inProgress.Contains(key))
            {
                var chain = string.Join(" -> ", _inProgress.Append(key));
                diagnostics.Add(new XlsxDiagnostic(
                    XlsxDiagnosticCode.VariableCycle,
                    XlsxDiagnosticSeverity.Error,
                    $"variables.{key}",
                    $"Variable cycle detected: {chain}. Variables must not reference each other in a loop."));
                return;
            }

            _inProgress.Add(key);
            var value = _raw![key] ?? string.Empty;
            var result = Expand(key, value, diagnostics);
            _inProgress.Remove(key);
            _resolved[key] = result;
        }

        private string Expand(string owner, string input, List<XlsxDiagnostic> diagnostics)
        {
            if (!input.Contains("{{"))
            {
                return input;
            }

            var sb = new StringBuilder();
            var pos = 0;
            while (pos < input.Length)
            {
                var open = input.IndexOf("{{", pos, StringComparison.Ordinal);
                if (open < 0)
                {
                    sb.Append(input, pos, input.Length - pos);
                    break;
                }

                var close = input.IndexOf("}}", open + 2, StringComparison.Ordinal);
                if (close < 0)
                {
                    sb.Append(input, pos, input.Length - pos);
                    diagnostics.Add(new XlsxDiagnostic(
                        XlsxDiagnosticCode.UnresolvedVariable,
                        XlsxDiagnosticSeverity.Error,
                        $"variables.{owner}",
                        $"Variable '{owner}' contains an unclosed '{{{{' placeholder; every placeholder must be '{{{{name}}}}'."));
                    break;
                }

                sb.Append(input, pos, open - pos);
                var inner = input.Substring(open + 2, close - open - 2).Trim();
                if (_raw!.ContainsKey(inner))
                {
                    ResolveVariable(inner, diagnostics);
                    if (_resolved.TryGetValue(inner, out var resolved))
                    {
                        sb.Append(resolved);
                    }
                    else
                    {
                        // A cycle broke resolution; keep the token verbatim (the cycle is
                        // already reported) so no fabricated content is planned.
                        sb.Append("{{").Append(inner).Append("}}");
                    }
                }
                else
                {
                    diagnostics.Add(new XlsxDiagnostic(
                        XlsxDiagnosticCode.UndefinedVariableReference,
                        XlsxDiagnosticSeverity.Error,
                        $"variables.{owner}",
                        $"Variable '{inner}' is referenced in variable '{owner}' but is not defined " +
                        "in the 'variables' block."));
                    sb.Append("{{").Append(inner).Append("}}");
                }

                pos = close + 2;
            }

            return sb.ToString();
        }
    }

    // ─── Worksheets ───────────────────────────────────────────────

    private static XlsxPlanWorksheet PlanWorksheet(
        WorksheetInstruction ws, int sheetIndex, IReadOnlyDictionary<string, string> variables,
        List<XlsxDiagnostic> diagnostics)
    {
        var sheetName = ws.Name;
        var startRow = ws.StartRow ?? 1;
        var columns = PlanColumns(ws, sheetIndex, diagnostics);

        var headers = new List<XlsxPlanCell>();
        var planRows = new List<XlsxPlanRow>();
        var cells = new List<XlsxPlanCell>();
        var writes = new Dictionary<(int Row, int Col), List<string>>();

        if (ws.Headers is { Count: > 0 })
        {
            for (var i = 0; i < ws.Headers.Count; i++)
            {
                var path = $"worksheets[{sheetIndex}].headers[{i}]";
                var reference = WorksheetBuilder.GetCellReference(i, startRow);
                var resolved = ResolveValue(
                    variables, ws.Headers[i]!, path, $"header column {i + 1} of sheet '{sheetName}'",
                    diagnostics);
                var header = new XlsxPlanCell
                {
                    Reference = reference,
                    Row = startRow,
                    Column = i + 1,
                    Value = resolved,
                    Type = XlsxCellType.String,
                    StyleRef = ws.HeaderStyle,
                    Source = path
                };
                CheckCellLimits(header, sheetName, diagnostics);
                RecordWrite(writes, (header.Row, header.Column), path, sheetName, diagnostics);
                headers.Add(header);
            }
        }

        var headerOffset = ws.Headers is { Count: > 0 } ? 1 : 0;
        if (ws.Rows is { Count: > 0 })
        {
            for (var i = 0; i < ws.Rows.Count; i++)
            {
                var row = ws.Rows[i]!;
                var rowIndex = startRow + headerOffset + i;
                var rowCells = new List<XlsxPlanCell>(row.Count);
                for (var j = 0; j < row.Count; j++)
                {
                    var path = $"worksheets[{sheetIndex}].rows[{i}][{j}]";
                    var reference = WorksheetBuilder.GetCellReference(j, rowIndex);
                    var resolved = ResolveValue(
                        variables, row[j]!, path, $"cell {reference} of sheet '{sheetName}'", diagnostics);
                    var column = ColumnAt(columns, j);
                    var isFormula = resolved.StartsWith('=');
                    var cell = new XlsxPlanCell
                    {
                        Reference = reference,
                        Row = rowIndex,
                        Column = j + 1,
                        Value = isFormula ? null : resolved,
                        Formula = isFormula ? resolved : null,
                        Type = isFormula
                            ? XlsxCellType.Auto
                            : XlsxCellTypeParser.ResolveAuto(column?.Type ?? XlsxCellType.Auto, resolved),
                        StyleRef = column?.StyleRef,
                        Source = path
                    };
                    CheckCellLimits(cell, sheetName, diagnostics);
                    RecordWrite(writes, (cell.Row, cell.Column), path, sheetName, diagnostics);
                    rowCells.Add(cell);
                }

                planRows.Add(new XlsxPlanRow { RowIndex = rowIndex, Cells = rowCells });
            }
        }

        if (ws.Cells is { Count: > 0 })
        {
            for (var i = 0; i < ws.Cells.Count; i++)
            {
                var instruction = ws.Cells[i]!;
                var path = $"worksheets[{sheetIndex}].cells[{i}]";
                var reference = WorksheetBuilder.NormalizeCellReference(instruction.Address);
                var row = WorksheetBuilder.GetRowIndex(reference);
                var col = WorksheetBuilder.GetColumnIndex(reference);

                var isFormula = instruction.Formula is not null;
                var resolvedFormula = isFormula
                    ? ResolveValue(variables, instruction.Formula!, path, $"formula of cell {reference} in sheet '{sheetName}'", diagnostics)
                    : null;
                var resolvedValue = !isFormula && instruction.Value is not null
                    ? ResolveValue(variables, instruction.Value, path, $"cell {reference} of sheet '{sheetName}'", diagnostics)
                    : null;

                var declaredType = instruction.Type is not null && XlsxCellTypeParser.TryParse(instruction.Type, out var parsed)
                    ? parsed
                    : XlsxCellType.Auto;
                var column = ColumnAt(columns, col - 1);
                // Formulas have no known concrete type until Excel evaluates them; a
                // declared type is honored, otherwise the type stays Auto.
                var effectiveType = resolvedFormula is not null
                    ? declaredType != XlsxCellType.Auto ? declaredType : XlsxCellType.Auto
                    : XlsxCellTypeParser.ResolveAuto(
                        declaredType != XlsxCellType.Auto ? declaredType : column?.Type ?? XlsxCellType.Auto,
                        resolvedValue);

                var cell = new XlsxPlanCell
                {
                    Reference = reference,
                    Row = row,
                    Column = col,
                    Value = resolvedValue,
                    Formula = resolvedFormula,
                    Type = effectiveType,
                    StyleRef = instruction.Style,
                    NumberFormat = instruction.NumberFormat,
                    Source = path
                };
                CheckCellLimits(cell, sheetName, diagnostics);
                RecordWrite(writes, (cell.Row, cell.Column), path, sheetName, diagnostics);
                cells.Add(cell);
            }
        }

        return new XlsxPlanWorksheet
        {
            Name = sheetName,
            StartRow = startRow,
            HeaderStyle = ws.HeaderStyle,
            Columns = columns,
            Headers = headers,
            Rows = planRows,
            Cells = cells,
            RowHeights = PlanRowHeights(ws),
            Merges = PlanMerges(ws),
            FreezePanes = PlanFreezePanes(ws),
            AutoFilterRange = PlanAutoFilter(ws),
            Tables = PlanTables(ws)
        };
    }

    private static IReadOnlyList<XlsxPlanColumn> PlanColumns(
        WorksheetInstruction ws, int sheetIndex, List<XlsxDiagnostic> diagnostics)
    {
        if (ws.Columns is null)
        {
            return Array.Empty<XlsxPlanColumn>();
        }

        var columns = new List<XlsxPlanColumn>(ws.Columns.Count);
        for (var i = 0; i < ws.Columns.Count; i++)
        {
            var instruction = ws.Columns[i]!;
            columns.Add(new XlsxPlanColumn
            {
                Index = i,
                ColumnLetter = WorksheetBuilder.GetColumnName(i),
                Name = instruction.Name,
                Width = instruction.Width,
                StyleRef = instruction.Style,
                Type = instruction.Type is not null && XlsxCellTypeParser.TryParse(instruction.Type, out var type)
                    ? type
                    : XlsxCellType.Auto
            });
        }

        return columns;
    }

    private static XlsxPlanColumn? ColumnAt(IReadOnlyList<XlsxPlanColumn> columns, int columnIndex) =>
        columnIndex >= 0 && columnIndex < columns.Count ? columns[columnIndex] : null;

    private static IReadOnlyDictionary<int, double> PlanRowHeights(WorksheetInstruction ws)
    {
        if (ws.RowHeights is null)
        {
            return new Dictionary<int, double>();
        }

        // Last write wins; duplicate rows were already reported as warnings by validation.
        var heights = new Dictionary<int, double>();
        foreach (var entry in ws.RowHeights)
        {
            if (entry is { Row: { } row, Height: { } height })
            {
                heights[row] = height;
            }
        }

        return heights;
    }

    private static IReadOnlyList<string> PlanMerges(WorksheetInstruction ws)
    {
        if (ws.Merges is null)
        {
            return Array.Empty<string>();
        }

        var merges = new List<string>(ws.Merges.Count);
        foreach (var range in ws.Merges)
        {
            var (start, end) = XlsxRangeUtilities.NormalizeRange(range!);
            merges.Add($"{start}:{end}");
        }

        return merges;
    }

    private static XlsxPlanFreezePanes? PlanFreezePanes(WorksheetInstruction ws)
    {
        var freeze = ws.FreezePanes;
        if (freeze is null)
        {
            return null;
        }

        if (freeze.Cell is not null)
        {
            var reference = WorksheetBuilder.NormalizeCellReference(freeze.Cell);
            return new XlsxPlanFreezePanes
            {
                Row = WorksheetBuilder.GetRowIndex(reference),
                Column = WorksheetBuilder.GetColumnIndex(reference)
            };
        }

        return new XlsxPlanFreezePanes
        {
            Row = freeze.Row ?? 1,
            Column = freeze.Column ?? 1
        };
    }

    private static string? PlanAutoFilter(WorksheetInstruction ws)
    {
        if (ws.AutoFilter is null)
        {
            return null;
        }

        var (start, end) = XlsxRangeUtilities.NormalizeRange(ws.AutoFilter);
        return $"{start}:{end}";
    }

    private static IReadOnlyList<XlsxPlanTable> PlanTables(WorksheetInstruction ws)
    {
        if (ws.Tables is null)
        {
            return Array.Empty<XlsxPlanTable>();
        }

        var tables = new List<XlsxPlanTable>(ws.Tables.Count);
        foreach (var table in ws.Tables)
        {
            var (start, end) = XlsxRangeUtilities.NormalizeRange(table!.Range);
            tables.Add(new XlsxPlanTable
            {
                Name = table.Name,
                Range = $"{start}:{end}"
            });
        }

        return tables;
    }

    // ─── Cell helpers ─────────────────────────────────────────────

    private static string ResolveValue(
        IReadOnlyDictionary<string, string> variables, string raw, string path, string context,
        List<XlsxDiagnostic> diagnostics)
    {
        if (!raw.Contains("{{"))
        {
            return raw;
        }

        var sb = new StringBuilder();
        var pos = 0;
        while (pos < raw.Length)
        {
            var open = raw.IndexOf("{{", pos, StringComparison.Ordinal);
            if (open < 0)
            {
                sb.Append(raw, pos, raw.Length - pos);
                break;
            }

            var close = raw.IndexOf("}}", open + 2, StringComparison.Ordinal);
            if (close < 0)
            {
                sb.Append(raw, pos, raw.Length - pos);
                diagnostics.Add(new XlsxDiagnostic(
                    XlsxDiagnosticCode.UnresolvedVariable,
                    XlsxDiagnosticSeverity.Error,
                    path,
                    $"Unresolved variable placeholder in {context}. Provide a value for it in the " +
                    "instruction set's 'variables' block."));
                break;
            }

            sb.Append(raw, pos, open - pos);
            var inner = raw.Substring(open + 2, close - open - 2).Trim();
            if (variables.TryGetValue(inner, out var resolved))
            {
                sb.Append(resolved);
            }
            else
            {
                diagnostics.Add(new XlsxDiagnostic(
                    XlsxDiagnosticCode.UnresolvedVariable,
                    XlsxDiagnosticSeverity.Error,
                    path,
                    $"Unresolved variable {{{{{inner}}}}} in {context}. Provide a value for it in the " +
                    "instruction set's 'variables' block."));
                sb.Append("{{").Append(inner).Append("}}");
            }

            pos = close + 2;
        }

        return sb.ToString();
    }

    private static void CheckCellLimits(XlsxPlanCell cell, string sheetName, List<XlsxDiagnostic> diagnostics)
    {
        if (cell.Formula is { Length: > XlsxRangeUtilities.MaxFormulaLength } formula)
        {
            diagnostics.Add(new XlsxDiagnostic(
                XlsxDiagnosticCode.FormulaTooLong,
                XlsxDiagnosticSeverity.Error,
                cell.Source,
                $"Formula in cell '{cell.Reference}' of sheet '{sheetName}' is {formula.Length} characters, " +
                $"exceeding Excel's {XlsxRangeUtilities.MaxFormulaLength:N0}-character formula limit."));
        }
        else if (cell.Value is { Length: > XlsxRangeUtilities.MaxCellTextLength } value)
        {
            diagnostics.Add(new XlsxDiagnostic(
                XlsxDiagnosticCode.TextTooLong,
                XlsxDiagnosticSeverity.Error,
                cell.Source,
                $"Cell '{cell.Reference}' of sheet '{sheetName}' contains {value.Length} characters, " +
                $"exceeding Excel's {XlsxRangeUtilities.MaxCellTextLength:N0}-character cell text limit."));
        }
    }

    private static void RecordWrite(
        Dictionary<(int Row, int Col), List<string>> writes, (int Row, int Col) key, string path,
        string sheetName, List<XlsxDiagnostic> diagnostics)
    {
        if (writes.TryGetValue(key, out var existing))
        {
            var reference = WorksheetBuilder.GetCellReference(key.Col - 1, key.Row);
            diagnostics.Add(new XlsxDiagnostic(
                XlsxDiagnosticCode.DuplicateWrite,
                XlsxDiagnosticSeverity.Warning,
                path,
                $"Cell {reference} of sheet '{sheetName}' is written more than once (first written at " +
                $"{existing[0]}); the last write wins."));
            existing.Add(path);
        }
        else
        {
            writes[key] = new List<string> { path };
        }
    }
}
