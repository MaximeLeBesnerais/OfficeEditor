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
    /// <summary>
    /// Default cap on the expanded length of a single variable (and of any resolved cell
    /// value) during planning. Expansion is aborted with a structured diagnostic the
    /// moment an append would cross the budget, so a hostile doubling chain like
    /// <c>a1="{{a0}}{{a0}}"</c> can never materialize gigabytes in memory — even when its
    /// terminal variable is referenced by the sheet. Callers can tune it via the
    /// <see cref="Plan(XlsxInstructionSet, int)"/> overload.
    /// </summary>
    public const int DefaultMaxVariableExpansionLength = 1_048_576;

    public static XlsxPlanResult Plan(XlsxInstructionSet instructions) =>
        Plan(instructions, DefaultMaxVariableExpansionLength);

    /// <summary>
    /// Plans the instruction set with an explicit per-expansion length budget (in
    /// characters). A negative/zero budget is a programmer error and throws; bad input is
    /// always folded into the returned diagnostics instead.
    /// </summary>
    public static XlsxPlanResult Plan(XlsxInstructionSet instructions, int maxVariableLength)
    {
        if (maxVariableLength < 1)
        {
            throw new ArgumentOutOfRangeException(
                nameof(maxVariableLength), maxVariableLength,
                "The maximum variable expansion length must be a positive integer.");
        }

        try
        {
            var validation = XlsxValidationEngine.Validate(instructions);
            if (!validation.IsValid)
            {
                return new XlsxPlanResult(null, validation);
            }

            var diagnostics = new List<XlsxDiagnostic>(validation.Diagnostics);

            // Only the variables the sheet actually references are resolved (plus whatever
            // they transitively reference); defined-but-unused variables are never expanded.
            var resolver = new VariableResolver(instructions.Variables, maxVariableLength);
            var variables = resolver.ResolveAll(CollectReferencedVariables(instructions), diagnostics);
            if (diagnostics.Any(d => d.IsError))
            {
                return new XlsxPlanResult(null, new XlsxValidationResult { Diagnostics = diagnostics });
            }

            var plans = new List<XlsxPlanWorksheet>(instructions.Worksheets.Count);
            for (var s = 0; s < instructions.Worksheets.Count; s++)
            {
                plans.Add(PlanWorksheet(instructions.Worksheets[s], s, variables, diagnostics, maxVariableLength));
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
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            // Bad input must never leak a raw exception: an unexpected planner failure
            // (e.g. a defensive regression in address math) is folded into a structured,
            // path-qualified diagnostic instead, matching the validator's non-throwing
            // contract for document input.
            return new XlsxPlanResult(null, new XlsxValidationResult
            {
                Diagnostics = new[]
                {
                    new XlsxDiagnostic(
                        XlsxDiagnosticCode.PlanningFailed,
                        XlsxDiagnosticSeverity.Error,
                        string.Empty,
                        $"Planning failed: {ex.Message}")
                }
            });
        }
    }

    /// <summary>
    /// Collects the names of every variable the instruction set actually references (as a
    /// '{{name}}' token) in headers, row values and discrete cell values/formulas. Only
    /// these — plus whatever they transitively reference — are resolved; variables the
    /// sheet never uses stay inert, so an exponential definition chain is never expanded
    /// merely because it was declared.
    /// </summary>
    private static IEnumerable<string> CollectReferencedVariables(XlsxInstructionSet set)
    {
        var names = new HashSet<string>(StringComparer.Ordinal);
        if (set.Worksheets is not { Count: > 0 })
        {
            return names;
        }

        foreach (var ws in set.Worksheets)
        {
            if (ws is null)
            {
                continue;
            }

            if (ws.Headers is { Count: > 0 })
            {
                foreach (var header in ws.Headers)
                {
                    if (header is not null)
                    {
                        CollectPlaceholders(header, names);
                    }
                }
            }

            if (ws.Rows is { Count: > 0 })
            {
                foreach (var row in ws.Rows)
                {
                    if (row is null)
                    {
                        continue;
                    }

                    foreach (var value in row)
                    {
                        if (value is not null)
                        {
                            CollectPlaceholders(value, names);
                        }
                    }
                }
            }

            if (ws.Cells is { Count: > 0 })
            {
                foreach (var cell in ws.Cells)
                {
                    if (cell is null)
                    {
                        continue;
                    }

                    if (cell.Value is not null)
                    {
                        CollectPlaceholders(cell.Value, names);
                    }

                    if (cell.Formula is not null)
                    {
                        CollectPlaceholders(cell.Formula, names);
                    }
                }
            }
        }

        return names;
    }

    private static void CollectPlaceholders(string text, HashSet<string> names)
    {
        if (!text.Contains("{{", StringComparison.Ordinal))
        {
            return;
        }

        var pos = 0;
        while (pos < text.Length)
        {
            var open = text.IndexOf("{{", pos, StringComparison.Ordinal);
            if (open < 0)
            {
                return;
            }

            var close = text.IndexOf("}}", open + 2, StringComparison.Ordinal);
            if (close < 0)
            {
                return;
            }

            var inner = text.Substring(open + 2, close - open - 2).Trim();
            if (inner.Length > 0)
            {
                names.Add(inner);
            }

            pos = close + 2;
        }
    }

    // ─── Variables ────────────────────────────────────────────────

    private sealed class VariableResolver
    {
        private readonly Dictionary<string, string>? _raw;
        private readonly Dictionary<string, string> _resolved = new(StringComparer.Ordinal);
        private readonly HashSet<string> _overBudget = new(StringComparer.Ordinal);
        private readonly HashSet<string> _inProgress = new(StringComparer.Ordinal);
        private readonly int _maxLength;

        public VariableResolver(Dictionary<string, string>? variables, int maxLength)
        {
            _raw = variables;
            _maxLength = maxLength;
        }

        public IReadOnlyDictionary<string, string> ResolveAll(
            IEnumerable<string> referencedVariables, List<XlsxDiagnostic> diagnostics)
        {
            if (_raw is null)
            {
                return new Dictionary<string, string>();
            }

            // Only variables the sheet actually references (directly or transitively) are
            // resolved. Unreferenced definitions are never expanded, so a hostile doubling
            // chain stays inert unless its terminal variable is used by the sheet.
            foreach (var key in referencedVariables)
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

            // An over-budget variable is never materialized: it resolves to an empty value
            // (the budget diagnostic is already emitted) so a referencing variable keeps
            // its token instead of concatenating a truncated giant prefix.
            _resolved[key] = _overBudget.Contains(key) ? string.Empty : result;
        }

        private string Expand(string owner, string input, List<XlsxDiagnostic> diagnostics)
        {
            if (!input.Contains("{{", StringComparison.Ordinal))
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
                    TryAppend(sb, owner, input, pos, input.Length - pos, diagnostics);
                    break;
                }

                var close = input.IndexOf("}}", open + 2, StringComparison.Ordinal);
                if (close < 0)
                {
                    TryAppend(sb, owner, input, pos, input.Length - pos, diagnostics);
                    diagnostics.Add(new XlsxDiagnostic(
                        XlsxDiagnosticCode.UnresolvedVariable,
                        XlsxDiagnosticSeverity.Error,
                        $"variables.{owner}",
                        $"Variable '{owner}' contains an unclosed '{{{{' placeholder; every placeholder must be '{{{{name}}}}'."));
                    break;
                }

                if (!TryAppend(sb, owner, input, pos, open - pos, diagnostics))
                {
                    break;
                }

                var inner = input.Substring(open + 2, close - open - 2).Trim();
                if (_raw!.ContainsKey(inner))
                {
                    ResolveVariable(inner, diagnostics);
                    if (_resolved.TryGetValue(inner, out var resolved))
                    {
                        if (_overBudget.Contains(inner))
                        {
                            // The referenced variable already exceeded the budget on its own
                            // expansion (reported there); keep its token so no fabricated
                            // content is planned and no second diagnostic is emitted.
                            sb.Append("{{").Append(inner).Append("}}");
                        }
                        else if (!TryAppend(sb, owner, resolved, 0, resolved.Length, diagnostics))
                        {
                            break;
                        }
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

        /// <summary>
        /// Appends a slice of <paramref name="value"/> to the buffer, aborting the
        /// expansion (with a structured diagnostic) the moment it would cross the budget.
        /// The buffer is therefore bounded by the budget plus one appended chunk, so even a
        /// pathological doubling chain cannot allocate more than a constant multiple of the
        /// budget.
        /// </summary>
        private bool TryAppend(StringBuilder sb, string owner, string value, int start, int length, List<XlsxDiagnostic> diagnostics)
        {
            if (sb.Length + length <= _maxLength)
            {
                sb.Append(value, start, length);
                return true;
            }

            if (_overBudget.Add(owner))
            {
                diagnostics.Add(new XlsxDiagnostic(
                    XlsxDiagnosticCode.VariableExpansionTooLarge,
                    XlsxDiagnosticSeverity.Error,
                    $"variables.{owner}",
                    $"Variable '{owner}' expands beyond {_maxLength:N0} characters; expansion " +
                    "is aborted instead of materializing the full value."));
            }

            return false;
        }
    }

    // ─── Worksheets ───────────────────────────────────────────────

    private static XlsxPlanWorksheet PlanWorksheet(
        WorksheetInstruction ws, int sheetIndex, IReadOnlyDictionary<string, string> variables,
        List<XlsxDiagnostic> diagnostics, int maxLength)
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
                    diagnostics, maxLength);
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
                    var raw = row[j]!;
                    var resolved = ResolveValue(
                        variables, raw, path, $"cell {reference} of sheet '{sheetName}'", diagnostics, maxLength);
                    var column = ColumnAt(columns, j);
                    var declaredType = column?.Type ?? XlsxCellType.Auto;
                    var authoredFormula = raw.StartsWith('=');

                    string? formula = null;
                    string? value = resolved;
                    var type = XlsxCellType.Auto;

                    if (authoredFormula)
                    {
                        // A row value authored with a leading '=' is an explicit formula;
                        // variable substitution happens inside it (legacy behaviour).
                        formula = resolved;
                        value = null;
                        type = XlsxCellType.Auto;
                    }
                    else if (resolved.StartsWith('='))
                    {
                        // Variable substitution produced a formula-like value.
                        if (declaredType == XlsxCellType.String)
                        {
                            // Explicit 'string' typing keeps the value literal; the promotion
                            // is suppressed so no formula is written.
                            type = XlsxCellType.String;
                            diagnostics.Add(new XlsxDiagnostic(
                                XlsxDiagnosticCode.VariableValueBecameFormula,
                                XlsxDiagnosticSeverity.Warning,
                                path,
                                $"Variable substitution made the value of cell {reference} in " +
                                $"sheet '{sheetName}' start with '=' ('{resolved}'), but the " +
                                $"column is typed 'string', so it is written literally. Remove " +
                                $"the column's 'string' type to treat it as a formula."));
                        }
                        else
                        {
                            // Legacy behaviour: the value is promoted to a formula, but a
                            // warning makes the implicit conversion explicit.
                            formula = resolved;
                            value = null;
                            type = XlsxCellType.Auto;
                            diagnostics.Add(new XlsxDiagnostic(
                                XlsxDiagnosticCode.VariableValueBecameFormula,
                                XlsxDiagnosticSeverity.Warning,
                                path,
                                $"Variable substitution turned the value of cell {reference} in " +
                                $"sheet '{sheetName}' into a formula ('{resolved}'). Type the " +
                                $"column as 'string' to keep it literal."));
                        }
                    }
                    else
                    {
                        type = XlsxCellTypeParser.ResolveAuto(declaredType, resolved);
                    }

                    var cell = new XlsxPlanCell
                    {
                        Reference = reference,
                        Row = rowIndex,
                        Column = j + 1,
                        Value = value,
                        Formula = formula,
                        Type = type,
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
                    ? ResolveValue(variables, instruction.Formula!, path, $"formula of cell {reference} in sheet '{sheetName}'", diagnostics, maxLength)
                    : null;
                var resolvedValue = !isFormula && instruction.Value is not null
                    ? ResolveValue(variables, instruction.Value, path, $"cell {reference} of sheet '{sheetName}'", diagnostics, maxLength)
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
        List<XlsxDiagnostic> diagnostics, int maxLength)
    {
        if (!raw.Contains("{{", StringComparison.Ordinal))
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
                TryAppend(sb, raw, pos, raw.Length - pos, maxLength, path, context, diagnostics);
                break;
            }

            var close = raw.IndexOf("}}", open + 2, StringComparison.Ordinal);
            if (close < 0)
            {
                TryAppend(sb, raw, pos, raw.Length - pos, maxLength, path, context, diagnostics);
                diagnostics.Add(new XlsxDiagnostic(
                    XlsxDiagnosticCode.UnresolvedVariable,
                    XlsxDiagnosticSeverity.Error,
                    path,
                    $"Unresolved variable placeholder in {context}. Provide a value for it in the " +
                    "instruction set's 'variables' block."));
                break;
            }

            if (!TryAppend(sb, raw, pos, open - pos, maxLength, path, context, diagnostics))
            {
                break;
            }

            var inner = raw.Substring(open + 2, close - open - 2).Trim();
            if (variables.TryGetValue(inner, out var resolved))
            {
                if (!TryAppend(sb, resolved, 0, resolved.Length, maxLength, path, context, diagnostics))
                {
                    break;
                }
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

    /// <summary>
    /// Appends a slice of <paramref name="value"/> to the cell-resolution buffer, aborting
    /// with a structured diagnostic the moment the combined result would cross the length
    /// budget — so a resolved value is never materialized past the cap even when many
    /// large placeholders combine in one cell.
    /// </summary>
    private static bool TryAppend(
        StringBuilder sb, string value, int start, int length, int maxLength,
        string path, string context, List<XlsxDiagnostic> diagnostics)
    {
        if (sb.Length + length <= maxLength)
        {
            sb.Append(value, start, length);
            return true;
        }

        diagnostics.Add(new XlsxDiagnostic(
            XlsxDiagnosticCode.VariableExpansionTooLarge,
            XlsxDiagnosticSeverity.Error,
            path,
            $"Resolved value in {context} expands beyond {maxLength:N0} characters; expansion " +
            "is aborted instead of materializing the full value."));
        return false;
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
