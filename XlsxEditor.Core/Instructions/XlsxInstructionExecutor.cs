using XlsxEditor.Core.Builders;
using XlsxEditor.Core.Exceptions;

namespace XlsxEditor.Core.Instructions;

/// <summary>
/// Executes an XlsxInstructionSet against a WorkbookBuilder, producing the workbook.
/// </summary>
public static class XlsxInstructionExecutor
{
    /// <summary>
    /// Applies all instructions from the set to the given workbook builder.
    /// Variables in cell values ({{…}}) are resolved at this point.
    /// </summary>
    public static void Execute(XlsxInstructionSet instructions, IWorkbookBuilder builder)
    {
        var variables = instructions.Variables ?? new Dictionary<string, string>();

        foreach (var wsInstruction in instructions.Worksheets)
        {
            var sheet = builder.AddWorksheet(wsInstruction.Name);

            // 1. Headers
            if (wsInstruction.Headers is { Count: > 0 })
            {
                sheet.AddHeaderRow(ResolveList(wsInstruction.Headers, variables));
            }

            // 2. Row data
            if (wsInstruction.Rows is { Count: > 0 })
            {
                int headerOffset = wsInstruction.Headers is { Count: > 0 } ? 1 : 0;
                for (int i = 0; i < wsInstruction.Rows.Count; i++)
                {
                    var rowIndex = i + headerOffset + 1;
                    var values = ResolveList(wsInstruction.Rows[i], variables);

                    for (int col = 0; col < values.Count; col++)
                    {
                        var cellRef = WorksheetBuilder.GetCellReference(col, rowIndex);
                        var val = values[col];
                        if (val.StartsWith("="))
                        {
                            sheet.AddCell(cellRef, val, true);
                        }
                        else
                        {
                            sheet.AddCell(cellRef, val);
                        }
                    }
                }
            }

            // 3. Discrete cells
            if (wsInstruction.Cells is { Count: > 0 })
            {
                foreach (var cell in wsInstruction.Cells)
                {
                    ApplyCellInstruction(sheet, cell, variables, wsInstruction.Name);
                }
            }
        }
    }

    private static List<string> ResolveList(List<string> values, Dictionary<string, string> variables)
    {
        return values.Select(v => ResolveVariables(v, variables)).ToList();
    }

    internal static string ResolveVariables(string input, Dictionary<string, string> variables)
    {
        if (string.IsNullOrEmpty(input) || !input.Contains("{{"))
        {
            return input;
        }

        var result = input;
        foreach (var (key, value) in variables)
        {
            var placeholder = $"{{{{{key}}}}}";
            result = result.Replace(placeholder, value);
        }
        return result;
    }

    private static void ApplyCellInstruction(
        IWorksheetBuilder sheet, CellInstruction cell, Dictionary<string, string> variables, string sheetName)
    {
        // Defense in depth: instruction sets built in code (bypassing the parser's
        // validation) must not silently drop 'type'/'numberFormat' either.
        XlsxInstructionValidator.RejectUnsupportedCellFields(cell, sheetName);

        if (!string.IsNullOrEmpty(cell.Formula))
        {
            var formula = ResolveVariables(cell.Formula, variables);
            sheet.AddCell(cell.Address, formula, true);
            return;
        }

        if (!string.IsNullOrEmpty(cell.Value))
        {
            var resolved = ResolveVariables(cell.Value, variables);

            if (!string.IsNullOrEmpty(cell.Style))
            {
                sheet.AddCell(cell.Address, resolved, cell.Style);
            }
            else
            {
                sheet.AddCell(cell.Address, resolved);
            }
        }
    }
}
