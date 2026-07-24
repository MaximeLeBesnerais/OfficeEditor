using System.Text.RegularExpressions;
using XlsxEditor.Core.Builders;
using XlsxEditor.Core.Exceptions;

namespace XlsxEditor.Core.Instructions;

/// <summary>
/// Loud validator — rejects malformed instruction sets with actionable messages.
/// </summary>
public static class XlsxInstructionValidator
{
    private static readonly Regex CellAddressRegex =
        new(@"^[A-Za-z]{1,3}[1-9][0-9]*$", RegexOptions.Compiled);

    private static readonly Regex InvalidSheetNameChars =
        new(@"[:\\\/\?\*\[\]]", RegexOptions.Compiled);

    public static void Validate(XlsxInstructionSet instructions)
    {
        if (instructions == null)
        {
            throw new XlsxException("Instruction set must not be null.");
        }

        if (string.IsNullOrWhiteSpace(instructions.Version))
        {
            throw new XlsxException("Instruction set must declare a 'version'.");
        }

        if (instructions.Version != "1.0")
        {
            throw new XlsxException(
                $"Unsupported instruction version '{instructions.Version}'. Only version '1.0' is supported.");
        }

        if (instructions.Worksheets.Count == 0)
        {
            throw new XlsxException(
                "Instruction set must contain at least one worksheet in 'worksheets'.");
        }

        // Excel worksheet names are case-insensitive: "Sales" and "SALES" collide.
        var seenSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var ws in instructions.Worksheets)
        {
            ValidateWorksheet(ws, seenSheetNames);
        }

        // Variables block: warn on empty but don't reject
        if (instructions.Variables != null && instructions.Variables.Count == 0)
        {
            throw new XlsxException("'variables' block is present but empty. Either remove it or provide values.");
        }
    }

    private static void ValidateWorksheet(WorksheetInstruction ws, HashSet<string> seenNames)
    {
        if (string.IsNullOrWhiteSpace(ws.Name))
        {
            throw new XlsxException("Every worksheet must have a non-empty 'name'.");
        }

        if (ws.Name.Length > 31)
        {
            throw new XlsxException(
                $"Worksheet name '{ws.Name}' exceeds Excel's 31-character limit.");
        }

        if (InvalidSheetNameChars.IsMatch(ws.Name))
        {
            throw new XlsxException(
                $"Worksheet name '{ws.Name}' contains illegal characters (: \\ / ? * [ ]).");
        }

        if (!seenNames.Add(ws.Name))
        {
            throw new XlsxException(
                $"Duplicate worksheet name '{ws.Name}'. Worksheet names must be unique (case-insensitive).");
        }

        var hasHeaders = ws.Headers is { Count: > 0 };
        var hasRows = ws.Rows is { Count: > 0 };
        var hasCells = ws.Cells is { Count: > 0 };

        if (!hasHeaders && !hasRows && !hasCells)
        {
            throw new XlsxException(
                $"Worksheet '{ws.Name}' must have at least one of 'headers', 'rows', or 'cells'.");
        }

        if (ws.Cells != null)
        {
            foreach (var cell in ws.Cells)
            {
                ValidateCellInstruction(cell, ws.Name);
            }
        }
    }

    private static void ValidateCellInstruction(CellInstruction cell, string sheetName)
    {
        if (string.IsNullOrWhiteSpace(cell.Address))
        {
            throw new XlsxException($"Cell instruction in sheet '{sheetName}' is missing 'address'.");
        }

        if (!CellAddressRegex.IsMatch(cell.Address))
        {
            throw new XlsxException(
                $"Invalid cell address '{cell.Address}' in sheet '{sheetName}'. " +
                "Expected a valid Excel reference like 'A1', 'AA10', etc.");
        }

        var hasValue = !string.IsNullOrEmpty(cell.Value);
        var hasFormula = !string.IsNullOrEmpty(cell.Formula);

        if (hasValue && hasFormula)
        {
            throw new XlsxException(
                $"Cell '{cell.Address}' in sheet '{sheetName}' has both 'value' and 'formula'. " +
                "A cell must use one or the other, not both.");
        }

        if (!hasValue && !hasFormula)
        {
            throw new XlsxException(
                $"Cell '{cell.Address}' in sheet '{sheetName}' has neither 'value' nor 'formula'. " +
                "Set one of them.");
        }

        if (hasFormula && !cell.Formula!.StartsWith("="))
        {
            throw new XlsxException(
                $"Formula in cell '{cell.Address}' sheet '{sheetName}' must start with '='. " +
                $"Got: '{cell.Formula}'");
        }

        if (cell.Type != null)
        {
            var validTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                { "number", "string", "boolean", "date" };
            if (!validTypes.Contains(cell.Type))
            {
                throw new XlsxException(
                    $"Cell '{cell.Address}' in sheet '{sheetName}' has unknown type '{cell.Type}'. " +
                    "Valid types: number, string, boolean, date.");
            }
        }
    }
}
