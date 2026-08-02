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

        if (instructions.Worksheets is null || instructions.Worksheets.Count == 0)
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
        if (ws == null)
        {
            throw new XlsxException("Every entry in 'worksheets' must be a worksheet object.");
        }

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

        if (ws.Name[0] == '\'' || ws.Name[^1] == '\'')
        {
            throw new XlsxException(
                $"Worksheet name '{ws.Name}' begins or ends with an apostrophe ('), " +
                "which Excel does not allow. Remove the leading or trailing apostrophe.");
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

        if (ws.Headers != null)
        {
            for (var i = 0; i < ws.Headers.Count; i++)
            {
                if (ws.Headers[i] == null)
                {
                    throw new XlsxException(
                        $"Header {i + 1} of worksheet '{ws.Name}' is null; headers must be strings.");
                }
            }
        }

        if (ws.Rows != null)
        {
            for (var i = 0; i < ws.Rows.Count; i++)
            {
                var row = ws.Rows[i];
                if (row == null)
                {
                    throw new XlsxException(
                        $"Row {i + 1} of worksheet '{ws.Name}' is null; each row must be an array of cell values.");
                }

                for (var j = 0; j < row.Count; j++)
                {
                    if (row[j] == null)
                    {
                        throw new XlsxException(
                            $"Cell {j + 1} of row {i + 1} in worksheet '{ws.Name}' is null; cell values must be strings.");
                    }
                }
            }
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
        if (cell == null)
        {
            throw new XlsxException($"Cell instruction in sheet '{sheetName}' must be a cell object.");
        }

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

        // Enforce Excel's real limits (columns A-XFD, rows 1-1,048,576) through the
        // same shared parser the builders use, so validation and execution can never
        // disagree about what is a legal address.
        WorksheetBuilder.NormalizeCellReference(cell.Address);

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

        RejectUnsupportedCellFields(cell, sheetName);
    }

    /// <summary>
    /// 'type' and 'numberFormat' were validated by the schema but then silently
    /// dropped by the executor. Per the owner-approved decision they are rejected
    /// loudly until the typed-cell work lands (Phase 2 of the XLSX roadmap).
    /// Called from both validation and execution so programmatically constructed
    /// instruction sets (which bypass the parser) are rejected too.
    /// </summary>
    internal static void RejectUnsupportedCellFields(CellInstruction cell, string sheetName)
    {
        if (cell.Type != null)
        {
            throw new XlsxException(
                $"Cell '{cell.Address}' in sheet '{sheetName}' sets 'type' ('{cell.Type}'), which is " +
                "not yet supported — previously it was silently ignored. Typed cells arrive in " +
                "Phase 2 of the XLSX roadmap (docs/roadmap-xlsx.md); remove the field for now.");
        }

        if (cell.NumberFormat != null)
        {
            throw new XlsxException(
                $"Cell '{cell.Address}' in sheet '{sheetName}' sets 'numberFormat' " +
                $"('{cell.NumberFormat}'), which is not yet supported — previously it was silently " +
                "ignored. Number formats arrive with typed cells in Phase 2 of the XLSX roadmap " +
                "(docs/roadmap-xlsx.md); remove the field for now.");
        }
    }
}
