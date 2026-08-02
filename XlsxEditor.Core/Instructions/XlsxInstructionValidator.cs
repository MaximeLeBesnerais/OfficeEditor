using XlsxEditor.Core.Exceptions;

namespace XlsxEditor.Core.Instructions;

/// <summary>
/// Throwing facade over <see cref="XlsxValidationEngine"/>. Kept for compatibility with
/// callers and tests that rely on a single <see cref="XlsxException"/> instead of the
/// structured <see cref="XlsxValidationResult"/> — it validates via the same engine and
/// throws the first error's message. Programmatic instruction sets (bypassing the parser)
/// are held to exactly the same rules as parsed JSON.
/// </summary>
public static class XlsxInstructionValidator
{
    public static void Validate(XlsxInstructionSet instructions)
    {
        var result = XlsxValidationEngine.Validate(instructions);
        if (!result.IsValid)
        {
            var first = result.Errors.First();
            throw new XlsxException(
                string.IsNullOrEmpty(first.Path) ? first.Message : $"{first.Message} (at {first.Path})");
        }
    }

    /// <summary>
    /// The v1 executor is not wired to apply typed cells / number formats yet (the style
    /// and typed-cell builder primitives land on another branch). It must fail loudly
    /// rather than silently drop those fields. The model, validator and planner accept
    /// them; only execution rejects. The message intentionally keeps the roadmap phrasing
    /// so existing consumers see the familiar guidance.
    /// </summary>
    internal static void RejectUnsupportedCellFields(CellInstruction cell, string sheetName)
    {
        if (cell.Type != null)
        {
            throw new XlsxException(
                $"Cell '{cell.Address}' in sheet '{sheetName}' sets 'type' ('{cell.Type}'), which is " +
                "not yet supported by the v1 executor — previously it was silently ignored. Typed cells " +
                "arrive in Phase 2 of the XLSX roadmap (docs/roadmap-xlsx.md); remove the field to " +
                "execute with the v1 engine.");
        }

        if (cell.NumberFormat != null)
        {
            throw new XlsxException(
                $"Cell '{cell.Address}' in sheet '{sheetName}' sets 'numberFormat' " +
                $"('{cell.NumberFormat}'), which is not yet supported by the v1 executor — previously it " +
                "was silently ignored. Number formats arrive with typed cells in Phase 2 of the XLSX " +
                "roadmap (docs/roadmap-xlsx.md); remove the field to execute with the v1 engine.");
        }
    }
}
