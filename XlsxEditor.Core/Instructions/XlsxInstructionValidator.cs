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
}
