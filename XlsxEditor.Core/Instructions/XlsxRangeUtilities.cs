using System.Text.RegularExpressions;
using XlsxEditor.Core.Builders;
using XlsxEditor.Core.Exceptions;

namespace XlsxEditor.Core.Instructions;

/// <summary>
/// Shared, builder-independent helpers for A1 ranges and Excel table-name rules.
/// Address parsing itself delegates to <see cref="WorksheetBuilder"/>'s internal
/// normalization so the validator, planner and (eventually) the executor can never
/// disagree about what a legal reference is.
/// </summary>
internal static class XlsxRangeUtilities
{
    /// <summary>Excel's real sheet bounds (columns A-XFD, rows 1-1,048,576).</summary>
    internal const int MaxColumns = 16384;
    internal const int MaxRows = 1_048_576;

    /// <summary>Excel cell text limit (characters).</summary>
    internal const int MaxCellTextLength = 32_767;

    /// <summary>Excel formula limit (characters, including the leading '=').</summary>
    internal const int MaxFormulaLength = 8_192;

    // Same grammar the builder uses for table display names.
    private static readonly Regex TableDisplayNamePattern =
        new(@"^[A-Za-z_\\][A-Za-z0-9._]*$", RegexOptions.Compiled);

    private static readonly Regex CellReferencePattern =
        new(@"^[A-Za-z]{1,3}[1-9][0-9]*$", RegexOptions.Compiled);

    private static readonly Regex R1C1ReferencePattern =
        new(@"^[Rr][1-9][0-9]*[Cc][1-9][0-9]*$", RegexOptions.Compiled);

    /// <summary>
    /// Splits an A1 range ("A1:C3") into its two canonical upper-case endpoint cells.
    /// Throws <see cref="XlsxException"/> for null, non-colon, malformed, out-of-bounds
    /// or (via the endpoints) empty input — mirroring the builder's merge-range contract.
    /// </summary>
    internal static (string start, string end) NormalizeRange(string range)
    {
        if (string.IsNullOrWhiteSpace(range))
        {
            throw new XlsxException("Range must not be null or whitespace.");
        }

        var bounds = range.Split(':');
        if (bounds.Length != 2)
        {
            throw new XlsxException(
                $"Invalid range '{range}'. Expected an A1-style range such as 'A1:C3' " +
                "(two cell references separated by a colon).");
        }

        return (
            WorksheetBuilder.NormalizeCellReference(bounds[0]),
            WorksheetBuilder.NormalizeCellReference(bounds[1]));
    }

    internal static bool IsValidTableName(string? name) =>
        !string.IsNullOrWhiteSpace(name) && TableDisplayNamePattern.IsMatch(name!);

    /// <summary>
    /// True when <paramref name="name"/> matches the grammar Excel parses as a cell
    /// reference (A1 style or R1C1 style) and the referenced cell exists within Excel's
    /// real sheet bounds. Mirrors the builder's rule so table names are rejected
    /// identically by validation and by the builder.
    /// </summary>
    internal static bool IsCellReferenceName(string name)
    {
        if (CellReferencePattern.IsMatch(name))
        {
            try
            {
                WorksheetBuilder.NormalizeCellReference(name);
                return true;
            }
            catch (XlsxException)
            {
                return false;
            }
        }

        if (!R1C1ReferencePattern.IsMatch(name))
        {
            return false;
        }

        var columnLetterIndex = name.IndexOfAny(new[] { 'C', 'c' });
        var rowText = name.AsSpan(1, columnLetterIndex - 1);
        var columnText = name.AsSpan(columnLetterIndex + 1);
        return long.TryParse(rowText, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var row)
            && long.TryParse(columnText, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var column)
            && row >= 1
            && row <= MaxRows
            && column >= 1
            && column <= MaxColumns;
    }

    /// <summary>
    /// Returns a 1-based column number from an A1 column letter ("A" = 1, "XFD" = 16384).
    /// Throws <see cref="XlsxException"/> for malformed or out-of-bounds letters.
    /// </summary>
    internal static int GetColumnNumber(string column)
    {
        var normalized = column.ToUpperInvariant();
        if (normalized.Length == 0 || normalized.Length > 3 || !normalized.All(char.IsAsciiLetter))
        {
            throw new XlsxException(
                $"Invalid column '{column}'. Expected an A1-style column letter such as " +
                "'A' or 'AB' (1-3 letters, no row number).");
        }

        var result = 0;
        foreach (var c in normalized)
        {
            result = result * 26 + c - 'A' + 1;
        }

        if (result > MaxColumns)
        {
            throw new XlsxException(
                $"Column '{column}' is beyond Excel's maximum column 'XFD'.");
        }

        return result;
    }
}
