namespace XlsxEditor.Core.Instructions;

/// <summary>
/// The discrete cell types the instruction vocabulary understands. <see cref="Auto"/>
/// lets the planner infer the concrete type from the resolved value.
/// </summary>
public enum XlsxCellType
{
    Auto,
    String,
    Number,
    Boolean,
    Date,
    DateTime
}

public static class XlsxCellTypeParser
{
    /// <summary>
    /// Parses a case-insensitive vocabulary value ("auto", "string", "number",
    /// "boolean", "date", "datetime") into a <see cref="XlsxCellType"/>. null, empty
    /// and "auto" all map to <see cref="XlsxCellType.Auto"/> (the default). Returns
    /// false for anything unrecognized so the validator can emit a path-qualified
    /// diagnostic instead of letting an unknown type silently fall back.
    /// </summary>
    public static bool TryParse(string? value, out XlsxCellType type)
    {
        switch (value?.Trim().ToLowerInvariant())
        {
            case null:
            case "":
            case "auto":
                type = XlsxCellType.Auto;
                return true;
            case "string":
            case "str":
            case "text":
                type = XlsxCellType.String;
                return true;
            case "number":
            case "num":
            case "numeric":
                type = XlsxCellType.Number;
                return true;
            case "boolean":
            case "bool":
                type = XlsxCellType.Boolean;
                return true;
            case "date":
                type = XlsxCellType.Date;
                return true;
            case "datetime":
            case "date-time":
            case "dateTime":
                type = XlsxCellType.DateTime;
                return true;
            default:
                type = default;
                return false;
        }
    }

    /// <summary>
    /// Applies a declared type to a concrete value, resolving <see cref="XlsxCellType.Auto"/>
    /// by inspecting the value: formulas stay formulas, "true"/"false" are booleans,
    /// ISO-ish dates become date/datetime, parseable numbers become numbers, everything
    /// else is a string.
    /// </summary>
    public static XlsxCellType ResolveAuto(XlsxCellType declared, string? value)
    {
        if (declared != XlsxCellType.Auto)
        {
            return declared;
        }

        if (string.IsNullOrEmpty(value))
        {
            return XlsxCellType.String;
        }

        if (value.StartsWith('='))
        {
            return XlsxCellType.String;
        }

        if (value.Equals("true", StringComparison.OrdinalIgnoreCase)
            || value.Equals("false", StringComparison.OrdinalIgnoreCase))
        {
            return XlsxCellType.Boolean;
        }

        if (double.TryParse(
                value,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture,
                out var number)
            && !double.IsNaN(number)
            && !double.IsInfinity(number))
        {
            return XlsxCellType.Number;
        }

        // Checked after numbers so a year-only string like "2024" stays a number while a
        // real date ("2024-01-15", "01/15/2024 10:30") falls through to a date type.
        if (System.DateTime.TryParse(
                value,
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.AllowWhiteSpaces,
                out var date))
        {
            return date.TimeOfDay == System.TimeSpan.Zero
                ? XlsxCellType.Date
                : XlsxCellType.DateTime;
        }

        return XlsxCellType.String;
    }
}
