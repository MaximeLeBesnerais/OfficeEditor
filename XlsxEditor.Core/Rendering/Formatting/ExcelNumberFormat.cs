using System.Globalization;

namespace XlsxEditor.Core.Rendering.Formatting;

public enum FormatConditionOperator
{
    None,
    GreaterThan,
    GreaterOrEqual,
    LessThan,
    LessOrEqual,
    Equal,
    NotEqual
}

public readonly record struct FormatCondition(FormatConditionOperator Operator, double Value);

public enum FormatTokenType
{
    Digit0,
    DigitHash,
    DigitQ,
    Dot,
    Comma,
    Percent,
    SciE,
    TextPlaceholder,
    Literal,
    General,
    Slash,
    StarChar,
    UnderChar,
    Color,
    Condition
}

public readonly record struct FormatToken(FormatTokenType Type, string Text);

public sealed class FormatSection
{
    public FormatCondition? Condition { get; set; }
    public string? ColorArgb { get; set; }
    public List<FormatToken> Tokens { get; } = [];
}

public sealed class ExcelNumberFormat
{
    public List<FormatSection> Sections { get; } = [];
    public bool IsGeneral { get; }
    public bool IsDateFormat { get; }

    public ExcelNumberFormat(string code)
    {
        if (string.IsNullOrEmpty(code))
        {
            IsGeneral = true;
            Sections.Add(new FormatSection());
            return;
        }

        if (code.Equals("General", StringComparison.OrdinalIgnoreCase))
        {
            IsGeneral = true;
            Sections.Add(new FormatSection());
            return;
        }

        IsGeneral = false;

        var sectionParts = SplitSections(code);

        foreach (var part in sectionParts)
        {
            var section = ParseSection(part);
            Sections.Add(section);
        }

        if (Sections.Count == 0)
            Sections.Add(new FormatSection());

        IsDateFormat = ExcelDateFormat.IsDateFormat(code);
    }

    private static List<string> SplitSections(string code)
    {
        var parts = new List<string>();
        int start = 0;
        bool inQuote = false;

        for (int i = 0; i < code.Length; i++)
        {
            char c = code[i];
            if (c == '"')
                inQuote = !inQuote;
            else if (c == ';' && !inQuote)
            {
                parts.Add(code.Substring(start, i - start));
                start = i + 1;
            }
        }

        parts.Add(code.Substring(start));
        return parts;
    }

    private static FormatSection ParseSection(string raw)
    {
        var section = new FormatSection();
        int i = 0;

        (i, section.Condition, section.ColorArgb) = ParseBracketedPrefix(raw, i);

        section.Tokens.AddRange(TokenizeSection(raw, i));
        return section;
    }

    private static (int newIndex, FormatCondition? condition, string? color) ParseBracketedPrefix(
        string raw, int start)
    {
        FormatCondition? condition = null;
        string? color = null;
        int i = start;

        while (i < raw.Length && raw[i] == '[')
        {
            int close = raw.IndexOf(']', i + 1);
            if (close < 0) break;

            string content = raw.Substring(i + 1, close - i - 1);
            bool consumed = false;

            var col = ParseColor(content);
            if (col != null)
            {
                color = col;
                i = close + 1;
                consumed = true;
            }

            if (!consumed)
            {
                var cond = ParseCondition(content);
                if (cond != null)
                {
                    condition = cond;
                    i = close + 1;
                    consumed = true;
                }
            }

            if (!consumed)
                break;
        }

        return (i, condition, color);
    }

    internal static string? ParseColor(string content)
    {
        return content.ToLowerInvariant() switch
        {
            "black" => "FF000000",
            "blue" => "FF0000FF",
            "cyan" => "FF00FFFF",
            "green" => "FF00FF00",
            "magenta" => "FFFF00FF",
            "red" => "FFFF0000",
            "white" => "FFFFFFFF",
            "yellow" => "FFFFFF00",
            _ => TryParseColorN(content)
        };
    }

    private static string? TryParseColorN(string content)
    {
        if (content.StartsWith("color", StringComparison.OrdinalIgnoreCase)
            && int.TryParse(content.AsSpan(5), NumberStyles.None, CultureInfo.InvariantCulture, out int n)
            && n >= 0 && n <= 56)
        {
            return $"color{n}";
        }
        return null;
    }

    private static FormatCondition? ParseCondition(string content)
    {
        if (content.Length < 2) return null;

        string valStr;
        FormatConditionOperator op;

        if (content.StartsWith(">="))
        {
            op = FormatConditionOperator.GreaterOrEqual;
            valStr = content[2..];
        }
        else if (content.StartsWith("<="))
        {
            op = FormatConditionOperator.LessOrEqual;
            valStr = content[2..];
        }
        else if (content.StartsWith("<>"))
        {
            op = FormatConditionOperator.NotEqual;
            valStr = content[2..];
        }
        else if (content.StartsWith('>'))
        {
            op = FormatConditionOperator.GreaterThan;
            valStr = content[1..];
        }
        else if (content.StartsWith('<'))
        {
            op = FormatConditionOperator.LessThan;
            valStr = content[1..];
        }
        else if (content.StartsWith('='))
        {
            op = FormatConditionOperator.Equal;
            valStr = content[1..];
        }
        else
        {
            return null;
        }

        if (double.TryParse(valStr, NumberStyles.Float | NumberStyles.AllowThousands,
                CultureInfo.InvariantCulture, out double val))
        {
            return new FormatCondition(op, val);
        }

        return null;
    }

    internal static List<FormatToken> TokenizeSection(string raw, int start)
    {
        var tokens = new List<FormatToken>();
        int i = start;

        while (i < raw.Length)
        {
            char c = raw[i];

            if (c == '"')
            {
                int end = raw.IndexOf('"', i + 1);
                if (end < 0) end = raw.Length - 1;
                tokens.Add(new FormatToken(FormatTokenType.Literal,
                    raw.Substring(i + 1, end - i - 1)));
                i = end + 1;
                continue;
            }

            if (c == '\\' && i + 1 < raw.Length)
            {
                tokens.Add(new FormatToken(FormatTokenType.Literal,
                    raw[i + 1].ToString()));
                i += 2;
                continue;
            }

            if (c == '*' && i + 1 < raw.Length)
            {
                tokens.Add(new FormatToken(FormatTokenType.StarChar,
                    raw[i + 1].ToString()));
                i += 2;
                continue;
            }

            if (c == '_' && i + 1 < raw.Length)
            {
                tokens.Add(new FormatToken(FormatTokenType.UnderChar,
                    raw[i + 1].ToString()));
                i += 2;
                continue;
            }

            if (c == '0')
            {
                tokens.Add(new FormatToken(FormatTokenType.Digit0, "0"));
                i++;
                continue;
            }

            if (c == '#')
            {
                tokens.Add(new FormatToken(FormatTokenType.DigitHash, "#"));
                i++;
                continue;
            }

            if (c == '?')
            {
                tokens.Add(new FormatToken(FormatTokenType.DigitQ, "?"));
                i++;
                continue;
            }

            if (c == '.')
            {
                tokens.Add(new FormatToken(FormatTokenType.Dot, "."));
                i++;
                continue;
            }

            if (c == ',')
            {
                tokens.Add(new FormatToken(FormatTokenType.Comma, ","));
                i++;
                continue;
            }

            if (c == '%')
            {
                tokens.Add(new FormatToken(FormatTokenType.Percent, "%"));
                i++;
                continue;
            }

            if (c == '/')
            {
                tokens.Add(new FormatToken(FormatTokenType.Slash, "/"));
                i++;
                continue;
            }

            if (c == '@')
            {
                tokens.Add(new FormatToken(FormatTokenType.TextPlaceholder, "@"));
                i++;
                continue;
            }

            if (c == 'E' || c == 'e')
            {
                if (i + 2 < raw.Length
                    && (raw[i + 1] == '+' || raw[i + 1] == '-')
                    && char.IsDigit(raw[i + 2]))
                {
                    int len = 1;
                    while (i + len < raw.Length && (raw[i + len] == '+' || raw[i + len] == '-'
                        || char.IsDigit(raw[i + len]) || raw[i + len] == '0'))
                        len++;
                    tokens.Add(new FormatToken(FormatTokenType.SciE,
                        raw.Substring(i, len)));
                    i += len;
                    continue;
                }
            }

            if (char.IsLetter(c) || c == '/' || c == ':' || c == ' ' || c == '-'
                || c == '(' || c == ')')
            {
                tokens.Add(new FormatToken(FormatTokenType.Literal, c.ToString()));
                i++;
                continue;
            }

            tokens.Add(new FormatToken(FormatTokenType.Literal, c.ToString()));
            i++;
        }

        return tokens;
    }
}
