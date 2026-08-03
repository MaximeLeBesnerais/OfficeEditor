using System.Globalization;
using System.Text;

namespace XlsxEditor.Core.Rendering.Formatting;

/// <summary>
/// Formats numeric values according to Excel number-format codes, producing display
/// strings (and optional ARGB colors) for a renderer. Culture-invariant; uses ASCII
/// hyphen for negative numbers, rounding half-away-from-zero to match Excel.
/// </summary>
public static class ExcelNumberFormatFormatter
{
    private const int MaxDisplayWidth = 255;
    private const int OverflowHashes = 30;

    public static (string Text, string? ColorArgb, bool UsedTextSection)
        FormatNumber(double value, string code)
    {
        var fmt = new ExcelNumberFormat(code);

        if (fmt.IsGeneral)
            return (FormatGeneral(value), null, false);

        return FormatWithSections(value, fmt, code, null);
    }

    public static (string Text, string? ColorArgb, bool UsedTextSection)
        FormatNumber(decimal value, string code)
    {
        return FormatNumber((double)value, code);
    }

    private static (string Text, string? ColorArgb, bool UsedTextSection)
        FormatWithSections(double value, ExcelNumberFormat fmt, string code, string? textValue)
    {
        var sections = fmt.Sections;

        if (sections.Count == 0)
            return (FormatGeneral(value), null, false);

        // Date/date-time formats use a single section; route them through the date
        // formatter which has its own dedicated tokenizer.
        if (fmt.IsDateFormat)
        {
            var dt = ExcelDateFormat.SerialToDateTime(value);
            return (ExcelDateFormat.FormatDate(dt, code), null, false);
        }

        FormatSection? selected = null;

        if (textValue != null)
        {
            if (sections.Count >= 4)
                selected = sections[3];
            else
            {
                foreach (var s in sections)
                {
                    if (HasTextPlaceholder(s))
                    {
                        selected = s;
                        break;
                    }
                }

                if (selected == null)
                    return (textValue, null, true);
            }
        }

        if (selected == null)
        {
            for (int i = 0; i < sections.Count; i++)
            {
                var s = sections[i];
                if (s.Condition != null && EvaluateCondition(s.Condition.Value, value))
                {
                    selected = s;
                    break;
                }
            }
        }

        if (selected == null)
        {
            if (sections.Count == 1)
            {
                selected = sections[0];
            }
            else if (sections.Count == 2)
            {
                selected = value >= 0 ? sections[0] : sections[1];
            }
            else if (sections.Count >= 3)
            {
                if (value > 0)
                    selected = sections[0];
                else if (value < 0)
                    selected = sections[1];
                else
                    selected = sections[2];
            }
        }

        if (selected == null)
            selected = sections[0];

        string? colorArgb = selected.ColorArgb ?? GetSectionColorFromTokens(selected);
        bool usedTextSection = false;

        string result;
        if (HasTextPlaceholder(selected) && textValue != null)
        {
            result = FormatTextSection(selected, textValue);
            usedTextSection = true;
        }
        else
        {
            // Excel only auto-adds a negative sign when there is a single section; an
            // explicit negative/zero/text section fully controls its own sign.
            bool autoSign = sections.Count == 1;
            result = FormatNumberSection(value, selected, autoSign);
        }

        if (result.Length > MaxDisplayWidth)
            result = new string('#', OverflowHashes);

        return (result, colorArgb, usedTextSection);
    }

    private static bool EvaluateCondition(FormatCondition condition, double value)
    {
        return condition.Operator switch
        {
            FormatConditionOperator.GreaterThan => value > condition.Value,
            FormatConditionOperator.GreaterOrEqual => value >= condition.Value,
            FormatConditionOperator.LessThan => value < condition.Value,
            FormatConditionOperator.LessOrEqual => value <= condition.Value,
            FormatConditionOperator.Equal => Math.Abs(value - condition.Value) < 1e-10,
            FormatConditionOperator.NotEqual => Math.Abs(value - condition.Value) >= 1e-10,
            _ => true
        };
    }

    private static string? GetSectionColorFromTokens(FormatSection section)
    {
        foreach (var t in section.Tokens)
        {
            if (t.Type == FormatTokenType.Color)
                return t.Text;
        }

        return null;
    }

    private static bool HasTextPlaceholder(FormatSection section)
    {
        foreach (var t in section.Tokens)
        {
            if (t.Type == FormatTokenType.TextPlaceholder)
                return true;
        }

        return false;
    }

    private static string FormatTextSection(FormatSection section, string textValue)
    {
        var sb = new StringBuilder();
        foreach (var token in section.Tokens)
        {
            if (token.Type == FormatTokenType.TextPlaceholder)
                sb.Append(textValue);
            else if (token.Type == FormatTokenType.Literal)
                sb.Append(token.Text);
            else if (token.Type is FormatTokenType.StarChar or FormatTokenType.UnderChar)
                continue;
            else
                sb.Append(token.Text);
        }

        return sb.ToString();
    }

    private static string FormatNumberSection(double value, FormatSection section, bool autoSign)
    {
        bool isNegative = value < 0;
        double workValue = Math.Abs(value);

        bool hasPercent = false;
        bool hasSci = false;
        string sciFormat = "";
        bool hasDecimal = false;
        int scaleCommas = 0;
        bool? grouping = null;

        int int0 = 0, intHash = 0;
        int dec0 = 0, decHash = 0;
        var literalsBefore = new List<string>();
        var literalsAfter = new List<string>();
        bool sawDigit = false;
        bool afterDecimal = false;
        bool hasDigitTokens = false;

        var tokens = section.Tokens;

        for (int i = 0; i < tokens.Count; i++)
        {
            var t = tokens[i];

            switch (t.Type)
            {
                case FormatTokenType.Percent:
                    hasPercent = true;
                    break;
                case FormatTokenType.SciE:
                    hasSci = true;
                    sciFormat = t.Text;
                    break;
                case FormatTokenType.Slash:
                    // Fraction formats are not yet supported; treated as literal.
                    if (sawDigit)
                        literalsAfter.Add(t.Text);
                    else
                        literalsBefore.Add(t.Text);
                    break;
                case FormatTokenType.Dot:
                    hasDecimal = true;
                    afterDecimal = true;
                    break;
                case FormatTokenType.Comma:
                    if (IsTrailingComma(tokens, i))
                        scaleCommas++;
                    else
                        grouping ??= true;
                    break;
                case FormatTokenType.Digit0:
                case FormatTokenType.DigitHash:
                case FormatTokenType.DigitQ:
                    hasDigitTokens = true;
                    sawDigit = true;
                    if (afterDecimal)
                    {
                        if (t.Type == FormatTokenType.Digit0) dec0++;
                        else if (t.Type == FormatTokenType.DigitHash) decHash++;
                    }
                    else
                    {
                        if (t.Type == FormatTokenType.Digit0) int0++;
                        else if (t.Type == FormatTokenType.DigitHash) intHash++;
                    }

                    break;
                case FormatTokenType.Literal:
                    if (sawDigit)
                        literalsAfter.Add(t.Text);
                    else
                        literalsBefore.Add(t.Text);
                    break;
            }
        }

        if (!hasDigitTokens)
        {
            var sbLit = new StringBuilder();
            foreach (var t in tokens)
            {
                if (t.Type == FormatTokenType.Literal)
                    sbLit.Append(t.Text);
                else if (t.Type is FormatTokenType.StarChar or FormatTokenType.UnderChar or FormatTokenType.Comma)
                    continue;
            }

            return sbLit.ToString();
        }

        if (hasPercent)
            workValue *= 100;

        if (hasSci && !string.IsNullOrEmpty(sciFormat))
            return FormatScientific(workValue, isNegative, dec0 + decHash);

        for (int s = 0; s < scaleCommas; s++)
            workValue /= 1000;

        int maxDec = dec0 + decHash;
        int minInt = int0;

        // Round the whole value once so carries propagate into the integer part.
        double rounded = Math.Round(workValue, maxDec, MidpointRounding.AwayFromZero);
        if (rounded == -0.0)
            rounded = 0.0;

        decimal roundedDec;
        try
        {
            roundedDec = (decimal)rounded;
        }
        catch (OverflowException)
        {
            roundedDec = rounded > 0 ? decimal.MaxValue : decimal.MinValue;
        }

        long intPart = (long)Math.Floor(roundedDec < 0 ? -roundedDec : roundedDec);
        decimal fracDec = (roundedDec < 0 ? -roundedDec : roundedDec) - intPart;
        long scaledFrac = maxDec > 0 ? (long)(fracDec * PowerOfTen(maxDec) + 0.5m) : 0;
        if (scaledFrac >= (long)PowerOfTen(maxDec))
        {
            scaledFrac -= (long)PowerOfTen(maxDec);
            intPart++;
        }

        string formattedInt = FormatIntegerPart(intPart, minInt, grouping ?? false);
        string formattedDec = FormatDecimalPart(scaledFrac, maxDec, dec0);

        var result = new StringBuilder();
        foreach (var lit in literalsBefore)
            result.Append(lit);

        result.Append(formattedInt);

        if (hasDecimal)
            result.Append('.');

        result.Append(formattedDec);

        foreach (var lit in literalsAfter)
            result.Append(lit);

        if (hasPercent)
            result.Append('%');

        if (isNegative && autoSign)
            result.Insert(0, '-');

        string str = result.ToString();
        return str.Length > MaxDisplayWidth ? new string('#', OverflowHashes) : str;
    }

    private static decimal PowerOfTen(int n)
    {
        decimal p = 1;
        for (int i = 0; i < n; i++)
            p *= 10m;
        return p;
    }

    private static bool IsTrailingComma(List<FormatToken> tokens, int index)
    {
        for (int j = index + 1; j < tokens.Count; j++)
        {
            var t = tokens[j];
            if (t.Type is FormatTokenType.Digit0 or FormatTokenType.DigitHash
                or FormatTokenType.DigitQ or FormatTokenType.Dot)
                return false;
            if (t.Type == FormatTokenType.Comma)
                continue;
            return true;
        }

        return true;
    }

    private static string FormatScientific(double value, bool isNegative, int mantissaDecimals)
    {
        string sign = isNegative ? "-" : "";
        string mantissa;
        string exponent;

        if (value == 0)
        {
            mantissa = mantissaDecimals > 0
                ? "0." + new string('0', mantissaDecimals)
                : "0";
            exponent = "00";
        }
        else
        {
            double exp = Math.Floor(Math.Log10(value));
            double man = value / Math.Pow(10, exp);
            man = Math.Round(man, mantissaDecimals, MidpointRounding.AwayFromZero);
            if (man >= 10.0)
            {
                man /= 10.0;
                exp += 1;
            }

            mantissa = mantissaDecimals > 0
                ? man.ToString("F" + mantissaDecimals, CultureInfo.InvariantCulture)
                : man.ToString("0", CultureInfo.InvariantCulture);
            exponent = ((int)exp).ToString("D2", CultureInfo.InvariantCulture);
        }

        return $"{sign}{mantissa}E+{exponent}";
    }

    private static string FormatIntegerPart(long intPart, int minDigits, bool grouping)
    {
        string digits = intPart.ToString(CultureInfo.InvariantCulture);
        int required = Math.Max(digits.Length, minDigits);
        var sb = new StringBuilder(required + (grouping ? required / 3 + 2 : 0));

        for (int i = 0; i < required - digits.Length; i++)
            sb.Append('0');

        for (int i = 0; i < digits.Length; i++)
        {
            if (grouping)
            {
                int remaining = digits.Length - i;
                if (remaining > 0 && remaining % 3 == 0 && sb.Length > 0 && sb[^1] != ',')
                    sb.Append(',');
            }

            sb.Append(digits[i]);
        }

        return sb.ToString();
    }

    private static string FormatDecimalPart(long scaledFrac, int maxDec, int minDec)
    {
        if (maxDec <= 0)
            return "";

        string s = scaledFrac.ToString("D" + maxDec, CultureInfo.InvariantCulture);

        // Trim trailing zeros beyond the forced minimum, preserving at least minDec.
        if (minDec < maxDec)
        {
            int trim = maxDec;
            while (trim > minDec && s[trim - 1] == '0')
                trim--;
            if (trim < maxDec)
                s = s[..trim];
        }

        return s;
    }

    internal static string FormatGeneral(double value)
    {
        if (double.IsNaN(value))
            return "";
        if (double.IsPositiveInfinity(value))
            return "Infinity";
        if (double.IsNegativeInfinity(value))
            return "-Infinity";

        double abs = Math.Abs(value);

        if (abs == 0)
            return "0";

        if (abs >= 1e11 || (abs < 1e-10 && abs > 0))
            return FormatScientificGeneral(value);

        string s = value.ToString("G15", CultureInfo.InvariantCulture);

        if (s.Contains('E') || s.Contains('e') || s.Length > 15)
            return FormatScientificGeneral(value);

        if (s.Contains('.'))
        {
            s = s.TrimEnd('0');
            if (s.EndsWith('.'))
                s = s[..^1];
        }

        if (s.Length > MaxDisplayWidth)
            return FormatScientificGeneral(value);

        return s;
    }

    private static string FormatScientificGeneral(double value)
    {
        if (value == 0)
            return "0";
        // Force scientific notation: "0.###E+00" guarantees an exponent.
        string s = value.ToString("0.###E+00", CultureInfo.InvariantCulture);
        if (s.Length > MaxDisplayWidth)
            return new string('#', OverflowHashes);
        return s;
    }
}
