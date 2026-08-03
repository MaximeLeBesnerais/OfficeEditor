using System.Globalization;

namespace XlsxEditor.Core.Rendering.Formatting;

public static class ExcelDateFormat
{
    private static readonly DateTime Epoch1900 = new(1899, 12, 31);
    private static readonly DateTime Epoch1904 = new(1904, 1, 1);
    private static readonly DateTime FakeLeapDay = new(1900, 2, 28);

    private static readonly string[] DayNames =
    [
        "Sunday", "Monday", "Tuesday", "Wednesday",
        "Thursday", "Friday", "Saturday"
    ];

    private static readonly string[] DayAbbrs =
    [
        "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"
    ];

    private static readonly string[] MonthNames =
    [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ];

    private static readonly string[] MonthAbbrs =
    [
        "Jan", "Feb", "Mar", "Apr", "May", "Jun",
        "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
    ];

    // Excel serial 1 == 1900-01-01, so the base epoch is 1899-12-31. Because Excel treats
    // 1900 as a leap year, serial 60 maps to the non-existent 1900-02-29; we represent that
    // fake day as 1900-02-28 so all serials round-trip (59 -> 1900-02-28, 60 -> 1900-02-28,
    // 61 -> 1900-03-01).

    public static DateTime SerialToDateTime(double serial, bool use1904System = false)
    {
        if (double.IsNaN(serial) || double.IsInfinity(serial))
            throw new ArgumentOutOfRangeException(nameof(serial), "Serial must be a finite number.");

        if (use1904System)
        {
            int days1904 = (int)Math.Floor(serial);
            double fraction = serial - days1904;
            return Epoch1904
                .AddDays(days1904)
                .Add(TimeSpan.FromDays(fraction));
        }

        int days = (int)Math.Floor(serial);
        double timeFraction = serial - days;

        if (days < 1)
            days = 1;

        if (days == 60)
            return FakeLeapDay.Add(TimeSpan.FromDays(timeFraction));

        DateTime date;
        if (days > 60)
            date = Epoch1900.AddDays(days - 1);
        else
            date = Epoch1900.AddDays(days);

        return date.Add(TimeSpan.FromDays(timeFraction));
    }

    public static double DateTimeToSerial(DateTime dateTime, bool use1904System = false)
    {
        if (use1904System)
            return (dateTime - Epoch1904).TotalDays;

        // Serial 60 is the fake leap day; .NET cannot represent 1900-02-29, so any real
        // date on/after 1900-03-01 adds one day to skip it.
        if (dateTime < new DateTime(1900, 3, 1))
        {
            double days = (dateTime.Date - Epoch1900).TotalDays;
            return days + dateTime.TimeOfDay.TotalDays;
        }

        double total = (dateTime.Date - Epoch1900).TotalDays + 1;
        return total + dateTime.TimeOfDay.TotalDays;
    }

    public static string FormatDate(DateTime dateTime, string code)
    {
        if (string.IsNullOrEmpty(code))
            return dateTime.ToString(CultureInfo.InvariantCulture);

        if (code.Equals("General", StringComparison.OrdinalIgnoreCase))
            return dateTime.ToString(CultureInfo.InvariantCulture);

        var tokens = TokenizeDateCode(code);
        var sb = new System.Text.StringBuilder();

        foreach (var token in tokens)
        {
            sb.Append(FormatDateToken(dateTime, token, code));
        }

        return sb.ToString();
    }

    public static bool IsDateFormat(string code)
    {
        if (string.IsNullOrEmpty(code))
            return false;
        if (code.Equals("General", StringComparison.OrdinalIgnoreCase))
            return false;

        var tokens = TokenizeDateCode(code);
        foreach (var t in tokens)
        {
            if (t.Type is DateTokenType.Year or DateTokenType.MonthNum or DateTokenType.DayNum
                or DateTokenType.Hour or DateTokenType.Second or DateTokenType.AmPm
                or DateTokenType.ElapsedHours or DateTokenType.ElapsedMinutes
                or DateTokenType.ElapsedSeconds or DateTokenType.MonthName
                or DateTokenType.MonthAbbr or DateTokenType.DayName or DateTokenType.DayAbbr)
                return true;
        }
        if (tokens.Count > 0)
        {
            foreach (var t in tokens)
            {
                if (t.Type == DateTokenType.Minute)
                    return true;
            }
        }
        return false;
    }

    public static bool HasTimeComponent(string code)
    {
        if (string.IsNullOrEmpty(code))
            return false;
        var tokens = TokenizeDateCode(code);
        foreach (var t in tokens)
        {
            if (t.Type is DateTokenType.Hour or DateTokenType.Minute or DateTokenType.Second
                or DateTokenType.AmPm or DateTokenType.ElapsedHours
                or DateTokenType.ElapsedMinutes or DateTokenType.ElapsedSeconds)
                return true;
        }
        return false;
    }

    private enum DateTokenType
    {
        Literal,
        Year,
        MonthNum,
        MonthName,
        MonthAbbr,
        DayNum,
        DayName,
        DayAbbr,
        Hour,
        Minute,
        Second,
        AmPm,
        ElapsedHours,
        ElapsedMinutes,
        ElapsedSeconds,
        StarChar,
        UnderChar,
        TextPlaceholder
    }

    private readonly record struct DateToken(DateTokenType Type, string Text, int Length);

    private static List<DateToken> TokenizeDateCode(string code)
    {
        var tokens = new List<DateToken>();
        int i = 0;
        bool hasHourOrSecond = false;
        bool hasAmPm = false;

        while (i < code.Length)
        {
            char c = code[i];

            if (c == '"')
            {
                int end = code.IndexOf('"', i + 1);
                if (end < 0) end = code.Length - 1;
                tokens.Add(new DateToken(DateTokenType.Literal, code.Substring(i + 1, end - i - 1), 0));
                i = end + 1;
                continue;
            }

            if (c == '\\' && i + 1 < code.Length)
            {
                tokens.Add(new DateToken(DateTokenType.Literal, code[i + 1].ToString(), 0));
                i += 2;
                continue;
            }

            if (c == '@')
            {
                tokens.Add(new DateToken(DateTokenType.TextPlaceholder, "@", 0));
                i++;
                continue;
            }

            if (c == '*' && i + 1 < code.Length)
            {
                tokens.Add(new DateToken(DateTokenType.StarChar, code[i + 1].ToString(), 0));
                i += 2;
                continue;
            }

            if (c == '_' && i + 1 < code.Length)
            {
                tokens.Add(new DateToken(DateTokenType.UnderChar, code[i + 1].ToString(), 0));
                i += 2;
                continue;
            }

            if (c == '[')
            {
                int close = code.IndexOf(']', i + 1);
                if (close > i)
                {
                    string brack = code.Substring(i + 1, close - i - 1).ToLowerInvariant();
                    if (brack == "h")
                        tokens.Add(new DateToken(DateTokenType.ElapsedHours, "[h]", 0));
                    else if (brack == "m")
                        tokens.Add(new DateToken(DateTokenType.ElapsedMinutes, "[m]", 0));
                    else if (brack == "s")
                        tokens.Add(new DateToken(DateTokenType.ElapsedSeconds, "[s]", 0));
                    else
                        tokens.Add(new DateToken(DateTokenType.Literal, code.Substring(i, close - i + 1), 0));
                    i = close + 1;
                    continue;
                }
            }

            if (c == 'y')
            {
                int len = CountRun(code, i, 'y');
                tokens.Add(new DateToken(DateTokenType.Year, i < code.Length ? code.Substring(i, len) : "", len));
                i += len;
                continue;
            }

            if (c == 'm' && !((hasHourOrSecond || hasAmPm) && CountRun(code, i, 'm') <= 2))
            {
                int len = CountRun(code, i, 'm');
                if (len >= 3)
                    tokens.Add(new DateToken(DateTokenType.MonthName, code.Substring(i, len), len));
                else
                    tokens.Add(new DateToken(DateTokenType.MonthNum, code.Substring(i, len), len));
                i += len;
                continue;
            }

            if (c == 'd')
            {
                int len = CountRun(code, i, 'd');
                if (len >= 3)
                    tokens.Add(new DateToken(DateTokenType.DayName, code.Substring(i, len), len));
                else
                    tokens.Add(new DateToken(DateTokenType.DayNum, code.Substring(i, len), len));
                i += len;
                continue;
            }

            if (c == 'h')
            {
                int len = CountRun(code, i, 'h');
                tokens.Add(new DateToken(DateTokenType.Hour, code.Substring(i, len), len));
                hasHourOrSecond = true;
                i += len;
                continue;
            }

            if (c == 's')
            {
                int len = CountRun(code, i, 's');
                tokens.Add(new DateToken(DateTokenType.Second, code.Substring(i, len), len));
                hasHourOrSecond = true;
                i += len;
                continue;
            }

            if (c == 'm')
            {
                int len = CountRun(code, i, 'm');
                tokens.Add(new DateToken(DateTokenType.Minute, code.Substring(i, len), len));
                i += len;
                continue;
            }

            if (char.ToUpperInvariant(c) == 'A' || char.ToUpperInvariant(c) == 'P')
            {
                if (i + 5 <= code.Length && code.AsSpan(i, 5).Equals("AM/PM", StringComparison.OrdinalIgnoreCase))
                {
                    tokens.Add(new DateToken(DateTokenType.AmPm, code.Substring(i, 5), 2));
                    hasAmPm = true;
                    i += 5;
                }
                else if (i + 3 <= code.Length && code.AsSpan(i, 3).Equals("A/P", StringComparison.OrdinalIgnoreCase))
                {
                    tokens.Add(new DateToken(DateTokenType.AmPm, code.Substring(i, 3), 1));
                    hasAmPm = true;
                    i += 3;
                }
                else
                {
                    tokens.Add(new DateToken(DateTokenType.Literal, c.ToString(), 0));
                    i++;
                }
                continue;
            }

            tokens.Add(new DateToken(DateTokenType.Literal, c.ToString(), 0));
            i++;
        }

        return tokens;
    }

    private static int CountRun(string s, int start, char c)
    {
        int count = 0;
        char lower = char.ToLowerInvariant(c);
        while (start + count < s.Length && char.ToLowerInvariant(s[start + count]) == lower)
            count++;
        return count;
    }

    private static string FormatDateToken(DateTime dt, DateToken token, string code)
    {
        switch (token.Type)
        {
            case DateTokenType.Literal:
                return token.Text;

            case DateTokenType.Year:
                return token.Length == 2
                    ? (dt.Year % 100).ToString("D2", CultureInfo.InvariantCulture)
                    : dt.Year.ToString("D4", CultureInfo.InvariantCulture);

            case DateTokenType.MonthNum:
                return token.Length == 2
                    ? dt.Month.ToString("D2", CultureInfo.InvariantCulture)
                    : dt.Month.ToString(CultureInfo.InvariantCulture);

            case DateTokenType.MonthName:
                if (dt.Month < 1 || dt.Month > 12) return "";
                return token.Length >= 4 ? MonthNames[dt.Month - 1] : MonthAbbrs[dt.Month - 1];

            case DateTokenType.MonthAbbr:
                if (dt.Month < 1 || dt.Month > 12) return "";
                return token.Length >= 4 ? MonthNames[dt.Month - 1] : MonthAbbrs[dt.Month - 1];

            case DateTokenType.DayNum:
                return token.Length == 2
                    ? dt.Day.ToString("D2", CultureInfo.InvariantCulture)
                    : dt.Day.ToString(CultureInfo.InvariantCulture);

            case DateTokenType.DayName:
                int dow = (int)dt.DayOfWeek;
                return token.Length >= 4 ? DayNames[dow] : DayAbbrs[dow];

            case DateTokenType.DayAbbr:
                return DayAbbrs[(int)dt.DayOfWeek];

            case DateTokenType.Hour:
            {
                int h = dt.Hour;
                int display = h;
                bool is12 = HasAmPm(code) || IsAmPmInTokens(token);
                if (is12)
                {
                    display = h == 0 ? 12 : (h > 12 ? h - 12 : h);
                }
                return token.Length >= 2
                    ? display.ToString("D2", CultureInfo.InvariantCulture)
                    : display.ToString(CultureInfo.InvariantCulture);
            }

            case DateTokenType.Minute:
                return token.Length >= 2
                    ? dt.Minute.ToString("D2", CultureInfo.InvariantCulture)
                    : dt.Minute.ToString(CultureInfo.InvariantCulture);

            case DateTokenType.Second:
                return token.Length >= 2
                    ? dt.Second.ToString("D2", CultureInfo.InvariantCulture)
                    : dt.Second.ToString(CultureInfo.InvariantCulture);

            case DateTokenType.AmPm:
            {
                bool pm = dt.Hour >= 12;
                string raw = token.Text;
                // raw is the matched slice, e.g. "AM/PM", "am/pm", "A/P", "a/p", "A", "a".
                bool upper = raw.All(c => !char.IsLetter(c) || char.IsUpper(c));
                string am = upper ? "AM" : "am";
                string pmStr = upper ? "PM" : "pm";
                if (raw.Contains('/'))
                    return pm ? pmStr : am;
                // Single-char form "A"/"a".
                return pm
                    ? (upper ? "P" : "p")
                    : (upper ? "A" : "a");
            }

            case DateTokenType.ElapsedHours:
                return ((int)(dt - Epoch1900).TotalHours).ToString(CultureInfo.InvariantCulture);

            case DateTokenType.ElapsedMinutes:
                return ((int)(dt - Epoch1900).TotalMinutes).ToString(CultureInfo.InvariantCulture);

            case DateTokenType.ElapsedSeconds:
                return ((int)(dt - Epoch1900).TotalSeconds).ToString(CultureInfo.InvariantCulture);

            case DateTokenType.StarChar:
                return "";

            case DateTokenType.UnderChar:
                return " ";

            case DateTokenType.TextPlaceholder:
                return dt.ToString(CultureInfo.InvariantCulture);

            default:
                return token.Text;
        }
    }

    private static bool HasAmPm(string code)
    {
        return code.Contains("AM/PM", StringComparison.OrdinalIgnoreCase)
            || code.Contains("am/pm", StringComparison.Ordinal)
            || code.Contains("A/P", StringComparison.Ordinal)
            || code.Contains("a/p", StringComparison.Ordinal);
    }

    private static bool IsAmPmInTokens(DateToken token)
    {
        return token.Type == DateTokenType.AmPm;
    }
}
