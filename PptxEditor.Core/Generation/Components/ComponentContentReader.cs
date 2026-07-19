using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace PptxEditor.Core.Generation.Components;

/// <summary>
/// Reads a component's raw JSON content bag (<see cref="Model.ComponentElement.Content"/>)
/// into the strongly typed payloads of <c>ComponentContents.cs</c>, collecting loud,
/// actionable errors (JSON path + suggestion) in the style of the schema parser instead
/// of failing fast. Pure JSON→record validation: no layout, no emitter concerns.
/// </summary>
internal sealed class ComponentContentReader
{
    private static readonly Regex HexColorPattern = new(@"^#[0-9a-fA-F]{6}$", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private readonly string _path;
    private readonly string _componentName;
    private readonly JsonElement? _content;
    private readonly IReadOnlyDictionary<string, string> _palette;
    private readonly List<string> _errors = [];

    private ComponentContentReader(string path, string componentName, JsonElement? content, IReadOnlyDictionary<string, string> palette)
    {
        _path = path;
        _componentName = componentName;
        _content = content;
        _palette = palette;
    }

    /// <summary>
    /// Creates a reader for <paramref name="content"/>, validating presence, object kind
    /// and unknown properties (with "Did you mean …?" suggestions) up front.
    /// </summary>
    public static ComponentContentReader For(
        string path,
        string componentName,
        JsonElement? content,
        IReadOnlySet<string> allowedProps,
        IReadOnlyDictionary<string, string> palette,
        bool required = true)
    {
        var reader = new ComponentContentReader(path, componentName, content, palette);
        if (content is not { } c)
        {
            if (required)
            {
                reader._errors.Add($"{path}.content: 'content' is required for component '{componentName}'.");
            }
            return reader;
        }
        if (c.ValueKind != JsonValueKind.Object)
        {
            reader._errors.Add($"{path}.content: must be an object (component '{componentName}' payload, plan.md §4).");
            return reader;
        }
        foreach (var property in c.EnumerateObject())
        {
            if (!allowedProps.Contains(property.Name))
            {
                reader.Error(property.Name, $"unknown property '{property.Name}' on a '{componentName}' component.", Suggest(property.Name, allowedProps));
            }
        }
        return reader;
    }

    /// <summary>True when a content object is present and readable.</summary>
    public bool HasContent => _content is { ValueKind: JsonValueKind.Object };

    /// <summary>Reads a string property.</summary>
    public string? String(string name, bool required = false)
    {
        if (!TryGet(name, out var v))
        {
            if (required)
            {
                Error(name, $"'{name}' is required.");
            }
            return null;
        }
        if (v.ValueKind != JsonValueKind.String)
        {
            Error(name, "must be a string.");
            return null;
        }
        return v.GetString()!;
    }

    /// <summary>Reads a numeric property (points) within [min, max].</summary>
    public double? Number(string name, double min = double.MinValue, double max = double.MaxValue)
    {
        if (!TryGet(name, out var v))
        {
            return null;
        }
        if (v.ValueKind == JsonValueKind.String && v.GetString() is { } s && s.EndsWith('%'))
        {
            Error(name, $"'{s}' is a percentage: percentages are not supported (v1 non-goal, plan.md §1); use pt numbers.");
            return null;
        }
        if (v.ValueKind != JsonValueKind.Number || !v.TryGetDouble(out var number))
        {
            Error(name, "must be a number (pt).");
            return null;
        }
        if (number < min || number > max)
        {
            Error(name, $"must be between {Format(min)} and {Format(max)} (got {Format(number)}).");
            return null;
        }
        return number;
    }

    /// <summary>Reads a boolean property with a default.</summary>
    public bool Bool(string name, bool defaultValue)
    {
        if (!TryGet(name, out var v))
        {
            return defaultValue;
        }
        if (v.ValueKind == JsonValueKind.True)
        {
            return true;
        }
        if (v.ValueKind == JsonValueKind.False)
        {
            return false;
        }
        Error(name, "must be a boolean.");
        return defaultValue;
    }

    /// <summary>Reads a color property: a palette token name or a #RRGGBB literal.</summary>
    public string? Color(string name)
    {
        if (!TryGet(name, out var v))
        {
            return null;
        }
        if (v.ValueKind != JsonValueKind.String)
        {
            Error(name, "must be a color string: a palette token name or #RRGGBB.");
            return null;
        }
        var raw = v.GetString()!;
        if (_palette.ContainsKey(raw) || HexColorPattern.IsMatch(raw))
        {
            return raw;
        }
        if (raw.StartsWith('#'))
        {
            Error(name, $"'{raw}' is not a valid hex color: expected #RRGGBB.");
            return null;
        }
        Error(name, $"unknown color '{raw}': not a palette token and not #RRGGBB hex.", Suggest(raw, _palette.Keys));
        return null;
    }

    /// <summary>Reads an enum property from a string map.</summary>
    public TEnum? Enum<TEnum>(string name, IReadOnlyDictionary<string, TEnum> map, string what) where TEnum : struct
    {
        if (!TryGet(name, out var v))
        {
            return null;
        }
        if (v.ValueKind != JsonValueKind.String)
        {
            Error(name, $"must be a string ({what}: {string.Join(", ", map.Keys)}).");
            return null;
        }
        var raw = v.GetString()!;
        if (map.TryGetValue(raw, out var value))
        {
            return value;
        }
        Error(name, $"'{raw}' is not a valid {what}. Valid values: {string.Join(", ", map.Keys)}.", Suggest(raw, map.Keys));
        return null;
    }

    /// <summary>Reads an array of strings.</summary>
    public IReadOnlyList<string>? StringArray(string name, bool required = false, int minCount = 0)
    {
        if (!TryGet(name, out var v))
        {
            if (required)
            {
                Error(name, $"'{name}' is required.");
            }
            return null;
        }
        if (v.ValueKind != JsonValueKind.Array)
        {
            Error(name, "must be an array of strings.");
            return null;
        }
        var items = new List<string>();
        var index = 0;
        foreach (var item in v.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.String)
            {
                Error($"{name}[{index}]", "must be a string.");
            }
            else
            {
                items.Add(item.GetString()!);
            }
            index++;
        }
        if (items.Count < minCount)
        {
            Error(name, $"must contain at least {minCount} item(s).");
            return null;
        }
        return items;
    }

    /// <summary>Reads an array of string arrays (table rows).</summary>
    public IReadOnlyList<IReadOnlyList<string>>? StringMatrix(string name, bool required = false)
    {
        if (!TryGet(name, out var v))
        {
            if (required)
            {
                Error(name, $"'{name}' is required.");
            }
            return null;
        }
        if (v.ValueKind != JsonValueKind.Array)
        {
            Error(name, "must be an array of rows (arrays of strings).");
            return null;
        }
        var rows = new List<IReadOnlyList<string>>();
        var rowIndex = 0;
        foreach (var row in v.EnumerateArray())
        {
            if (row.ValueKind != JsonValueKind.Array)
            {
                Error($"{name}[{rowIndex}]", "must be an array of strings (one cell per column).");
                rowIndex++;
                continue;
            }
            var cells = new List<string>();
            var cellIndex = 0;
            foreach (var cell in row.EnumerateArray())
            {
                if (cell.ValueKind != JsonValueKind.String)
                {
                    Error($"{name}[{rowIndex}][{cellIndex}]", "must be a string.");
                }
                else
                {
                    cells.Add(cell.GetString()!);
                }
                cellIndex++;
            }
            rows.Add(cells);
            rowIndex++;
        }
        return rows;
    }

    /// <summary>Reads an array of positive numbers (grow weights).</summary>
    public IReadOnlyList<double>? NumberArray(string name)
    {
        if (!TryGet(name, out var v))
        {
            return null;
        }
        if (v.ValueKind != JsonValueKind.Array)
        {
            Error(name, "must be an array of numbers.");
            return null;
        }
        var numbers = new List<double>();
        var index = 0;
        foreach (var item in v.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.Number || !item.TryGetDouble(out var number))
            {
                Error($"{name}[{index}]", "must be a number.");
            }
            else if (number <= 0)
            {
                Error($"{name}[{index}]", $"must be > 0 (got {Format(number)}).");
            }
            else
            {
                numbers.Add(number);
            }
            index++;
        }
        return numbers;
    }

    /// <summary>Throws a <see cref="ComponentException"/> carrying every collected error.</summary>
    public void ThrowIfInvalid()
    {
        if (_errors.Count > 0)
        {
            throw new ComponentException(_path, string.Join("; ", _errors));
        }
    }

    private bool TryGet(string name, out JsonElement value)
    {
        if (_content is { ValueKind: JsonValueKind.Object } content)
        {
            foreach (var property in content.EnumerateObject())
            {
                if (string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    value = property.Value;
                    return true;
                }
            }
        }
        value = default;
        return false;
    }

    private void Error(string name, string message, string? suggestion = null)
        => _errors.Add(suggestion is null
            ? $"{_path}.content.{name}: {message}"
            : $"{_path}.content.{name}: {message} {suggestion}");

    private static string Format(double value) => value.ToString("0.###", CultureInfo.InvariantCulture);

    private static string? Suggest(string name, IEnumerable<string> candidates)
    {
        string? best = null;
        var bestDistance = int.MaxValue;
        foreach (var candidate in candidates)
        {
            var distance = Levenshtein(name, candidate);
            if (distance < bestDistance)
            {
                bestDistance = distance;
                best = candidate;
            }
        }
        return bestDistance <= 2 ? $"Did you mean '{best}'?" : null;
    }

    private static int Levenshtein(string a, string b)
    {
        a = a.ToLowerInvariant();
        b = b.ToLowerInvariant();
        var previous = new int[b.Length + 1];
        var current = new int[b.Length + 1];
        for (var j = 0; j <= b.Length; j++)
        {
            previous[j] = j;
        }
        for (var i = 1; i <= a.Length; i++)
        {
            current[0] = i;
            for (var j = 1; j <= b.Length; j++)
            {
                var cost = a[i - 1] == b[j - 1] ? 0 : 1;
                current[j] = Math.Min(Math.Min(current[j - 1] + 1, previous[j] + 1), previous[j - 1] + cost);
            }
            (previous, current) = (current, previous);
        }
        return previous[b.Length];
    }
}
