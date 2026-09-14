using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace OfficeEditor.Core.Generation;

/// <summary>An authored element snapshot. Properties are a detached JSON copy.</summary>
public sealed record GenerationJsonElement(string Id, string Kind, string Path, string? ParentId, JsonObject Properties);

/// <summary>Normalized source and the count of newly assigned element IDs.</summary>
public sealed record GenerationJsonIdResult(string Json, int AddedCount);

/// <summary>
/// Adds persistent, readable IDs to recognized PPTX/DOCX/XLSX generation JSON.
/// IDs are editable names, unique across the document. Save the returned JSON:
/// deterministic defaults alone cannot preserve identity after reordering ID-less input.
/// </summary>
public static class GenerationJsonIds
{
    private static readonly Regex Pattern = new(@"\A[a-z][a-z0-9_-]*\z", RegexOptions.CultureInvariant);
    private static readonly JsonSerializerOptions WriteOptions = new() { WriteIndented = true };

    public static bool IsGenerationDocument(string json)
    {
        ArgumentNullException.ThrowIfNull(json);
        try
        {
            using var document = JsonDocument.Parse(json);
            return document.RootElement.ValueKind == JsonValueKind.Object
                && document.RootElement.EnumerateObject().Any(p =>
                    p.Name.Equals("version", StringComparison.OrdinalIgnoreCase))
                && document.RootElement.EnumerateObject().Count(p =>
                    p.Value.ValueKind == JsonValueKind.Array && IsRootCollection(p.Name)) == 1;
        }
        catch (JsonException)
        {
            return false;
        }
    }

    /// <summary>Preserves the original bytes-as-text when no IDs need adding.</summary>
    public static GenerationJsonIdResult Normalize(string json)
    {
        var root = Read(json);
        var missing = Walk(root).Where(e => Key(e.Node, "id") is null).ToDictionary(e => e.Path);
        int count = Assign(root);
        return new(count == 0 ? json : InsertIds(json, missing), count);
    }

    /// <summary>Atomically updates a generation source file only when IDs are missing.</summary>
    public static GenerationJsonIdResult NormalizeFile(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        var original = File.ReadAllText(path);
        var result = Normalize(original);
        if (result.AddedCount == 0) return result;
        var fullPath = Path.GetFullPath(path);
        var temporary = fullPath + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            File.WriteAllText(temporary, result.Json);
            if (File.ReadAllText(fullPath) != original)
                throw new IOException("Source JSON changed during ID normalization; refusing to overwrite it.");
            File.Move(temporary, fullPath, overwrite: true);
        }
        finally
        {
            if (File.Exists(temporary)) File.Delete(temporary);
        }
        return result;
    }

    public static IReadOnlyList<GenerationJsonElement> Inspect(string json)
    {
        var root = Read(json);
        Assign(root);
        return Walk(root).Select(e => new GenerationJsonElement(Id(e.Node)!, e.Kind, e.Path,
            e.Parent is null ? null : Id(e.Parent), (JsonObject)e.Node.DeepClone())).ToArray();
    }

    /// <summary>Returns a detached element. Throws when the ID does not exist.</summary>
    public static GenerationJsonElement GetElement(string json, string id) =>
        Inspect(json).SingleOrDefault(e => e.Id == id)
        ?? throw new KeyNotFoundException($"No element with id '{id}'.");

    /// <summary>Renaming changes identity; duplicate/invalid new names fail without modifying input.</summary>
    public static string Rename(string json, string id, string newId) =>
        SetProperties(json, id, new JsonObject { ["id"] = newId });

    /// <summary>
    /// Replaces the supplied properties on one authored element, retaining other properties.
    /// Nested objects/arrays are replaced whole. Validate the returned source with its format
    /// parser before generation; this helper validates identity, not visual vocabulary.
    /// </summary>
    public static string SetProperties(string json, string id, JsonObject properties)
    {
        ArgumentNullException.ThrowIfNull(properties);
        var root = Read(json);
        Assign(root);
        var target = Walk(root).SingleOrDefault(e => Id(e.Node) == id)
            ?? throw new KeyNotFoundException($"No element with id '{id}'.");
        foreach (var property in properties)
        {
            var key = Key(target.Node, property.Key) ?? property.Key;
            target.Node[key] = property.Value?.DeepClone();
        }
        Assign(root);
        return root.ToJsonString(WriteOptions) + Environment.NewLine;
    }

    // Used by format parsers so their existing errors for non-document JSON stay intact.
    internal static string NormalizeIfRecognized(string json) =>
        IsGenerationDocument(json) ? Normalize(json).Json : json;

    private static JsonObject Read(string json)
    {
        if (!IsGenerationDocument(json))
            throw new JsonException("Expected generation JSON with version and one of slides, sections or worksheets.");
        var root = JsonNode.Parse(json)!.AsObject();
        return root;
    }

    private static int Assign(JsonObject root)
    {
        var elements = Walk(root).ToArray();
        var taken = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var element in elements)
        {
            var key = Key(element.Node, "id");
            if (key is null) continue;
            if (element.Node[key] is not JsonValue value || !value.TryGetValue<string>(out var id) || !Pattern.IsMatch(id))
                throw new JsonException($"{element.Path}.id: expected [a-z][a-z0-9_-]*.");
            if (!taken.TryAdd(id, element.Path))
                throw new JsonException($"{element.Path}.id: duplicate id '{id}', first declared at {taken[id]}.");
        }
        var counters = new Dictionary<string, int>(StringComparer.Ordinal);
        int added = 0;
        foreach (var element in elements.Where(e => Key(e.Node, "id") is null))
        {
            string type = Regex.Replace(element.Kind.ToLowerInvariant(), "[^a-z0-9_-]", "-");
            if (type.Length == 0 || type[0] is < 'a' or > 'z') type = "element";
            int next = counters.GetValueOrDefault(type);
            string id;
            do { id = $"{type}-{next++}"; } while (taken.ContainsKey(id));
            counters[type] = next;
            taken.Add(id, element.Path);
            element.Node["id"] = id;
            added++;
        }
        return added;
    }

    // Insert at JSON token offsets rather than reserialize the file: preserve comments-free
    // JSON formatting, Unicode, numeric spellings, property order and the final newline.
    private static string InsertIds(string json, IReadOnlyDictionary<string, Entry> missing)
    {
        byte[] input = Encoding.UTF8.GetBytes(json);
        var reader = new Utf8JsonReader(input);
        var frames = new Stack<JsonFrame>();
        var additions = new List<(int Offset, byte[] Text)>();
        while (reader.Read())
        {
            if (reader.TokenType is JsonTokenType.EndObject or JsonTokenType.EndArray)
            {
                frames.Pop();
                continue;
            }
            if (reader.TokenType == JsonTokenType.PropertyName)
            {
                frames.Peek().Property = reader.GetString()!;
                continue;
            }
            string path = "$";
            if (frames.TryPeek(out var parent))
                path = parent.IsArray ? $"{parent.Path}[{parent.Index++}]" : parent.Path + "." + parent.Property;
            if (reader.TokenType is JsonTokenType.StartObject or JsonTokenType.StartArray)
            {
                frames.Push(new JsonFrame(path, reader.TokenType == JsonTokenType.StartArray));
                if (reader.TokenType != JsonTokenType.StartObject || !missing.TryGetValue(path, out var entry)) continue;
                int start = checked((int)reader.TokenStartIndex + 1);
                int next = start;
                while (next < input.Length && input[next] is (byte)' ' or (byte)'\r' or (byte)'\n' or (byte)'\t') next++;
                string whitespace = Encoding.UTF8.GetString(input, start, next - start);
                string separator = whitespace.Length == 0 ? " " : whitespace;
                string property = $"\"id\": {JsonSerializer.Serialize(Id(entry.Node))}";
                if (input[next] != (byte)'}') property += "," + separator;
                additions.Add((next, Encoding.UTF8.GetBytes(property)));
            }
        }
        using var output = new MemoryStream();
        int cursor = 0;
        foreach (var addition in additions)
        {
            output.Write(input, cursor, addition.Offset - cursor);
            output.Write(addition.Text);
            cursor = addition.Offset;
        }
        output.Write(input, cursor, input.Length - cursor);
        return Encoding.UTF8.GetString(output.ToArray());
    }

    private sealed class JsonFrame(string path, bool isArray)
    {
        public string Path { get; } = path;
        public bool IsArray { get; } = isArray;
        public int Index { get; set; }
        public string Property { get; set; } = string.Empty;
    }

    private sealed record Entry(JsonObject Node, string Kind, string Path, JsonObject? Parent);

    private static IEnumerable<Entry> Walk(JsonObject root)
    {
        string rootName = root.First(p => IsRootCollection(p.Key) && p.Value is JsonArray).Key;
        string format = rootName.ToLowerInvariant();
        return VisitArray(root[rootName]!.AsArray(), "$." + rootName, null, format,
            format switch { "slides" => "container", "sections" => "section", _ => "worksheet" });
    }

    private static IEnumerable<Entry> VisitArray(JsonArray array, string path, JsonObject? parent, string format, string defaultKind)
    {
        for (int i = 0; i < array.Count; i++)
        {
            if (array[i] is not JsonObject node) continue; // the format validator diagnoses invalid entries
            string kind = defaultKind;
            if (format != "worksheets" && Get(node, "type") is JsonValue type && type.TryGetValue<string>(out var name))
                kind = name;
            if (format == "slides" && parent is null && kind.Equals("container", StringComparison.OrdinalIgnoreCase)) kind = "slide";
            string nodePath = $"{path}[{i}]";
            yield return new(node, kind, nodePath, parent);
            string[] children = format switch
            {
                "slides" => ["children"],
                "sections" => ["header", "footer", "blocks", "positioned"],
                _ when defaultKind == "worksheet" => ["cells", "columns", "tables"],
                _ => []
            };
            foreach (string field in node.Select(p => p.Key).Where(key => children.Contains(key, StringComparer.OrdinalIgnoreCase)))
            {
                if (Get(node, field) is not JsonArray nested) continue;
                string childKind = field.ToLowerInvariant() switch { "cells" => "cell", "columns" => "column", "tables" => "table", _ => "element" };
                foreach (var entry in VisitArray(nested, nodePath + "." + field, node, format, childKind))
                    yield return entry;
            }
        }
    }

    private static bool IsRootCollection(string key) =>
        key.Equals("slides", StringComparison.OrdinalIgnoreCase)
        || key.Equals("sections", StringComparison.OrdinalIgnoreCase)
        || key.Equals("worksheets", StringComparison.OrdinalIgnoreCase);

    private static string? Key(JsonObject obj, string name)
    {
        var keys = obj.Select(p => p.Key).Where(k => k.Equals(name, StringComparison.OrdinalIgnoreCase)).ToArray();
        if (keys.Length > 1) throw new JsonException($"Ambiguous duplicate property '{name}'.");
        return keys.SingleOrDefault();
    }

    private static JsonNode? Get(JsonObject obj, string name) => Key(obj, name) is { } key ? obj[key] : null;
    private static string? Id(JsonObject obj) => Get(obj, "id")?.GetValue<string>();
}
