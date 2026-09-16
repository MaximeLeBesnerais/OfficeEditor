using System.Text.Json;
using System.Text.Json.Nodes;

namespace OfficeEditor.Core.Generation;

/// <summary>Stateless source editing. Returned JSON must be saved and validated by the format generator.</summary>
public static class GenerationJsonEditor
{
    /// <summary>
    /// Applies a sequential batch of set/rename operations to a private source copy.
    /// A failed operation returns no partial result. An empty batch normalizes IDs.
    /// This validates identity and operation syntax; format vocabulary is validated at generation.
    /// </summary>
    public static string Apply(string json, string operationsJson)
    {
        using var document = JsonDocument.Parse(operationsJson);
        if (document.RootElement.ValueKind != JsonValueKind.Array)
            throw new JsonException("operations must be an array.");
        using var source = JsonDocument.Parse(json);
        RejectDuplicateProperties(source.RootElement);
        var result = GenerationJsonIds.Normalize(json).Json;
        int index = 0;
        foreach (var operation in document.RootElement.EnumerateArray())
        {
            try
            {
                if (operation.ValueKind != JsonValueKind.Object)
                    throw new JsonException("Expected an operation object.");
                var type = RequiredString(operation, "type");
                var id = RequiredString(operation, "id");
                string[] allowed = type switch
                {
                    "set" => ["type", "id", "properties"],
                    "rename" => ["type", "id", "newId"],
                    _ => throw new JsonException("type must be set or rename.")
                };
                var seen = new HashSet<string>(StringComparer.Ordinal);
                foreach (var property in operation.EnumerateObject())
                    if (!allowed.Contains(property.Name) || !seen.Add(property.Name))
                        throw new JsonException($"Unknown or duplicate property '{property.Name}'.");
                if (type == "rename")
                    result = GenerationJsonIds.Rename(result, id, RequiredString(operation, "newId"));
                else
                {
                    if (!operation.TryGetProperty("properties", out var properties) || properties.ValueKind != JsonValueKind.Object)
                        throw new JsonException("properties must be an object.");
                    RejectDuplicateProperties(properties);
                    result = GenerationJsonIds.SetProperties(result, id, JsonNode.Parse(properties.GetRawText())!.AsObject());
                }
            }
            catch (Exception ex) when (ex is JsonException or KeyNotFoundException or ArgumentException)
            {
                throw new JsonException($"operations[{index}]: {ex.Message}", ex);
            }
            index++;
        }
        return result;
    }

    private static void RejectDuplicateProperties(JsonElement element)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var property in element.EnumerateObject())
            {
                if (!names.Add(property.Name)) throw new JsonException($"Duplicate source property '{property.Name}'.");
                RejectDuplicateProperties(property.Value);
            }
        }
        else if (element.ValueKind == JsonValueKind.Array)
            foreach (var child in element.EnumerateArray()) RejectDuplicateProperties(child);
    }

    private static string RequiredString(JsonElement operation, string name)
    {
        if (!operation.TryGetProperty(name, out var value) || value.ValueKind != JsonValueKind.String
            || string.IsNullOrWhiteSpace(value.GetString()))
            throw new JsonException($"{name} must be a non-empty string.");
        return value.GetString()!;
    }
}
