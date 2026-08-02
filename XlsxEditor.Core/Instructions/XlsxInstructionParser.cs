using System.Text.Json;
using System.Text.Json.Serialization;
using XlsxEditor.Core.Exceptions;

namespace XlsxEditor.Core.Instructions;

/// <summary>
/// Parses JSON into an XlsxInstructionSet with loud validation.
/// </summary>
public static class XlsxInstructionParser
{
    /// <summary>
    /// Unknown properties anywhere in the instruction set are a typo or a future
    /// vocabulary leaking in early; both must fail loudly instead of being silently
    /// dropped (the pre-hardening behavior).
    /// </summary>
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow
    };

    public static XlsxInstructionSet Parse(string json)
    {
        if (string.IsNullOrWhiteSpace(json))
        {
            throw new XlsxException("Instruction JSON must not be null or empty.");
        }

        XlsxInstructionSet instructionSet;
        try
        {
            instructionSet = JsonSerializer.Deserialize<XlsxInstructionSet>(json, JsonOptions)
                ?? throw new XlsxException("Instruction JSON deserialized to null. Check that the document is valid JSON and starts with a JSON object.");
        }
        catch (JsonException ex)
        {
            throw new XlsxException($"Invalid JSON in instruction set: {ex.Message}", ex);
        }

        XlsxInstructionValidator.Validate(instructionSet);
        return instructionSet;
    }

    public static XlsxInstructionSet ParseFromFile(string filePath)
    {
        if (!File.Exists(filePath))
        {
            throw new XlsxException($"Instruction file not found: '{filePath}'.");
        }

        var json = File.ReadAllText(filePath);
        return Parse(json);
    }
}
