using System.Text.Json;
using System.Text.Json.Serialization;
using XlsxEditor.Core.Exceptions;

namespace XlsxEditor.Core.Instructions;

/// <summary>
/// The result of parsing without throwing: the deserialized instruction set (null when
/// the JSON could not be deserialized) plus the structured validation outcome.
/// </summary>
public sealed record XlsxParseResult(XlsxInstructionSet? InstructionSet, XlsxValidationResult Validation)
{
    public bool IsValid => Validation.IsValid;
}

/// <summary>
/// Parses JSON into an <see cref="XlsxInstructionSet"/>. The throwing <see cref="Parse"/>
/// / <see cref="ParseFromFile"/> entry points are adapters over
/// <see cref="ParseAndValidate"/>, which never throws for semantic issues and returns the
/// full structured validation result instead.
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
        var parsed = ParseAndValidate(json);
        if (!parsed.IsValid)
        {
            throw new XlsxException(parsed.Validation.Errors.First().Message);
        }

        return parsed.InstructionSet!;
    }

    /// <summary>
    /// Deserializes and validates without throwing for semantic problems. JSON syntax
    /// and shape errors (unmapped members, wrong property kinds) are folded into the
    /// returned <see cref="XlsxValidationResult"/> as an <see cref="XlsxDiagnosticCode.InvalidJson"/>
    /// error so callers get one uniform, path-qualified diagnostic stream.
    /// </summary>
    public static XlsxParseResult ParseAndValidate(string json)
    {
        if (string.IsNullOrWhiteSpace(json))
        {
            return Invalid(
                XlsxDiagnosticCode.InvalidJson,
                "Instruction JSON must not be null or empty.");
        }

        XlsxInstructionSet? instructionSet = null;
        string? jsonError = null;
        try
        {
            instructionSet = JsonSerializer.Deserialize<XlsxInstructionSet>(json, JsonOptions);
        }
        catch (JsonException ex)
        {
            jsonError = ex.Message;
        }

        if (instructionSet is null && jsonError is null)
        {
            jsonError = "the document is valid JSON but did not deserialize to an instruction " +
                        "set (expected a JSON object with 'version' and 'worksheets').";
        }

        if (jsonError is not null)
        {
            return Invalid(
                XlsxDiagnosticCode.InvalidJson,
                $"Invalid JSON in instruction set: {jsonError}");
        }

        var validation = XlsxValidationEngine.Validate(instructionSet);
        return new XlsxParseResult(instructionSet, validation);
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

    private static XlsxParseResult Invalid(XlsxDiagnosticCode code, string message) =>
        new(
            null,
            new XlsxValidationResult
            {
                Diagnostics = new[]
                {
                    new XlsxDiagnostic(code, XlsxDiagnosticSeverity.Error, string.Empty, message)
                }
            });
}
