using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using XlsxEditor.Core.Builders;
using XlsxEditor.Core.Exceptions;

namespace XlsxEditor.Core.Instructions;

/// <summary>
/// The outcome of a generation: the produced workbook bytes (null when the input was
/// rejected) plus the canonical validation/planning diagnostics collected along the way.
/// Warnings do not block generation; any error does.
/// </summary>
public sealed record XlsxGenerateResult(byte[]? Bytes, XlsxValidationResult Validation)
{
    public bool IsValid => Validation.IsValid;

    public bool HasErrors => Validation.HasErrors;

    public bool HasWarnings => Validation.HasWarnings;
}

/// <summary>
/// Generation switches. <see cref="ValidatePackage"/> is off by default because package
/// validation is a test-time/QA concern, not a production hot-path cost.
/// </summary>
public sealed record XlsxGenerateOptions
{
    /// <summary>Runs the OpenXML SDK validator over the finished package and reports findings.</summary>
    public bool ValidatePackage { get; init; }
}

/// <summary>
/// High-level JSON/model → XLSX generator. Parses and plans a JSON instruction set (or
/// an already-built <see cref="XlsxInstructionSet"/>), then produces workbook bytes via
/// the shared <see cref="XlsxInstructionExecutor"/>.
///
/// Input streams are never disposed and are left at their end position after reading;
/// output streams are never disposed, are flushed, and are left with their position at
/// the end of the written bytes. File outputs are written atomically: the bytes land in a
/// sibling temp file and are moved over the destination only on success, so a failure
/// never leaves a partial or corrupt workbook behind (and an existing file is preserved
/// when generation fails).
/// </summary>
public static class XlsxGenerator
{
    // ─── Generate to bytes (result carries the bytes + diagnostics) ──

    /// <summary>Generates a workbook from a JSON instruction string.</summary>
    public static XlsxGenerateResult Generate(string json, XlsxGenerateOptions? options = null)
    {
        var parsed = XlsxInstructionParser.ParseAndValidate(json);
        if (!parsed.IsValid)
        {
            return new XlsxGenerateResult(null, parsed.Validation);
        }

        return Generate(parsed.InstructionSet!, options);
    }

    /// <summary>
    /// Generates a workbook from a JSON instruction stream. The stream is read to the end
    /// and left open (position at the end); it is never disposed.
    /// </summary>
    public static XlsxGenerateResult Generate(Stream jsonStream, XlsxGenerateOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(jsonStream);
        return Generate(ReadAllText(jsonStream), options);
    }

    /// <summary>Generates a workbook from a JSON instruction file.</summary>
    public static XlsxGenerateResult GenerateFromFile(string jsonPath, XlsxGenerateOptions? options = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(jsonPath);
        return Generate(File.ReadAllText(jsonPath), options);
    }

    /// <summary>Generates a workbook from an already-deserialized instruction set.</summary>
    public static XlsxGenerateResult Generate(XlsxInstructionSet instructions, XlsxGenerateOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(instructions);

        var plan = XlsxPlanner.Plan(instructions);
        if (!plan.IsValid)
        {
            return new XlsxGenerateResult(null, plan.Validation);
        }

        var generateOptions = options ?? new XlsxGenerateOptions();
        byte[] bytes;
        try
        {
            bytes = BuildBytes(plan.Plan!);
        }
        catch (XlsxException ex)
        {
            // Post-planning execution failure (e.g. an unsupported style aspect): fold it
            // into the canonical diagnostics instead of losing it as a bare exception.
            return new XlsxGenerateResult(null, WithError(plan.Validation, ex.Message));
        }

        if (generateOptions.ValidatePackage)
        {
            var validated = ValidatePackage(plan.Validation, bytes);
            return new XlsxGenerateResult(validated.HasErrors ? null : bytes, validated);
        }

        return new XlsxGenerateResult(bytes, plan.Validation);
    }

    // ─── Outputs ───────────────────────────────────────────────────

    /// <summary>
    /// Generates a workbook from JSON and writes it to <paramref name="outputPath"/>
    /// atomically (temp file + move). When generation fails nothing is written and the
    /// destination, if any, is left untouched; the returned result carries the diagnostics.
    /// </summary>
    public static XlsxGenerateResult GenerateToFile(string json, string outputPath, XlsxGenerateOptions? options = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        var result = Generate(json, options);
        WriteAtomically(result, outputPath);
        return result;
    }

    /// <summary>Generates a workbook from an instruction set and writes it atomically.</summary>
    public static XlsxGenerateResult GenerateToFile(XlsxInstructionSet instructions, string outputPath, XlsxGenerateOptions? options = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        var result = Generate(instructions, options);
        WriteAtomically(result, outputPath);
        return result;
    }

    /// <summary>
    /// Generates a workbook from JSON and writes it to <paramref name="outputStream"/>,
    /// flushing and leaving the stream open (position at the end of the written bytes).
    /// Nothing is written when generation fails.
    /// </summary>
    public static XlsxGenerateResult GenerateToStream(string json, Stream outputStream, XlsxGenerateOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(outputStream);
        var result = Generate(json, options);
        WriteToStream(result, outputStream);
        return result;
    }

    /// <summary>Generates a workbook from an instruction set and writes it to a stream.</summary>
    public static XlsxGenerateResult GenerateToStream(XlsxInstructionSet instructions, Stream outputStream, XlsxGenerateOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(outputStream);
        var result = Generate(instructions, options);
        WriteToStream(result, outputStream);
        return result;
    }

    // ─── Throwing convenience ──────────────────────────────────────

    /// <summary>Generates a workbook from JSON and returns the bytes, throwing the first error.</summary>
    public static byte[] GenerateBytes(string json, XlsxGenerateOptions? options = null)
    {
        return BytesOrThrow(Generate(json, options));
    }

    /// <summary>Generates a workbook from an instruction set and returns the bytes, throwing the first error.</summary>
    public static byte[] GenerateBytes(XlsxInstructionSet instructions, XlsxGenerateOptions? options = null)
    {
        return BytesOrThrow(Generate(instructions, options));
    }

    // ─── Plumbing ──────────────────────────────────────────────────

    private static byte[] BuildBytes(XlsxPlan plan)
    {
        using var builder = (WorkbookBuilder)WorkbookBuilder.Create();
        XlsxInstructionExecutor.Execute(plan, builder);
        return builder.SaveToBytes();
    }

    private static byte[] BytesOrThrow(XlsxGenerateResult result)
    {
        if (!result.IsValid)
        {
            var first = result.Validation.Errors.First();
            throw new XlsxException(
                string.IsNullOrEmpty(first.Path) ? first.Message : $"{first.Message} (at {first.Path})");
        }

        return result.Bytes!;
    }

    private static void WriteAtomically(XlsxGenerateResult result, string outputPath)
    {
        if (result.Bytes is null)
        {
            return;
        }

        var fullPath = Path.GetFullPath(outputPath);
        var directory = Path.GetDirectoryName(fullPath);
        if (string.IsNullOrEmpty(directory))
        {
            throw new XlsxException($"Cannot resolve the directory of output path '{outputPath}'.");
        }

        var tempPath = Path.Combine(directory, $".{Path.GetFileName(fullPath)}.{Guid.NewGuid():N}.tmp");
        try
        {
            File.WriteAllBytes(tempPath, result.Bytes);
            File.Move(tempPath, fullPath, overwrite: true);
        }
        catch
        {
            try
            {
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
            }
            catch
            {
                // Best-effort cleanup; the original failure propagates.
            }

            throw;
        }
    }

    private static void WriteToStream(XlsxGenerateResult result, Stream outputStream)
    {
        if (result.Bytes is null)
        {
            return;
        }

        outputStream.Write(result.Bytes, 0, result.Bytes.Length);
        outputStream.Flush();
    }

    private static XlsxValidationResult ValidatePackage(XlsxValidationResult validation, byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = SpreadsheetDocument.Open(stream, false);
        var errors = new OpenXmlValidator().Validate(document).ToList();
        if (errors.Count == 0)
        {
            return validation;
        }

        var diagnostics = validation.Diagnostics.ToList();
        foreach (var error in errors)
        {
            diagnostics.Add(new XlsxDiagnostic(
                XlsxDiagnosticCode.PackageInvalid,
                XlsxDiagnosticSeverity.Error,
                error.Path?.XPath ?? string.Empty,
                error.Description));
        }

        return new XlsxValidationResult { Diagnostics = diagnostics };
    }

    private static XlsxValidationResult WithError(XlsxValidationResult validation, string message)
    {
        var diagnostics = validation.Diagnostics.ToList();
        diagnostics.Add(new XlsxDiagnostic(
            XlsxDiagnosticCode.GenerationFailed,
            XlsxDiagnosticSeverity.Error,
            string.Empty,
            message));
        return new XlsxValidationResult { Diagnostics = diagnostics };
    }

    private static string ReadAllText(Stream stream)
    {
        using var reader = new StreamReader(
            stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, leaveOpen: true);
        return reader.ReadToEnd();
    }
}
