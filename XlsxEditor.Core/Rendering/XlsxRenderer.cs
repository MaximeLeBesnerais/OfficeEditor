using OfficeEditor.Core.Services;
using XlsxEditor.Core.Exceptions;
using XlsxEditor.Core.Instructions;
using XlsxEditor.Core.Rendering.Emit;
using XlsxEditor.Core.Rendering.Models;
using XlsxEditor.Core.Rendering.Read;

namespace XlsxEditor.Core.Rendering;

/// <summary>
/// High-level facade that renders spreadsheets to PDF / PNG / SVG using the repository's
/// own Typst pipeline. Inputs can be an arbitrary .xlsx (file/stream/bytes), an immutable
/// <see cref="XlsxRenderWorkbook"/>, or JSON instruction sets (which are first generated
/// to an in-memory workbook, then read back). No external office suite is involved.
/// </summary>
public static class XlsxRenderer
{
    /// <summary>Renders a parsed workbook model to the requested output format.</summary>
    public static CompileResult Render(XlsxRenderWorkbook workbook, CompileOptions? options = null)
    {
        if (workbook == null)
            throw new ArgumentNullException(nameof(workbook));

        var source = new XlsxToTypstConverter(workbook).GenerateTypstSource();
        using var compiler = new TypstCompilerService();
        return compiler.Compile(source, options ?? new CompileOptions());
    }

    /// <summary>Renders an arbitrary .xlsx read from a path.</summary>
    public static XlsxRenderResult RenderFile(string xlsxPath, CompileOptions? options = null)
    {
        var read = new XlsxReader().Read(xlsxPath);
        var compile = Render(read.Workbook, options);
        return new XlsxRenderResult(compile, read.Issues);
    }

    /// <summary>Renders an arbitrary .xlsx read from a stream.</summary>
    public static XlsxRenderResult RenderStream(Stream xlsxStream, CompileOptions? options = null)
    {
        var read = new XlsxReader().Read(xlsxStream);
        var compile = Render(read.Workbook, options);
        return new XlsxRenderResult(compile, read.Issues);
    }

    /// <summary>Renders an arbitrary .xlsx read from a byte buffer.</summary>
    public static XlsxRenderResult RenderBytes(byte[] xlsxBytes, CompileOptions? options = null)
    {
        var read = new XlsxReader().Read(xlsxBytes);
        var compile = Render(read.Workbook, options);
        return new XlsxRenderResult(compile, read.Issues);
    }

    /// <summary>
    /// Renders a JSON instruction set directly to the requested output format without
    /// producing an intermediate .xlsx file on disk. Validation failures abort rendering.
    /// </summary>
    public static XlsxRenderResult RenderJson(string json, CompileOptions? options = null)
    {
        var generated = XlsxGenerator.Generate(json);
        if (!generated.IsValid || generated.Bytes is null)
            throw new XlsxException($"Cannot render invalid XLSX JSON: {string.Join("; ", generated.Validation.Errors.Select(e => e.Message))}");

        return RenderBytes(generated.Bytes, options);
    }

    /// <summary>Renders a JSON instruction set loaded from a file path.</summary>
    public static XlsxRenderResult RenderJsonFile(string jsonPath, CompileOptions? options = null)
    {
        var generated = XlsxGenerator.GenerateFromFile(jsonPath);
        if (!generated.IsValid || generated.Bytes is null)
            throw new XlsxException($"Cannot render invalid XLSX JSON: {string.Join("; ", generated.Validation.Errors.Select(e => e.Message))}");

        return RenderBytes(generated.Bytes, options);
    }

    /// <summary>Renders a JSON instruction set to the requested output format.</summary>
    public static XlsxRenderResult RenderInstructions(XlsxInstructionSet instructions, CompileOptions? options = null)
    {
        var generated = XlsxGenerator.Generate(instructions);
        if (!generated.IsValid || generated.Bytes is null)
            throw new XlsxException($"Cannot render invalid XLSX instructions: {string.Join("; ", generated.Validation.Errors.Select(e => e.Message))}");

        return RenderBytes(generated.Bytes, options);
    }
}

/// <summary>Combines a Typst compile result with the reader issues encountered while loading.</summary>
public sealed record XlsxRenderResult(CompileResult Compile, IReadOnlyList<XlsxReadIssue> ReadIssues)
{
    public bool Success => Compile.Success;
}
