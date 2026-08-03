using System.Text.Json;
using OfficeEditor.Core.Services;
using XlsxEditor.Core.Exceptions;
using XlsxEditor.Core.Instructions;
using XlsxEditor.Core.Rendering;
using XlsxEditor.Core.Rendering.Emit;
using XlsxEditor.Core.Rendering.Models;
using XlsxEditor.Core.Rendering.Read;

const string Usage =
    "Usage: convert-xlsx <input.xlsx|input.json> <output-path> [--format pdf|png|svg|typ|json] [--font-path <path>] [--ppi <n>]";

if (args.Length < 2)
{
    Console.WriteLine(Usage);
    return 1;
}

var inputPath = args[0];
var outputPath = args[1];
var format = XlsxOutputFormat.Pdf;
string? fontPath = null;
float ppi = 150;

for (var i = 2; i < args.Length; i++)
{
    switch (args[i])
    {
        case "--format":
            if (i + 1 >= args.Length)
            {
                Console.WriteLine("Error: Missing value for --format.");
                Console.WriteLine(Usage);
                return 1;
            }

            var formatValue = args[++i];
            if (!TryParseFormat(formatValue, out format))
            {
                Console.WriteLine($"Error: Unsupported format '{formatValue}'. Expected 'pdf', 'png', 'svg', 'typ', or 'json'.");
                return 1;
            }

            break;

        case "--font-path":
            if (i + 1 >= args.Length)
            {
                Console.WriteLine("Error: Missing value for --font-path.");
                Console.WriteLine(Usage);
                return 1;
            }

            fontPath = args[++i];
            break;

        case "--ppi":
            if (i + 1 >= args.Length || !float.TryParse(args[i + 1], System.Globalization.CultureInfo.InvariantCulture, out ppi))
            {
                Console.WriteLine("Error: Missing or invalid value for --ppi.");
                Console.WriteLine(Usage);
                return 1;
            }

            i++;
            break;

        default:
            Console.WriteLine($"Error: Unknown argument '{args[i]}'.");
            Console.WriteLine(Usage);
            return 1;
    }
}

if (!File.Exists(inputPath))
{
    Console.WriteLine($"Error: File not found: {inputPath}");
    return 1;
}

try
{
    var isJson = inputPath.EndsWith(".json", StringComparison.OrdinalIgnoreCase);

    if (format == XlsxOutputFormat.Json)
    {
        // .xlsx -> JSON instruction set (best-effort serialization).
        var json = XlsxJsonSerializer.Serialize(inputPath);
        EnsureDirectory(outputPath);
        File.WriteAllText(outputPath, JsonSerializer.Serialize(json, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine($"JSON saved to: {outputPath}");
        return 0;
    }

    if (format == XlsxOutputFormat.Typ)
    {
        var workbook = LoadWorkbook(inputPath, isJson);
        var source = new XlsxToTypstConverter(workbook).GenerateTypstSource();
        EnsureDirectory(outputPath);
        File.WriteAllText(outputPath, source);
        Console.WriteLine($"Typst source saved to: {outputPath}");
        return 0;
    }

    Console.WriteLine($"Converting {inputPath}...");
    var compileOptions = new CompileOptions
    {
        Format = format switch
        {
            XlsxOutputFormat.Png => OutputFormat.Png,
            XlsxOutputFormat.Svg => OutputFormat.Svg,
            _ => OutputFormat.Pdf
        },
        Ppi = ppi,
        FontDirectory = fontPath
    };

    XlsxRenderResult result;
    if (isJson)
    {
        result = XlsxRenderer.RenderJsonFile(inputPath, compileOptions);
    }
    else
    {
        result = XlsxRenderer.RenderFile(inputPath, compileOptions);
    }

    foreach (var issue in result.ReadIssues)
    {
        Console.WriteLine($"{issue.Severity}: {issue.Message}" + (issue.Location == null ? "" : $" ({issue.Location})"));
    }

    if (!result.Success)
    {
        Console.WriteLine($"Error: Typst compilation failed. {result.Compile.ErrorMessage}");
        return 1;
    }

    EnsureDirectory(outputPath);

    if (format == XlsxOutputFormat.Pdf)
    {
        if (result.Compile.Pages.Length == 0)
        {
            Console.WriteLine("Error: Typst compilation produced no PDF output.");
            return 1;
        }

        File.WriteAllBytes(outputPath, result.Compile.Pages[0]);
        Console.WriteLine($"PDF saved to: {outputPath}");
    }
    else
    {
        // PNG/SVG produce one buffer per page; write into a directory named by the output path.
        Directory.CreateDirectory(outputPath);
        for (int i = 0; i < result.Compile.Pages.Length; i++)
        {
            string ext = format == XlsxOutputFormat.Png ? "png" : "svg";
            string pagePath = Path.Combine(outputPath, $"page-{i + 1:D3}.{ext}");
            File.WriteAllBytes(pagePath, result.Compile.Pages[i]);
        }

        Console.WriteLine($"{result.Compile.Pages.Length} page(s) saved to: {outputPath}");
    }

    return 0;
}
catch (XlsxException ex)
{
    Console.WriteLine($"Error: {ex.Message}");
    return 1;
}
catch (Exception ex)
{
    Console.WriteLine($"Error: {ex.Message}");
    return 1;
}

static XlsxRenderWorkbook LoadWorkbook(string inputPath, bool isJson)
{
    if (isJson)
    {
        var generated = XlsxGenerator.GenerateFromFile(inputPath);
        if (!generated.IsValid || generated.Bytes is null)
        {
            var errors = string.Join("; ", generated.Validation.Errors.Select(e => e.Message));
            throw new XlsxException($"Invalid XLSX JSON: {errors}");
        }

        return new XlsxReader().Read(generated.Bytes).Workbook;
    }

    return new XlsxReader().Read(inputPath).Workbook;
}

static bool TryParseFormat(string value, out XlsxOutputFormat format)
{
    switch (value.ToLowerInvariant())
    {
        case "pdf": format = XlsxOutputFormat.Pdf; return true;
        case "png": format = XlsxOutputFormat.Png; return true;
        case "svg": format = XlsxOutputFormat.Svg; return true;
        case "typ": format = XlsxOutputFormat.Typ; return true;
        case "json": format = XlsxOutputFormat.Json; return true;
        default: format = default; return false;
    }
}

static void EnsureDirectory(string path)
{
    var directory = Path.GetDirectoryName(path);
    if (!string.IsNullOrEmpty(directory))
    {
        Directory.CreateDirectory(directory);
    }
}

internal enum XlsxOutputFormat
{
    Pdf,
    Png,
    Svg,
    Typ,
    Json
}
