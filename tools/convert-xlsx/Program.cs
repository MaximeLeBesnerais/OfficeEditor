using System.Globalization;
using System.Text.Json;
using OfficeEditor.Core.Rendering;
using XlsxEditor.Core.Exceptions;
using XlsxEditor.Core.Rendering;

const string Usage =
    "Usage: convert-xlsx <input.xlsx|input.json> <output-path> [--format pdf|png|svg|typ|json] [--font-path <path>] [--ppi <n>]\n" +
    "  --format: pdf (single file, default), png/svg (directory of page-NNN.ext), typ (Typst source), json (XLSX instructions).\n" +
    "  --font-path: extra font file/directory (path-separator-joined list allowed).\n" +
    "  --ppi: raster density for png/svg output in pixels per inch (default 150).";

if (args.Length < 2)
{
    Console.WriteLine(Usage);
    return 1;
}

var inputPath = args[0];
var outputPath = args[1];
var format = XlsxOutputFormat.Pdf;
string? fontPath = null;
var ppi = 150f;

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
            if (i + 1 >= args.Length || !float.TryParse(args[i + 1], CultureInfo.InvariantCulture, out ppi))
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

if (format == XlsxOutputFormat.Json)
{
    return ConvertToJson(inputPath, outputPath);
}

Console.WriteLine($"Converting {inputPath}...");

// The facade dispatches by source extension and routes .json inputs through the XLSX
// generation vocabulary, so a single code path serves both .xlsx and .json sources.
var renderer = new DocumentRenderer();
var result = renderer.Render(new DocumentRenderRequest
{
    SourcePath = inputPath,
    Format = format switch
    {
        XlsxOutputFormat.Png => DocumentOutputFormat.Png,
        XlsxOutputFormat.Svg => DocumentOutputFormat.Svg,
        XlsxOutputFormat.Typ => DocumentOutputFormat.Typ,
        _ => DocumentOutputFormat.Pdf
    },
    Ppi = ppi,
    FontPath = fontPath
});

if (!result.Success)
{
    Console.WriteLine($"Error: {result.ErrorMessage}");
    return 1;
}

foreach (var warning in result.Warnings)
{
    Console.WriteLine($"Warning: {warning}");
}

if (format is XlsxOutputFormat.Pdf or XlsxOutputFormat.Typ)
{
    if (result.Pages.Length == 0)
    {
        Console.WriteLine("Error: Render produced no output.");
        return 1;
    }

    EnsureDirectory(outputPath);
    File.WriteAllBytes(outputPath, result.Pages[0]);
    var label = format == XlsxOutputFormat.Pdf ? "PDF" : "Typst source";
    Console.WriteLine($"{label} saved to: {outputPath}");
}
else
{
    // PNG/SVG produce one buffer per page; write into a directory named by the output path.
    Directory.CreateDirectory(outputPath);
    var extension = format == XlsxOutputFormat.Png ? "png" : "svg";
    for (var i = 0; i < result.Pages.Length; i++)
    {
        var pagePath = Path.Combine(outputPath, $"page-{i + 1:D3}.{extension}");
        File.WriteAllBytes(pagePath, result.Pages[i]);
    }

    Console.WriteLine($"{result.Pages.Length} page(s) saved to: {outputPath}");
}

return 0;

static int ConvertToJson(string inputPath, string outputPath)
{
    try
    {
        // .xlsx -> JSON instruction set (best-effort serialization).
        var json = XlsxJsonSerializer.Serialize(inputPath);
        EnsureDirectory(outputPath);
        File.WriteAllText(outputPath, JsonSerializer.Serialize(json, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine($"JSON saved to: {outputPath}");
        return 0;
    }
    catch (XlsxException ex)
    {
        Console.WriteLine($"Error: {ex.Message}");
        return 1;
    }
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
