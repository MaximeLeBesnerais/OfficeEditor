using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocxEditor.Core.Converters;
using OfficeEditor.Core.Rendering;
using OfficeEditor.Core.Services;

const string Usage =
    "Usage: convert-docx <input.docx> <output-path> [--format pdf|png|svg|typ] [--font-path <path>] [--ppi <n>]\n" +
    "  --format: pdf (single file, default), png/svg (directory of page-NNN.ext), typ (Typst source + assets).\n" +
    "  --font-path: extra font file/directory (path-separator-joined list allowed).\n" +
    "  --ppi: raster density for png/svg output in pixels per inch (default 150).";

if (args.Length < 2)
{
    Console.WriteLine(Usage);
    return 1;
}

var inputPath = args[0];
var outputPath = args[1];
var format = DocumentOutputFormat.Pdf;
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
                Console.WriteLine($"Error: Unsupported format '{formatValue}'. Expected 'pdf', 'png', 'svg', or 'typ'.");
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

// The Typst-source export keeps the document's extracted assets (images, fonts) alongside
// the source file so the output is standalone-compilable. The shared facade returns only the
// source bytes, so this format stays on the converter path; every compiled format goes
// through the facade.
if (format == DocumentOutputFormat.Typ)
{
    return ConvertToTypstSource(inputPath, outputPath);
}

Console.WriteLine($"Converting {inputPath}...");

var renderer = new DocumentRenderer();
var result = renderer.Render(new DocumentRenderRequest
{
    SourcePath = inputPath,
    Format = format,
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

if (format == DocumentOutputFormat.Pdf)
{
    if (result.Pages.Length == 0)
    {
        Console.WriteLine("Error: Render produced no PDF output.");
        return 1;
    }

    EnsureDirectory(outputPath);
    File.WriteAllBytes(outputPath, result.Pages[0]);
    Console.WriteLine($"PDF saved to: {outputPath}");
}
else
{
    // PNG/SVG produce one buffer per page; write into a directory named by the output path.
    Directory.CreateDirectory(outputPath);
    var extension = format == DocumentOutputFormat.Png ? "png" : "svg";
    for (var i = 0; i < result.Pages.Length; i++)
    {
        var pagePath = Path.Combine(outputPath, $"page-{i + 1:D3}.{extension}");
        File.WriteAllBytes(pagePath, result.Pages[i]);
    }

    Console.WriteLine($"{result.Pages.Length} page(s) saved to: {outputPath}");
}

return 0;

static int ConvertToTypstSource(string inputPath, string outputPath)
{
    string? tempDirectory = null;

    try
    {
        Console.WriteLine($"Converting {inputPath}...");

        using var doc = WordprocessingDocument.Open(inputPath, false);
        using var converter = new DocxToTypstConverter(doc);
        var document = converter.Convert();
        tempDirectory = document.TempDirectory;
        var typstSource = converter.GenerateTypstSource(document);
        foreach (var diagnostic in document.Diagnostics)
        {
            Console.WriteLine($"Warning: {diagnostic}");
        }

        EnsureDirectory(outputPath);
        File.WriteAllText(outputPath, typstSource);
        CopyAssets(document.AssetsDirectory, outputPath);
        Console.WriteLine($"Typst source saved to: {outputPath}");
        return 0;
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error: {ex.Message}");
        return 1;
    }
    finally
    {
        TryDeleteDirectory(tempDirectory);
    }
}

static bool TryParseFormat(string value, out DocumentOutputFormat format)
{
    switch (value.ToLowerInvariant())
    {
        case "pdf": format = DocumentOutputFormat.Pdf; return true;
        case "png": format = DocumentOutputFormat.Png; return true;
        case "svg": format = DocumentOutputFormat.Svg; return true;
        case "typ": format = DocumentOutputFormat.Typ; return true;
        default: format = default; return false;
    }
}

static void TryDeleteDirectory(string? path)
{
    if (string.IsNullOrWhiteSpace(path) || !Directory.Exists(path))
    {
        return;
    }

    try
    {
        Directory.Delete(path, recursive: true);
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Warning: Could not delete temp directory '{path}': {ex.Message}");
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

static void CopyAssets(string? sourceAssetsDirectory, string typstPath)
{
    if (string.IsNullOrWhiteSpace(sourceAssetsDirectory) || !Directory.Exists(sourceAssetsDirectory))
    {
        return;
    }

    var typstDirectory = Path.GetDirectoryName(typstPath);
    var targetAssetsDirectory = Path.Combine(string.IsNullOrEmpty(typstDirectory) ? "." : typstDirectory, "assets");

    if (Directory.Exists(targetAssetsDirectory))
    {
        try
        {
            Directory.Delete(targetAssetsDirectory, recursive: true);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Warning: Could not clear existing assets directory '{targetAssetsDirectory}': {ex.Message}");
        }
    }

    Directory.CreateDirectory(targetAssetsDirectory);

    foreach (var sourcePath in Directory.EnumerateFiles(sourceAssetsDirectory))
    {
        var targetPath = Path.Combine(targetAssetsDirectory, Path.GetFileName(sourcePath));
        File.Copy(sourcePath, targetPath, overwrite: true);
    }
}
