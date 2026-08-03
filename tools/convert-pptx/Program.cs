using System.Globalization;
using OfficeEditor.Core.Rendering;

const string Usage =
    "Usage: convert-pptx <input.pptx> <output-path> [--format pdf|png|svg|typ] [--font-path <path>] [--ppi <n>]\n" +
    "  --format: pdf (single file, default), png/svg (directory of page-NNN.ext), typ (Typst source).\n" +
    "  --font-path: extra font file/directory (path-separator-joined list allowed).\n" +
    "               Defaults to the system font directories; embedded PPTX fonts always win.\n" +
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

Console.WriteLine($"Converting {inputPath}...");

var renderer = new DocumentRenderer();
var result = renderer.Render(new DocumentRenderRequest
{
    SourcePath = inputPath,
    Format = format,
    Ppi = ppi,
    // Mirror the tool's historical default: when no explicit font path is given, hand the
    // Typst compiler the system font directories (the embedded deck fonts always come first).
    FontPath = fontPath ?? DefaultSystemFontPath()
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
else if (format == DocumentOutputFormat.Typ)
{
    if (result.Pages.Length == 0)
    {
        Console.WriteLine("Error: Render produced no Typst source.");
        return 1;
    }

    EnsureDirectory(outputPath);
    File.WriteAllBytes(outputPath, result.Pages[0]);
    Console.WriteLine($"Typst source saved to: {outputPath}");
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

static void EnsureDirectory(string path)
{
    var directory = Path.GetDirectoryName(path);
    if (!string.IsNullOrEmpty(directory))
    {
        Directory.CreateDirectory(directory);
    }
}

static string? DefaultSystemFontPath()
{
    string[] candidates =
    [
        "/System/Library/Fonts",
        "/Library/Fonts",
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Library", "Fonts"),
        "/usr/share/fonts",
        "/usr/local/share/fonts",
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".fonts"),
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".local", "share", "fonts"),
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts"),
    ];

    var existing = candidates.Where(Directory.Exists).ToList();
    return existing.Count > 0 ? string.Join(Path.PathSeparator, existing) : null;
}
