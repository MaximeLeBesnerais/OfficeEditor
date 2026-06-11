using DocumentFormat.OpenXml.Packaging;
using DocxEditor.Core.Converters;
using OfficeEditor.Core.Services;

const string Usage = "Usage: convert-docx <input.docx> <output-path> [--format pdf|typ] [--font-path <path>]";

if (args.Length < 2)
{
    Console.WriteLine(Usage);
    return 1;
}

var inputPath = args[0];
var outputPath = args[1];
var format = DocxOutputFormat.Pdf;
string? fontPath = null;

for (var i = 2; i < args.Length; i++)
{
    if (args[i] == "--format")
    {
        if (i + 1 >= args.Length)
        {
            Console.WriteLine("Error: Missing value for --format.");
            Console.WriteLine(Usage);
            return 1;
        }

        var formatValue = args[++i];
        if (formatValue.Equals("pdf", StringComparison.OrdinalIgnoreCase))
        {
            format = DocxOutputFormat.Pdf;
        }
        else if (formatValue.Equals("typ", StringComparison.OrdinalIgnoreCase))
        {
            format = DocxOutputFormat.Typ;
        }
        else
        {
            Console.WriteLine($"Error: Unsupported format '{formatValue}'. Expected 'pdf' or 'typ'.");
            return 1;
        }
    }
    else if (args[i] == "--font-path")
    {
        if (i + 1 >= args.Length)
        {
            Console.WriteLine("Error: Missing value for --font-path.");
            Console.WriteLine(Usage);
            return 1;
        }

        fontPath = args[++i];
    }
    else
    {
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
    Console.WriteLine($"Converting {inputPath}...");

    using var doc = WordprocessingDocument.Open(inputPath, false);
    using var converter = new DocxToTypstConverter(doc);
    var document = converter.Convert();
    var typstSource = converter.GenerateTypstSource(document);
    foreach (var diagnostic in document.Diagnostics)
    {
        Console.WriteLine($"Warning: {diagnostic}");
    }

    var sourceOnly = format == DocxOutputFormat.Typ;
    var typstPath = sourceOnly ? outputPath : Path.ChangeExtension(outputPath, ".typ");
    EnsureDirectory(typstPath);
    File.WriteAllText(typstPath, typstSource);
    CopyAssets(document.AssetsDirectory, typstPath);
    Console.WriteLine($"Typst source saved to: {typstPath}");

    if (sourceOnly)
    {
        return 0;
    }

    Console.WriteLine("Compiling with Typst...");
    using var compiler = new TypstCompilerService();
    var options = new CompileOptions
    {
        Format = OutputFormat.Pdf,
        WorkingDirectory = string.IsNullOrWhiteSpace(document.TempDirectory) ? null : document.TempDirectory,
        FontDirectory = fontPath
    };

    var result = compiler.Compile(typstSource, options);
    if (!result.Success)
    {
        Console.WriteLine($"Error: Typst compilation failed. {result.ErrorMessage}");
        return 1;
    }

    if (result.Pages.Length == 0)
    {
        Console.WriteLine("Error: Typst compilation succeeded but produced no PDF output.");
        return 1;
    }

    EnsureDirectory(outputPath);
    File.WriteAllBytes(outputPath, result.Pages[0]);
    Console.WriteLine($"PDF saved to: {outputPath}");
    return 0;
}
catch (Exception ex)
{
    Console.WriteLine($"Error: {ex.Message}");
    return 1;
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
    Directory.CreateDirectory(targetAssetsDirectory);

    foreach (var sourcePath in Directory.EnumerateFiles(sourceAssetsDirectory))
    {
        var targetPath = Path.Combine(targetAssetsDirectory, Path.GetFileName(sourcePath));
        File.Copy(sourcePath, targetPath, overwrite: true);
    }
}

internal enum DocxOutputFormat
{
    Pdf,
    Typ
}
