using DocumentFormat.OpenXml.Packaging;
using PptxEditor.Core.Converters;
using OfficeEditor.Core.Services;

if (args.Length < 2)
{
    Console.WriteLine("Usage: convert-pptx <input.pptx> <output-path> [--format pdf|png] [--font-path <path>]");
    Console.WriteLine("  --font-path: extra font file/directory (path-separator-joined list allowed).");
    Console.WriteLine("               Defaults to the system font directories; embedded PPTX fonts always win.");
    return 1;
}

var inputPath = args[0];
var outputPath = args[1];
var format = OutputFormat.Pdf;
string? fontPath = null;

for (var i = 2; i < args.Length; i++)
{
    if ((args[i] == "--format" || args[i] == "--foramt") && i + 1 < args.Length)
    {
        format = args[++i].Equals("png", StringComparison.OrdinalIgnoreCase)
            ? OutputFormat.Png
            : OutputFormat.Pdf;
    }
    else if (args[i] == "--font-path" && i + 1 < args.Length)
    {
        fontPath = args[++i];
    }
}

if (!File.Exists(inputPath))
{
    Console.WriteLine($"Error: File not found: {inputPath}");
    return 1;
}

Console.WriteLine($"Converting {inputPath}...");

using var doc = PresentationDocument.Open(inputPath, false);
using var converter = new PptxToTypstConverter(doc);
var presentation = converter.Convert();

// Generate Typst source
var typstSource = converter.GenerateTypstSource(presentation);

// Save Typst source for debugging
var typstPath = format == OutputFormat.Png
    ? Path.Combine(outputPath, Path.GetFileNameWithoutExtension(inputPath) + ".typ")
    : Path.ChangeExtension(outputPath, ".typ");
var typstDirectory = Path.GetDirectoryName(typstPath);
if (!string.IsNullOrEmpty(typstDirectory))
{
    Directory.CreateDirectory(typstDirectory);
}
File.WriteAllText(typstPath, typstSource);
Console.WriteLine($"Typst source saved to: {typstPath}");

// Compile with Typst
Console.WriteLine("Compiling with Typst...");
using var compiler = new TypstCompilerService();
var embeddedFontPath = presentation.FontFiles.Count > 0 ? Path.GetDirectoryName(presentation.FontFiles[0]) : null;

// Default to the system font directories (same pattern as the API's
// Demo:FontDirectory) so theme fonts missing from the deck can resolve against
// installed fonts instead of Typst's embedded serif fallback. TypstBridge keeps
// system fonts disabled natively, so without this the fallback chain emitted by
// the converter (… "Arial", "Helvetica", …) finds nothing.
var requestedFontPath = fontPath ?? DefaultSystemFontPath();
var compilerFontPath = requestedFontPath != null && embeddedFontPath != null
    ? string.Join(Path.PathSeparator, embeddedFontPath, requestedFontPath)
    : requestedFontPath ?? embeddedFontPath;

var options = new CompileOptions
{
    Format = format,
    WorkingDirectory = presentation.TempDirectory,
    FontDirectory = compilerFontPath
};

var result = compiler.Compile(typstSource, options);

if (!result.Success)
{
    Console.WriteLine($"Error: {result.ErrorMessage}");
    return 1;
}

if (format == OutputFormat.Png)
{
    Directory.CreateDirectory(outputPath);

    var width = Math.Max(2, result.Pages.Length.ToString().Length);
    for (var i = 0; i < result.Pages.Length; i++)
    {
        var pagePath = Path.Combine(outputPath, $"slide-{(i + 1).ToString().PadLeft(width, '0')}.png");
        File.WriteAllBytes(pagePath, result.Pages[i]);
    }

    Console.WriteLine($"PNG slides saved to: {outputPath}");
}
else
{
    File.WriteAllBytes(outputPath, result.Pages[0]);
    Console.WriteLine($"PDF saved to: {outputPath}");
}

return 0;

static string? DefaultSystemFontPath()
{
    string[] candidates =
    [
        "/System/Library/Fonts",
        "/Library/Fonts",
        "/usr/share/fonts",
        "/usr/local/share/fonts",
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".fonts"),
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".local", "share", "fonts"),
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts"),
    ];

    var existing = candidates.Where(Directory.Exists).ToList();
    return existing.Count > 0 ? string.Join(Path.PathSeparator, existing) : null;
}
