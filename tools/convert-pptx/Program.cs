using DocumentFormat.OpenXml.Packaging;
using PptxEditor.Core.Converters;
using OfficeEditor.Core.Services;

if (args.Length < 2)
{
    Console.WriteLine("Usage: convert-pptx <input.pptx> <output.pdf>");
    return 1;
}

var inputPath = args[0];
var outputPath = args[1];

if (!File.Exists(inputPath))
{
    Console.WriteLine($"Error: File not found: {inputPath}");
    return 1;
}

Console.WriteLine($"Converting {inputPath}...");

using var doc = PresentationDocument.Open(inputPath, false);
var converter = new PptxToTypstConverter(doc);
var presentation = converter.Convert();

// Generate Typst source
var typstSource = converter.GenerateTypstSource(presentation);

// Save Typst source for debugging
var typstPath = Path.ChangeExtension(outputPath, ".typ");
File.WriteAllText(typstPath, typstSource);
Console.WriteLine($"Typst source saved to: {typstPath}");

// Compile with Typst
Console.WriteLine("Compiling with Typst...");
using var compiler = new TypstCompilerService();

var options = new CompileOptions
{
    Format = OutputFormat.Pdf,
    WorkingDirectory = presentation.TempDirectory,
    FontDirectory = presentation.FontFiles.Count > 0 ? Path.GetDirectoryName(presentation.FontFiles[0]) : null
};

var result = compiler.Compile(typstSource, options);

if (!result.Success)
{
    Console.WriteLine($"Error: {result.ErrorMessage}");
    return 1;
}

// Save PDF
File.WriteAllBytes(outputPath, result.Pages[0]);
Console.WriteLine($"PDF saved to: {outputPath}");

return 0;
