using DocxExamples;
using PptxExamples;
using XlsxExamples;

Console.WriteLine("╔════════════════════════════════════════════════════════╗");
Console.WriteLine("║       OfficeEditor Examples - Complete Suite           ║");
Console.WriteLine("╚════════════════════════════════════════════════════════╝");
Console.WriteLine();

Console.WriteLine("This program demonstrates all OfficeEditor features:");
Console.WriteLine("  • DOCX - Word document creation and editing");
Console.WriteLine("  • PPTX - PowerPoint presentations with Typst export");
Console.WriteLine("  • XLSX - Excel workbooks with formulas");
Console.WriteLine();

// Create output directory
var outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
Directory.CreateDirectory(outputDir);

Console.WriteLine($"Output directory: {outputDir}");
Console.WriteLine();

// Run all examples
try
{
    Console.WriteLine("Press any key to start DOCX examples...");
    Console.ReadKey(true);
    DocxExamples.Program.Main(args);
    
    Console.WriteLine("\nPress any key to start PPTX examples...");
    Console.ReadKey(true);
    PptxExamples.Program.Main(args);
    
    Console.WriteLine("\nPress any key to start XLSX examples...");
    Console.ReadKey(true);
    XlsxExamples.Program.Main(args);
    
    Console.WriteLine("\n╔════════════════════════════════════════════════════════╗");
    Console.WriteLine("║              All Examples Completed!                   ║");
    Console.WriteLine("╚════════════════════════════════════════════════════════╝");
    Console.WriteLine();
    Console.WriteLine($"Check the output folder: {outputDir}");
    Console.WriteLine();
    Console.WriteLine("Generated files:");
    Console.WriteLine("  📄 DOCX - Word documents with various content");
    Console.WriteLine("  📊 PPTX - PowerPoint presentations");
    Console.WriteLine("  📈 XLSX - Excel workbooks with data and formulas");
    Console.WriteLine("  📑 PDF  - Exported from PPTX via Typst");
    Console.WriteLine("  🖼️  PNG  - Slide thumbnails from PPTX");
    Console.WriteLine("  📝 TYP  - Typst source code from PPTX");
}
catch (Exception ex)
{
    Console.WriteLine($"\n❌ Error: {ex.Message}");
    Console.WriteLine(ex.StackTrace);
}

Console.WriteLine("\nPress any key to exit...");
Console.ReadKey(true);
