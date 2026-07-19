using System.Text.Json;
using DocxEditor.Core.Builders;
using PptxEditor.Core.Builders;
using PptxEditor.Core.Generation.Archetypes;
using PptxEditor.Core.Generation.Components;
using PptxEditor.Core.Generation.Emit.Ooxml;
using PptxEditor.Core.Generation.Layout;
using PptxEditor.Core.Generation.Schema;
using XlsxEditor.Core.Builders;
using OfficeEditor.Core.Models;
using Spectre.Console;

namespace OfficeEditor.Cli;

class Program
{
    static int Main(string[] args)
    {
        if (args.Length == 0)
        {
            ShowHelp();
            return 1;
        }

        var command = args[0].ToLowerInvariant();

        try
        {
            switch (command)
            {
                case "create":
                    return HandleCreate(args) ? 0 : 1;
                case "edit":
                    return HandleEdit(args) ? 0 : 1;
                case "detect":
                    return HandleDetect(args) ? 0 : 1;
                case "merge":
                    return HandleMerge(args) ? 0 : 1;
                case "generate":
                    return HandleGenerate(args) ? 0 : 1;
                case "help":
                case "--help":
                case "-h":
                    ShowHelp();
                    return 0;
                default:
                    AnsiConsole.MarkupLine("[red]Unknown command.[/]");
                    ShowHelp();
                    return 1;
            }
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error: {Markup.Escape(ex.Message)}[/]");
            return 1;
        }
    }

    static bool HandleGenerate(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Input JSON file path is required.[/]");
            AnsiConsole.MarkupLine("Usage: officeeditor generate <input.json> [--output <output.pptx>]");
            return false;
        }

        var inputPath = args[1];
        var outputPath = GetArgumentValue(args, "--output") ?? Path.ChangeExtension(inputPath, ".pptx");

        if (!File.Exists(inputPath))
        {
            AnsiConsole.MarkupLine($"[red]File not found: {Markup.Escape(inputPath)}[/]");
            return false;
        }

        var json = File.ReadAllText(inputPath);

        AnsiConsole.Status()
            .Start("Generating...", ctx =>
            {
                var result = new GenerationDocumentParser().Validate(json);
                if (!result.IsValid)
                {
                    AnsiConsole.MarkupLine("[red]Validation errors:[/]");
                    foreach (var e in result.Errors)
                        AnsiConsole.MarkupLine($"  [red]{Markup.Escape(e.ToString())}[/]");
                    return;
                }

                var doc = result.Document!;
                var archetyped = ArchetypeExpander.Expand(doc);
                var componentized = ComponentExpander.Expand(archetyped);
                var layout = new LayoutResolver().Resolve(componentized);
                var emitResult = new OoxmlEmitter().Emit(layout);
                File.WriteAllBytes(outputPath, emitResult.Bytes);

                AnsiConsole.MarkupLine($"[green]{layout.Slides.Count} slides[/]  " +
                    $"[green]{emitResult.Bytes.Length} bytes[/]  " +
                    (layout.Warnings.Count + emitResult.Warnings.Count == 0
                        ? "[green]0 warnings[/]"
                        : $"[yellow]{layout.Warnings.Count + emitResult.Warnings.Count} warnings[/]"));
            });

        AnsiConsole.MarkupLine($"[green]Saved: {Markup.Escape(outputPath)}[/]");
        return true;
    }

    static void ShowHelp()
    {
        AnsiConsole.WriteLine("OfficeEditor CLI - Unified Office Document Editor");
        AnsiConsole.WriteLine();
        AnsiConsole.WriteLine("Usage:");
        AnsiConsole.WriteLine("  officeeditor create <output.file> [--type docx|pptx|xlsx] [--text \"content\"] [--title \"title\"] [--sheet \"name\"]");
        AnsiConsole.WriteLine("  officeeditor generate <input.json> [--output <output.pptx>]");
        AnsiConsole.WriteLine("  officeeditor detect <template.file>");
        AnsiConsole.WriteLine("  officeeditor merge <template.file> <data.json> <output.file>");
        AnsiConsole.WriteLine();
        AnsiConsole.WriteLine("Supported formats: .docx (Word), .pptx (PowerPoint), .xlsx (Excel)");
        AnsiConsole.WriteLine("Format is auto-detected from file extension.");
    }

    static bool HandleCreate(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Output file path is required.[/]");
            return false;
        }

        var outputPath = args[1];
        var format = DetectFormat(outputPath);
        
        if (format == DocumentFormat.Unknown)
        {
            AnsiConsole.MarkupLine("[red]Unsupported file format. Use .docx, .pptx, or .xlsx[/]");
            return false;
        }

        AnsiConsole.Status()
            .Start($"Creating {format} document...", ctx =>
            {
                switch (format)
                {
                    case DocumentFormat.Docx:
                        CreateDocx(outputPath, args);
                        break;
                    case DocumentFormat.Pptx:
                        CreatePptx(outputPath, args);
                        break;
                    case DocumentFormat.Xlsx:
                        CreateXlsx(outputPath, args);
                        break;
                }
            });

        AnsiConsole.MarkupLine($"[green]Document created: {Markup.Escape(outputPath)}[/]");
        return true;
    }

    static void CreateDocx(string path, string[] args)
    {
        var text = GetArgumentValue(args, "--text") ?? "Hello World";
        var style = GetArgumentValue(args, "--style");

        using var builder = DocumentBuilder.Create(path);
        builder.AddParagraph(text, style);
        builder.Save();
    }

    static void CreatePptx(string path, string[] args)
    {
        var title = GetArgumentValue(args, "--title") ?? "Presentation";

        using var builder = PresentationBuilder.Create(path);
        builder.AddSlide();
        builder.CurrentSlide.AddTitle(title);
        builder.Save();
    }

    static void CreateXlsx(string path, string[] args)
    {
        var sheetName = GetArgumentValue(args, "--sheet") ?? "Sheet1";

        using var builder = WorkbookBuilder.Create(path);
        builder.AddWorksheet(sheetName);
        builder.Save();
    }

    static bool HandleEdit(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Input file and instructions file are required.[/]");
            return false;
        }

        var inputPath = args[1];
        var instructionsPath = GetArgumentValue(args, "--instructions");
        var format = DetectFormat(inputPath);

        if (string.IsNullOrEmpty(instructionsPath))
        {
            AnsiConsole.MarkupLine("[red]--instructions parameter is required.[/]");
            return false;
        }

        AnsiConsole.Status()
            .Start("Editing document...", ctx =>
            {
                // For V1, edit is format-specific
                // Full implementation would parse and execute instructions
                AnsiConsole.MarkupLine($"[yellow]Edit not yet implemented for {format}[/]");
            });
        return false;
    }

    static bool HandleDetect(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Document path is required.[/]");
            return false;
        }

        var documentPath = args[1];
        var format = DetectFormat(documentPath);

        if (format == DocumentFormat.Unknown)
        {
            AnsiConsole.MarkupLine("[red]Unsupported file format.[/]");
            return false;
        }

        AnsiConsole.Status()
            .Start("Scanning for variables...", ctx =>
            {
                List<VariableInfo> variables = format switch
                {
                    DocumentFormat.Docx => DetectDocxVariables(documentPath),
                    DocumentFormat.Pptx => DetectPptxVariables(documentPath),
                    DocumentFormat.Xlsx => DetectXlsxVariables(documentPath),
                    _ => new List<VariableInfo>()
                };
                
                if (variables.Count == 0)
                {
                    AnsiConsole.MarkupLine("[yellow]No variables found in document.[/]");
                    return;
                }

                AnsiConsole.MarkupLine($"[green]Found {variables.Count} variables:[/]");
                
                var table = new Table();
                table.AddColumn("Variable");
                table.AddColumn("Default Value");
                table.AddColumn("Location");

                foreach (var variable in variables)
                {
                    table.AddRow(
                        variable.Name,
                        variable.DefaultValue ?? "(none)",
                        variable.Location
                    );
                }

                AnsiConsole.Write(table);
            });
        return true;
    }

    static List<VariableInfo> DetectDocxVariables(string path)
    {
        using var builder = DocumentBuilder.Open(path);
        return builder.DetectVariables();
    }

    static List<VariableInfo> DetectPptxVariables(string path)
    {
        using var builder = PresentationBuilder.Open(path);
        return builder.DetectVariables();
    }

    static List<VariableInfo> DetectXlsxVariables(string path)
    {
        using var builder = WorkbookBuilder.Open(path);
        return builder.DetectVariables();
    }

    static bool HandleMerge(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Template file, data file, and output file are required.[/]");
            return false;
        }

        var templatePath = args[1];
        var dataPath = args[2];
        var outputPath = args[3];
        var format = DetectFormat(templatePath);
        if (format == DocumentFormat.Unknown)
        {
            AnsiConsole.MarkupLine("[red]Unsupported file format.[/]");
            return false;
        }

        var json = File.ReadAllText(dataPath);
        var data = JsonSerializer.Deserialize<Dictionary<string, string>>(json);

        if (data == null)
        {
            AnsiConsole.MarkupLine("[red]Invalid data file.[/]");
            return false;
        }

        AnsiConsole.Status()
            .Start("Merging variables...", ctx =>
            {
                File.Copy(templatePath, outputPath, true);
                
                switch (format)
                {
                    case DocumentFormat.Docx:
                        using (var builder = DocumentBuilder.Open(outputPath))
                        {
                            builder.MergeVariables(data);
                            builder.Save();
                        }
                        break;
                    case DocumentFormat.Pptx:
                        using (var builder = PresentationBuilder.Open(outputPath))
                        {
                            builder.MergeVariables(data);
                            builder.Save();
                        }
                        break;
                    case DocumentFormat.Xlsx:
                        using (var builder = WorkbookBuilder.Open(outputPath))
                        {
                            builder.MergeVariables(data);
                            builder.Save();
                        }
                        break;
                }

                AnsiConsole.MarkupLine($"[green]Document merged: {Markup.Escape(outputPath)}[/]");
            });
        return true;
    }

    static DocumentFormat DetectFormat(string path)
    {
        var extension = Path.GetExtension(path).ToLowerInvariant();
        return extension switch
        {
            ".docx" => DocumentFormat.Docx,
            ".pptx" => DocumentFormat.Pptx,
            ".xlsx" => DocumentFormat.Xlsx,
            _ => DocumentFormat.Unknown
        };
    }

    static string? GetArgumentValue(string[] args, string key)
    {
        for (int i = 0; i < args.Length - 1; i++)
        {
            if (args[i].Equals(key, StringComparison.OrdinalIgnoreCase))
            {
                return args[i + 1];
            }
        }
        return null;
    }
}

enum DocumentFormat
{
    Unknown,
    Docx,
    Pptx,
    Xlsx
}
