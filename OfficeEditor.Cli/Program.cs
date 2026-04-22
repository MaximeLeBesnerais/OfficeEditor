using System.Text.Json;
using DocxEditor.Core.Builders;
using PptxEditor.Core.Builders;
using XlsxEditor.Core.Builders;
using OfficeEditor.Core.Models;
using Spectre.Console;

namespace OfficeEditor.Cli;

class Program
{
    static void Main(string[] args)
    {
        if (args.Length == 0)
        {
            ShowHelp();
            return;
        }

        var command = args[0].ToLowerInvariant();

        try
        {
            switch (command)
            {
                case "create":
                    HandleCreate(args);
                    break;
                case "edit":
                    HandleEdit(args);
                    break;
                case "detect":
                    HandleDetect(args);
                    break;
                case "merge":
                    HandleMerge(args);
                    break;
                case "help":
                case "--help":
                case "-h":
                    ShowHelp();
                    break;
                default:
                    AnsiConsole.MarkupLine("[red]Unknown command.[/]");
                    ShowHelp();
                    break;
            }
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error: {ex.Message}[/]");
            Environment.Exit(1);
        }
    }

    static void ShowHelp()
    {
        AnsiConsole.WriteLine("OfficeEditor CLI - Unified Office Document Editor");
        AnsiConsole.WriteLine();
        AnsiConsole.WriteLine("Usage:");
        AnsiConsole.WriteLine("  officeeditor create <output.file> [--type docx|pptx|xlsx] [--text \"content\"] [--title \"title\"] [--sheet \"name\"]");
        AnsiConsole.WriteLine("  officeeditor edit <input.file> --instructions <file.json|file.yaml>");
        AnsiConsole.WriteLine("  officeeditor detect <template.file>");
        AnsiConsole.WriteLine("  officeeditor merge <template.file> <data.json> <output.file>");
        AnsiConsole.WriteLine();
        AnsiConsole.WriteLine("Supported formats: .docx (Word), .pptx (PowerPoint), .xlsx (Excel)");
        AnsiConsole.WriteLine("Format is auto-detected from file extension.");
    }

    static void HandleCreate(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Output file path is required.[/]");
            return;
        }

        var outputPath = args[1];
        var format = DetectFormat(outputPath);
        
        if (format == DocumentFormat.Unknown)
        {
            AnsiConsole.MarkupLine("[red]Unsupported file format. Use .docx, .pptx, or .xlsx[/]");
            return;
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

        AnsiConsole.MarkupLine($"[green]Document created: {outputPath}[/]");
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

    static void HandleEdit(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Input file and instructions file are required.[/]");
            return;
        }

        var inputPath = args[1];
        var instructionsPath = GetArgumentValue(args, "--instructions");
        var format = DetectFormat(inputPath);

        if (string.IsNullOrEmpty(instructionsPath))
        {
            AnsiConsole.MarkupLine("[red]--instructions parameter is required.[/]");
            return;
        }

        AnsiConsole.Status()
            .Start("Editing document...", ctx =>
            {
                // For V1, edit is format-specific
                // Full implementation would parse and execute instructions
                AnsiConsole.MarkupLine($"[yellow]Edit not yet implemented for {format}[/]");
            });
    }

    static void HandleDetect(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Document path is required.[/]");
            return;
        }

        var documentPath = args[1];
        var format = DetectFormat(documentPath);

        if (format == DocumentFormat.Unknown)
        {
            AnsiConsole.MarkupLine("[red]Unsupported file format.[/]");
            return;
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

    static void HandleMerge(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Template file, data file, and output file are required.[/]");
            return;
        }

        var templatePath = args[1];
        var dataPath = args[2];
        var outputPath = args[3];
        var format = DetectFormat(templatePath);

        var json = File.ReadAllText(dataPath);
        var data = JsonSerializer.Deserialize<Dictionary<string, string>>(json);

        if (data == null)
        {
            AnsiConsole.MarkupLine("[red]Invalid data file.[/]");
            return;
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

                AnsiConsole.MarkupLine($"[green]Document merged: {outputPath}[/]");
            });
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
