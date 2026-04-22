using System.Text.Json;
using DocxEditor.Core.Builders;
using DocxEditor.Core.Instructions;
using DocxEditor.Core.Models;
using DocxEditor.Core.Serialization;
using Spectre.Console;

namespace DocxEditor.Cli;

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
                case "template":
                    HandleTemplate(args);
                    break;
                case "validate":
                    HandleValidate(args);
                    break;
                case "detect":
                    HandleDetect(args);
                    break;
                case "merge":
                    HandleMerge(args);
                    break;
                case "markdown":
                    HandleMarkdown(args);
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
        AnsiConsole.WriteLine("DocxEditor CLI");
        AnsiConsole.WriteLine();
        AnsiConsole.WriteLine("Usage:");
        AnsiConsole.WriteLine("  docxeditor create <output.docx> [--text \"content\"] [--style \"styleId\"]");
        AnsiConsole.WriteLine("  docxeditor edit <input.docx> --instructions <file.json|file.yaml>");
        AnsiConsole.WriteLine("  docxeditor template <template.docx> <output.docx> --instructions <file.json|file.yaml>");
        AnsiConsole.WriteLine("  docxeditor validate <instructions.json|instructions.yaml>");
        AnsiConsole.WriteLine("  docxeditor detect <template.docx>");
        AnsiConsole.WriteLine("  docxeditor merge <template.docx> <data.json> <output.docx>");
        AnsiConsole.WriteLine("  docxeditor markdown <input.md> <output.docx> [--style-map styles.json]");
    }

    static void HandleCreate(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Output file path is required.[/]");
            return;
        }

        var outputPath = args[1];
        var text = GetArgumentValue(args, "--text") ?? "Hello World";
        var style = GetArgumentValue(args, "--style");

        AnsiConsole.Status()
            .Start("Creating document...", ctx =>
            {
                using var builder = DocumentBuilder.Create(outputPath);
                builder.AddParagraph(text, style);
                builder.Save();
            });

        AnsiConsole.MarkupLine($"[green]Document created: {outputPath}[/]");
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

        if (string.IsNullOrEmpty(instructionsPath))
        {
            AnsiConsole.MarkupLine("[red]--instructions parameter is required.[/]");
            return;
        }

        var instructions = LoadInstructions(instructionsPath);
        
        AnsiConsole.Status()
            .Start("Editing document...", ctx =>
            {
                using var builder = DocumentBuilder.Open(inputPath);
                var engine = new InstructionEngine();
                engine.Execute(builder, instructions);
                builder.Save();
            });

        AnsiConsole.MarkupLine($"[green]Document edited: {inputPath}[/]");
    }

    static void HandleTemplate(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Template file, output file, and instructions file are required.[/]");
            return;
        }

        var templatePath = args[1];
        var outputPath = args[2];
        var instructionsPath = GetArgumentValue(args, "--instructions");

        if (string.IsNullOrEmpty(instructionsPath))
        {
            AnsiConsole.MarkupLine("[red]--instructions parameter is required.[/]");
            return;
        }

        var instructions = LoadInstructions(instructionsPath);
        
        AnsiConsole.Status()
            .Start("Processing template...", ctx =>
            {
                File.Copy(templatePath, outputPath, true);
                using var builder = DocumentBuilder.Open(outputPath);
                var engine = new InstructionEngine();
                engine.Execute(builder, instructions);
                builder.Save();
            });

        AnsiConsole.MarkupLine($"[green]Document created from template: {outputPath}[/]");
    }

    static void HandleValidate(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Instructions file path is required.[/]");
            return;
        }

        var instructionsPath = args[1];

        try
        {
            var instructions = LoadInstructions(instructionsPath);
            AnsiConsole.MarkupLine($"[green]Valid instructions file. Found {instructions.Operations.Count} operations.[/]");
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Invalid instructions file: {ex.Message}[/]");
        }
    }

    static void HandleDetect(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Document path is required.[/]");
            return;
        }

        var documentPath = args[1];

        AnsiConsole.Status()
            .Start("Scanning for variables...", ctx =>
            {
                using var builder = DocumentBuilder.Open(documentPath);
                var variables = builder.DetectVariables();
                
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

    static void HandleMerge(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Template file, data file, and output pattern are required.[/]");
            return;
        }

        var templatePath = args[1];
        var dataPath = args[2];
        var outputPattern = args[3];

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
                var outputPath = outputPattern;
                foreach (var kvp in data)
                {
                    outputPath = outputPath.Replace($"{{{kvp.Key}}}", kvp.Value);
                }

                File.Copy(templatePath, outputPath, true);
                
                using var builder = DocumentBuilder.Open(outputPath);
                builder.MergeVariables(data);
                builder.Save();

                AnsiConsole.MarkupLine($"[green]Document merged: {outputPath}[/]");
            });
    }

    static void HandleMarkdown(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Input markdown file and output docx file are required.[/]");
            return;
        }

        var inputPath = args[1];
        var outputPath = args[2];
        var styleMapPath = GetArgumentValue(args, "--style-map");

        var markdown = File.ReadAllText(inputPath);
        StyleMapping? styleMap = null;

        if (!string.IsNullOrEmpty(styleMapPath))
        {
            var styleJson = File.ReadAllText(styleMapPath);
            styleMap = JsonSerializer.Deserialize<StyleMapping>(styleJson);
        }

        AnsiConsole.Status()
            .Start("Converting markdown...", ctx =>
            {
                using var builder = DocumentBuilder.Create(outputPath);
                builder.AddMarkdown(markdown, styleMap);
                builder.Save();
            });

        AnsiConsole.MarkupLine($"[green]Document created from markdown: {outputPath}[/]");
    }

    static DocxEditor.Core.Models.DocumentInstructions LoadInstructions(string path)
    {
        var content = File.ReadAllText(path);
        
        if (path.EndsWith(".yaml", StringComparison.OrdinalIgnoreCase) || 
            path.EndsWith(".yml", StringComparison.OrdinalIgnoreCase))
        {
            var parser = new YamlInstructionParser();
            return parser.Parse(content);
        }
        else
        {
            var parser = new JsonInstructionParser();
            return parser.Parse(content);
        }
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
