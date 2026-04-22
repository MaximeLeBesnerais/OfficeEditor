using DocxEditor.Core.Builders;
using DocxEditor.Core.Instructions;
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
