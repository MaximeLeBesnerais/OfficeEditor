using System.Text.Json;
using DocxEditor.Core.Builders;
using DocxEditor.Core.Generation;
using DocxEditor.Core.Generation.Assets;
using DocxEditor.Core.Generation.Design;
using DocxEditor.Core.Generation.Model;
using DocxEditor.Core.Generation.Schema;
using PptxEditor.Core.Builders;
using PptxEditor.Core.Generation.Archetypes;
using PptxEditor.Core.Generation.Components;
using PptxEditor.Core.Generation.Emit.Ooxml;
using PptxEditor.Core.Generation.Layout;
using PptxEditor.Core.Generation.Schema;
using XlsxEditor.Core.Builders;
using XlsxEditor.Core.Instructions;
using XlsxEditor.Core.Rendering;
using OfficeEditor.Core.Models;
using OfficeEditor.Core.Services;
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
            AnsiConsole.MarkupLine("Usage: officeeditor generate <input.json> [--output <output.pptx|output.docx|output.xlsx|output.pdf|output.png>] [--theme <name>]");
            return false;
        }

        var inputPath = args[1];
        string? outputPath = null;
        string? theme = null;

        for (var i = 2; i < args.Length; i++)
        {
            var arg = args[i];
            if (arg.Equals("--output", StringComparison.OrdinalIgnoreCase))
            {
                if (i + 1 >= args.Length)
                {
                    AnsiConsole.MarkupLine("[red]--output requires a file path ending in .pptx, .docx, or .xlsx.[/]");
                    return false;
                }
                outputPath = args[++i];
            }
            else if (arg.Equals("--theme", StringComparison.OrdinalIgnoreCase))
            {
                if (i + 1 >= args.Length)
                {
                    AnsiConsole.MarkupLine("[red]--theme requires a theme name.[/]");
                    return false;
                }
                theme = args[++i];
            }
            else
            {
                AnsiConsole.MarkupLine($"[red]Unknown option: {Markup.Escape(arg)}[/]");
                return false;
            }
        }

        var resolvedOutputPath = outputPath ?? Path.ChangeExtension(inputPath, ".pptx");
        var outputExtension = Path.GetExtension(resolvedOutputPath);
        if (!outputExtension.Equals(".pptx", StringComparison.OrdinalIgnoreCase) &&
            !outputExtension.Equals(".docx", StringComparison.OrdinalIgnoreCase) &&
            !outputExtension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) &&
            !outputExtension.Equals(".pdf", StringComparison.OrdinalIgnoreCase) &&
            !outputExtension.Equals(".png", StringComparison.OrdinalIgnoreCase))
        {
            AnsiConsole.MarkupLine("[red]Unsupported output format. Use a file path ending in .pptx, .docx, .xlsx, .pdf, or .png.[/]");
            return false;
        }

        if (theme is not null && !outputExtension.Equals(".docx", StringComparison.OrdinalIgnoreCase))
        {
            AnsiConsole.MarkupLine("[red]--theme is only supported for DOCX output (a .docx output path).[/]");
            return false;
        }

        if (theme is not null && DesignThemeCatalog.TryGet(theme) is null)
        {
            AnsiConsole.MarkupLine($"[red]Unknown theme '{Markup.Escape(theme)}'. Known themes: {string.Join(", ", DesignThemeCatalog.Themes.Keys)}.[/]");
            return false;
        }

        if (!File.Exists(inputPath))
        {
            AnsiConsole.MarkupLine($"[red]File not found: {Markup.Escape(inputPath)}[/]");
            return false;
        }

        var json = File.ReadAllText(inputPath);

        if (outputExtension.Equals(".docx", StringComparison.OrdinalIgnoreCase))
            return GenerateDocx(inputPath, resolvedOutputPath, json, theme);

        if (outputExtension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            return GenerateXlsx(resolvedOutputPath, json);

        if (outputExtension.Equals(".pdf", StringComparison.OrdinalIgnoreCase) ||
            outputExtension.Equals(".png", StringComparison.OrdinalIgnoreCase))
            return GenerateXlsxRender(resolvedOutputPath, outputExtension, json);

        var generated = false;

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
                File.WriteAllBytes(resolvedOutputPath, emitResult.Bytes);

                AnsiConsole.MarkupLine($"[green]{layout.Slides.Count} slides[/]  " +
                    $"[green]{emitResult.Bytes.Length} bytes[/]  " +
                    (layout.Warnings.Count + emitResult.Warnings.Count == 0
                        ? "[green]0 warnings[/]"
                        : $"[yellow]{layout.Warnings.Count + emitResult.Warnings.Count} warnings[/]"));
                generated = true;
            });

        if (!generated)
            return false;

        AnsiConsole.MarkupLine($"[green]Saved: {Markup.Escape(resolvedOutputPath)}[/]");
        return true;
    }

    static bool GenerateDocx(string inputPath, string outputPath, string json, string? theme)
    {
        var validation = new DocxGenerationDocumentParser().Validate(json);
        if (!validation.IsValid)
        {
            AnsiConsole.MarkupLine("[red]Validation errors:[/]");
            foreach (var error in validation.Errors)
                AnsiConsole.MarkupLine($"  [red]{Markup.Escape(error.ToString())}[/]");
            return false;
        }

        var document = validation.Document!;
        if (theme is not null)
        {
            var design = document.Design ?? new DesignTokens();
            document = document with { Design = design with { Theme = theme } };
        }

        var inputDirectory = Path.GetDirectoryName(Path.GetFullPath(inputPath))!;
        var sourceOptions = new ImageSourceOptions { AllowedRoot = inputDirectory };
        string? templatePath = null;
        if (document.TemplatePath is { } configuredTemplate)
            templatePath = ResolveTemplatePath(configuredTemplate, sourceOptions, inputDirectory);

        DocxEditor.Core.Generation.Contracts.DocxGenerationResult? result = null;
        AnsiConsole.Status()
            .Start("Generating DOCX...", _ =>
            {
                result = new DocxGenerator().Generate(
                    document,
                    outputPath,
                    new DocxGeneratorOptions
                    {
                        TemplatePath = templatePath,
                        ImageSourceOptions = sourceOptions
                    });
            });

        var warnings = validation.Warnings.Concat(result!.Warnings).ToList();
        foreach (var warning in warnings)
            AnsiConsole.MarkupLine($"[yellow]Warning: {Markup.Escape(warning.ToString())}[/]");

        AnsiConsole.MarkupLine($"[green]{result.Document.Sections.Count} sections[/]  " +
            (warnings.Count == 0 ? "[green]0 warnings[/]" : $"[yellow]{warnings.Count} warnings[/]"));
        AnsiConsole.MarkupLine($"[green]Saved: {Markup.Escape(outputPath)}[/]");
        return true;
    }

    static bool GenerateXlsx(string outputPath, string json)
    {
        XlsxGenerateResult? result = null;
        AnsiConsole.Status()
            .Start("Generating XLSX...", _ =>
            {
                result = XlsxGenerator.GenerateToFile(json, outputPath);
            });

        foreach (var diagnostic in result!.Validation.Errors)
        {
            var location = string.IsNullOrEmpty(diagnostic.Path) ? string.Empty : $" at {diagnostic.Path}";
            AnsiConsole.MarkupLine($"[red]{Markup.Escape(diagnostic.Message)}{Markup.Escape(location)}[/]");
        }
        foreach (var diagnostic in result.Validation.Warnings)
        {
            var location = string.IsNullOrEmpty(diagnostic.Path) ? string.Empty : $" at {diagnostic.Path}";
            AnsiConsole.MarkupLine($"[yellow]Warning: {Markup.Escape(diagnostic.Message)}{Markup.Escape(location)}[/]");
        }

        if (!result.IsValid || result.Bytes is null)
        {
            AnsiConsole.MarkupLine("[red]XLSX generation failed; no output was written.[/]");
            return false;
        }

        AnsiConsole.MarkupLine($"[green]{result.Bytes.Length} bytes[/]  " +
            (result.Validation.Warnings.Any()
                ? $"[yellow]{result.Validation.Warnings.Count()} warnings[/]"
                : "[green]0 warnings[/]"));
        AnsiConsole.MarkupLine($"[green]Saved: {Markup.Escape(outputPath)}[/]");
        return true;
    }

    static bool GenerateXlsxRender(string outputPath, string outputExtension, string json)
    {
        var format = outputExtension.Equals(".png", StringComparison.OrdinalIgnoreCase)
            ? OutputFormat.Png
            : OutputFormat.Pdf;

        XlsxEditor.Core.Rendering.XlsxRenderResult? result = null;
        AnsiConsole.Status()
            .Start("Rendering XLSX via Typst...", _ =>
            {
                result = XlsxEditor.Core.Rendering.XlsxRenderer.RenderJson(json, new CompileOptions { Format = format });
            });

        if (result is null || !result.Success)
        {
            AnsiConsole.MarkupLine($"[red]Rendering failed: {Markup.Escape(result?.Compile.ErrorMessage ?? "unknown error")}[/]");
            return false;
        }

        var directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(directory))
            Directory.CreateDirectory(directory);

        if (format == OutputFormat.Pdf)
        {
            File.WriteAllBytes(outputPath, result.Compile.Pages[0]);
            AnsiConsole.MarkupLine($"[green]PDF saved: {Markup.Escape(outputPath)}[/] ({result.Compile.Pages.Length} page(s))");
        }
        else
        {
            Directory.CreateDirectory(outputPath);
            for (int i = 0; i < result.Compile.Pages.Length; i++)
            {
                string pagePath = Path.Combine(outputPath, $"page-{i + 1:D3}.png");
                File.WriteAllBytes(pagePath, result.Compile.Pages[i]);
            }

            AnsiConsole.MarkupLine($"[green]{result.Compile.Pages.Length} PNG page(s) saved to: {Markup.Escape(outputPath)}[/]");
        }

        return true;
    }

    static string ResolveTemplatePath(
        string configuredPath,
        ImageSourceOptions sourceOptions,
        string inputDirectory)
    {
        try
        {
            var resolution = DocxEditor.Core.Generation.Assets.ImageSourceResolver.Resolve(configuredPath, sourceOptions);
            if (resolution.Kind != ImageSourceKind.LocalFile || resolution.FilePath is null)
                throw new ArgumentException("Template must be a local .docx file path.");
            return resolution.FilePath;
        }
        catch (ImageSourceException ex)
        {
            throw new ArgumentException(
                $"Invalid template path '{configuredPath}'. Templates must be local files inside the input JSON directory '{inputDirectory}'. {ex.Message}",
                ex);
        }
    }

    static void ShowHelp()
    {
        AnsiConsole.WriteLine("OfficeEditor CLI - Unified Office Document Editor");
        AnsiConsole.WriteLine();
        AnsiConsole.WriteLine("Usage:");
        AnsiConsole.WriteLine("  officeeditor create <output.file> [--type docx|pptx|xlsx] [--text \"content\"] [--title \"title\"] [--sheet \"name\"]");
        AnsiConsole.WriteLine("  officeeditor generate <input.json> [--output <output.pptx|output.docx|output.xlsx>] [--theme <name>]");
        AnsiConsole.WriteLine("  officeeditor detect <template.file>");
        AnsiConsole.WriteLine("  officeeditor merge <template.file> <data.json> <output.file>");
        AnsiConsole.WriteLine();
        AnsiConsole.WriteLine("generate options:");
        AnsiConsole.WriteLine("  --output <path>  Output file. Extension selects the format (.pptx|.docx|.xlsx);");
        AnsiConsole.WriteLine("                   defaults to <input.json> with a .pptx extension.");
        AnsiConsole.WriteLine("  --theme <name>   DOCX only: named design theme to apply (editorial|corporate).");
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
