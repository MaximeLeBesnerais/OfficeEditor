using System.Text.Json;
using DocxEditor.Core.Builders;
using DocxEditor.Core.Instructions;
using DocxEditor.Core.Markdown;
using DocxEditor.Core.Markdown.Model;
using DocxEditor.Core.Markdown.Rendering;
using DocxEditor.Core.Models;
using DocxEditor.Core.Serialization;
using DocumentFormat.OpenXml.Packaging;
using OfficeEditor.Core.Models;
using Spectre.Console;

namespace DocxEditor.Cli;

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
                    return HandleCreate(args);
                case "edit":
                    return HandleEdit(args);
                case "template":
                    return HandleTemplate(args);
                case "validate":
                    return HandleValidate(args);
                case "detect":
                    return HandleDetect(args);
                case "merge":
                    return HandleMerge(args);
                case "markdown":
                    return HandleMarkdown(args);
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
        AnsiConsole.WriteLine("  docxeditor markdown <input.md> <output.docx> [--template <template.docx>] [--style-map <styles.json>] [--strict]");
        AnsiConsole.WriteLine();
        AnsiConsole.WriteLine("markdown options:");
        AnsiConsole.WriteLine("  --template <path>   Base the document on an existing template; its styles are preserved");
        AnsiConsole.WriteLine("                      and custom style-map style names must resolve against it.");
        AnsiConsole.WriteLine("  --style-map <path>  JSON mapping of markdown element -> Word style name or StyleId.");
        AnsiConsole.WriteLine("  --strict            Treat unresolved constructs and style references as errors.");
    }

    static int HandleCreate(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Output file path is required.[/]");
            return 1;
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
        return 0;
    }

    static int HandleEdit(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Input file and instructions file are required.[/]");
            return 1;
        }

        var inputPath = args[1];
        var instructionsPath = GetArgumentValue(args, "--instructions");

        if (string.IsNullOrEmpty(instructionsPath))
        {
            AnsiConsole.MarkupLine("[red]--instructions parameter is required.[/]");
            return 1;
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
        return 0;
    }

    static int HandleTemplate(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Template file, output file, and instructions file are required.[/]");
            return 1;
        }

        var templatePath = args[1];
        var outputPath = args[2];
        var instructionsPath = GetArgumentValue(args, "--instructions");

        if (string.IsNullOrEmpty(instructionsPath))
        {
            AnsiConsole.MarkupLine("[red]--instructions parameter is required.[/]");
            return 1;
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
        return 0;
    }

    static int HandleValidate(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Instructions file path is required.[/]");
            return 1;
        }

        var instructionsPath = args[1];

        try
        {
            var instructions = LoadInstructions(instructionsPath);
            AnsiConsole.MarkupLine($"[green]Valid instructions file. Found {instructions.Operations.Count} operations.[/]");
            return 0;
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Invalid instructions file: {Markup.Escape(ex.Message)}[/]");
            return 1;
        }
    }

    static int HandleDetect(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Document path is required.[/]");
            return 1;
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
        return 0;
    }

    static int HandleMerge(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Template file, data file, and output pattern are required.[/]");
            return 1;
        }

        var templatePath = args[1];
        var dataPath = args[2];
        var outputPattern = args[3];

        var json = File.ReadAllText(dataPath);
        var data = JsonSerializer.Deserialize<Dictionary<string, string>>(json);

        if (data == null)
        {
            AnsiConsole.MarkupLine("[red]Invalid data file.[/]");
            return 1;
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
        return 0;
    }

    static int HandleMarkdown(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Input markdown file and output docx file are required.[/]");
            return 1;
        }

        var inputPath = args[1];
        var outputPath = args[2];
        string? templatePath = null;
        string? styleMapPath = null;
        var strict = false;

        for (var i = 3; i < args.Length; i++)
        {
            var arg = args[i];
            if (arg.Equals("--template", StringComparison.OrdinalIgnoreCase))
            {
                if (i + 1 >= args.Length)
                {
                    AnsiConsole.MarkupLine("[red]--template requires a .docx template path.[/]");
                    return 1;
                }
                templatePath = args[++i];
            }
            else if (arg.Equals("--style-map", StringComparison.OrdinalIgnoreCase))
            {
                if (i + 1 >= args.Length)
                {
                    AnsiConsole.MarkupLine("[red]--style-map requires a JSON file path.[/]");
                    return 1;
                }
                styleMapPath = args[++i];
            }
            else if (arg.Equals("--strict", StringComparison.OrdinalIgnoreCase))
            {
                strict = true;
            }
            else
            {
                AnsiConsole.MarkupLine($"[red]Unknown option: {Markup.Escape(arg)}[/]");
                return 1;
            }
        }

        if (!File.Exists(inputPath))
        {
            AnsiConsole.MarkupLine($"[red]Input markdown file not found: {Markup.Escape(inputPath)}[/]");
            return 1;
        }
        if (templatePath is not null && !File.Exists(templatePath))
        {
            AnsiConsole.MarkupLine($"[red]Template file not found: {Markup.Escape(templatePath)}[/]");
            return 1;
        }
        if (styleMapPath is not null && !File.Exists(styleMapPath))
        {
            AnsiConsole.MarkupLine($"[red]Style-map file not found: {Markup.Escape(styleMapPath)}[/]");
            return 1;
        }

        if (!Path.GetExtension(outputPath).Equals(".docx", StringComparison.OrdinalIgnoreCase))
        {
            AnsiConsole.MarkupLine("[red]Output file must end in .docx.[/]");
            return 1;
        }

        var markdown = File.ReadAllText(inputPath);
        var styleMap = LoadStyleMap(styleMapPath);

        // Local image paths are resolved relative to the input markdown directory and are
        // confined to it: sources escaping the directory (or absolute paths) are rejected.
        var inputDirectory = Path.GetDirectoryName(Path.GetFullPath(inputPath))!;
        var imageSourceOptions = new DocxEditor.Core.Generation.Assets.ImageSourceOptions { AllowedRoot = inputDirectory };

        // Custom style names must resolve against the template when one is supplied. The
        // resolver is read-only over the template: unresolved/ambiguous references surface
        // as path-qualified warnings (permissive) or errors (strict), never silent fallbacks.
        // Only the author's explicit style-map entries plus the styles the input actually
        // renders are validated; default-map entries for constructs the input never uses
        // (e.g. "TableGrid" in a document without tables) are never rejected.
        if (templatePath is not null)
        {
            var parsedForKeys = new RichMarkdownParser(
                strict ? MarkdownParseOptions.StrictMode : MarkdownParseOptions.Default)
                .Parse(markdown);
            var referencedStyleKeys = RichMarkdownRenderer.CollectReferencedStyleKeys(parsedForKeys.Document);
            var styleDiagnostics = ResolveTemplateStyles(
                templatePath, styleMap, strict,
                referencedStyleKeys, userProvided: styleMapPath is not null);
            foreach (var diagnostic in styleDiagnostics)
                PrintStyleDiagnostic(styleMapPath ?? templatePath, diagnostic);

            if (strict && styleDiagnostics.Any(d => d.Severity == MarkdownDiagnosticSeverity.Error))
            {
                AnsiConsole.MarkupLine("[red]Markdown conversion failed: template style resolution errors above.[/]");
                return 1;
            }
        }

        var renderOptions = new MarkdownRenderOptions
        {
            StyleMapping = styleMap,
            Strict = strict,
            ParseOptions = strict ? MarkdownParseOptions.StrictMode : MarkdownParseOptions.Default,
            ImageSourceOptions = imageSourceOptions
        };

        // Generate atomically: the document is fully written to a sibling temp file and only
        // then moved over the destination, so a failure never truncates an existing output.
        // Render diagnostics are evaluated against the temp file BEFORE it is published: on any
        // error only this CLI's own temp file is deleted and the destination is left untouched.
        var fullOutputPath = Path.GetFullPath(outputPath);
        var outputDirectory = Path.GetDirectoryName(fullOutputPath)
            ?? throw new ArgumentException($"Cannot resolve the directory of output path '{outputPath}'.");
        Directory.CreateDirectory(outputDirectory);
        var tempPath = Path.Combine(outputDirectory, $".{Path.GetFileName(fullOutputPath)}.{Guid.NewGuid():N}.tmp");

        MarkdownRenderResult? renderResult = null;
        try
        {
            AnsiConsole.Status()
                .Start("Converting markdown...", _ =>
                {
                    if (templatePath is not null)
                    {
                        File.Copy(templatePath, tempPath, true);
                        using var builder = DocumentBuilder.Open(tempPath);
                        builder.AddRichMarkdown(markdown, renderOptions);
                        renderResult = builder.LastRichMarkdownResult;
                        builder.Save();
                    }
                    else
                    {
                        using var builder = DocumentBuilder.Create(tempPath);
                        builder.AddRichMarkdown(markdown, renderOptions);
                        renderResult = builder.LastRichMarkdownResult;
                        builder.Save();
                    }
                });

            if (renderResult is not null)
            {
                foreach (var diagnostic in renderResult.Diagnostics)
                    PrintMarkdownDiagnostic(inputPath, diagnostic);

                if (renderResult.HasErrors)
                {
                    AnsiConsole.MarkupLine("[red]Markdown conversion failed: errors above.[/]");
                    return 1;
                }
            }

            // Publish only after the diagnostics pass; the move is the single atomic replace.
            File.Move(tempPath, fullOutputPath, overwrite: true);
        }
        finally
        {
            TryDeleteTempFile(tempPath);
        }

        AnsiConsole.MarkupLine($"[green]Document created from markdown: {outputPath}[/]");
        return 0;
    }

    /// <summary>
    /// Best-effort cleanup of the CLI's own sibling temp file. Only <paramref name="tempPath"/>
    /// is touched — the destination and any unrelated files are never deleted. Cleanup failures
    /// are swallowed so a successful conversion is never reported as failed.
    /// </summary>
    private static void TryDeleteTempFile(string tempPath)
    {
        try
        {
            if (File.Exists(tempPath))
            {
                File.Delete(tempPath);
            }
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }

    static StyleMapping LoadStyleMap(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return StyleMapping.Default;
        }

        var json = File.ReadAllText(path);
        var map = new Dictionary<string, string>(StringComparer.Ordinal);

        using (var document = JsonDocument.Parse(json))
        {
            var root = document.RootElement;
            if (root.ValueKind != JsonValueKind.Object)
            {
                throw new ArgumentException(
                    $"Style-map '{path}' must be a JSON object mapping markdown element names to Word style names or StyleIds.");
            }

            // Accept both a flat mapping and the { "styleMap": {…} } wrapper shape.
            var mapping = root;
            if (root.TryGetProperty("styleMap", out var nested) && nested.ValueKind == JsonValueKind.Object)
            {
                mapping = nested;
            }

            foreach (var property in mapping.EnumerateObject())
            {
                if (property.Value.ValueKind != JsonValueKind.String)
                {
                    throw new ArgumentException(
                        $"Style-map '{path}': value for '{property.Name}' must be a string style name or StyleId.");
                }
                map[property.Name] = property.Value.GetString()!;
            }
        }

        return new StyleMapping { StyleMap = map };
    }

    static IReadOnlyList<MarkdownStyleDiagnostic> ResolveTemplateStyles(
        string templatePath,
        StyleMapping styleMap,
        bool strict,
        IReadOnlySet<string> referencedStyleKeys,
        bool userProvided)
    {
        var diagnostics = new List<MarkdownStyleDiagnostic>();
        using var document = WordprocessingDocument.Open(templatePath, false);
        var styles = document.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        var resolver = new MarkdownStyleResolver(
            styles,
            strict ? MarkdownStyleResolverOptions.StrictMode : MarkdownStyleResolverOptions.Default);

        foreach (var entry in styleMap.StyleMap)
        {
            // A user-supplied style map is validated wholesale: every entry is the author's
            // explicit intent, even for constructs the input happens not to use. The built-in
            // default map is validated only for the constructs the input actually renders, so a
            // default mapping for a construct that never appears cannot fail the conversion.
            if (!userProvided && !referencedStyleKeys.Contains(entry.Key))
            {
                continue;
            }

            var kind = MarkdownStyleKinds.ForElement(entry.Key);
            diagnostics.AddRange(resolver.Resolve(entry.Value, kind, element: entry.Key).Diagnostics);
        }

        return diagnostics;
    }

    static void PrintStyleDiagnostic(string sourcePath, MarkdownStyleDiagnostic diagnostic)
    {
        var location = string.IsNullOrEmpty(diagnostic.Element)
            ? sourcePath
            : $"{sourcePath} (element '{diagnostic.Element}')";
        var text = $"{location}: {diagnostic.Message}";
        if (diagnostic.Severity == MarkdownDiagnosticSeverity.Error)
        {
            AnsiConsole.MarkupLine($"[red]{Markup.Escape(text)}[/]");
        }
        else
        {
            AnsiConsole.MarkupLine($"[yellow]Warning: {Markup.Escape(text)}[/]");
        }
    }

    static void PrintMarkdownDiagnostic(string inputPath, MarkdownDiagnostic diagnostic)
    {
        var location = diagnostic.Span is { Line: >= 0 } span
            ? $"{inputPath}({span.Line + 1})"
            : inputPath;
        var text = $"{location}: {diagnostic.Message}";
        if (diagnostic.Severity == MarkdownDiagnosticSeverity.Error)
        {
            AnsiConsole.MarkupLine($"[red]{Markup.Escape(text)}[/]");
        }
        else
        {
            AnsiConsole.MarkupLine($"[yellow]Warning: {Markup.Escape(text)}[/]");
        }
    }

    static DocxEditor.Core.Models.DocumentInstructions LoadInstructions(string path)
    {
        var content = File.ReadAllText(path);

        if (path.EndsWith(".yaml", StringComparison.OrdinalIgnoreCase) ||
            path.EndsWith(".yml", StringComparison.OrdinalIgnoreCase))
        {
            var parser = new DocxYamlInstructionParser();
            return parser.Parse(content);
        }
        else
        {
            // Validate before parsing so unknown keys and missing/invalid fields abort
            // with field-level errors instead of being silently dropped at execution time.
            var validator = new DocxInstructionValidator();
            var errors = validator.Validate(content);
            if (errors.Count > 0)
            {
                throw new ArgumentException(
                    $"Invalid instruction file '{path}':{Environment.NewLine} - {string.Join(Environment.NewLine + " - ", errors)}");
            }

            var parser = new DocxJsonInstructionParser();
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
