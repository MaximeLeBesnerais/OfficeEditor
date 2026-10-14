using OfficeEditor.Core.Generation;
using OfficeEditor.Core.Rendering;
using System.Text.Json;
using DocxEditor.Core.Builders;
using DocxEditor.Core.Generation;
using DocxEditor.Core.Generation.Assets;
using DocxEditor.Core.Generation.Design;
using DocxEditor.Core.Generation.Model;
using DocxEditor.Core.Generation.Schema;
using PptxEditor.Core.Builders;
using PptxEditor.Core.Generation;
using PptxEditor.Core.Generation.Schema;
using XlsxEditor.Core.Builders;
using XlsxEditor.Core.Instructions;
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
                case "inspect":
                    if (args.Length != 2) throw new ArgumentException("Usage: officeeditor inspect <source.json>");
                    Console.WriteLine(JsonSerializer.Serialize(GenerationJsonIds.Inspect(File.ReadAllText(args[1])), new JsonSerializerOptions(JsonSerializerDefaults.Web)));
                    return 0;
                case "ids":
                    return HandleIds(args) ? 0 : 1;
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

    static bool HandleIds(string[] args)
    {
        var flags = args.Skip(2).ToArray();
        if (args.Length < 2 || flags.Distinct().Count() != flags.Length ||
            flags.Any(flag => flag is not ("--recursive" or "--dry-run" or "--backup")))
            throw new ArgumentException("Usage: officeeditor ids <file.json|directory> [--recursive] [--dry-run] [--backup]");
        bool dryRun = flags.Contains("--dry-run"), backup = flags.Contains("--backup");
        var path = args[1];
        var files = Directory.Exists(path)
            ? Directory.EnumerateFiles(path, "*.json", new EnumerationOptions
            {
                RecurseSubdirectories = flags.Contains("--recursive"),
                AttributesToSkip = FileAttributes.ReparsePoint,
                IgnoreInaccessible = false
            }).Where(file => !Path.GetRelativePath(path, file).Split(Path.DirectorySeparatorChar)
                .Any(part => part is ".git" or "bin" or "obj" or "node_modules"))
            : new[] { path };
        // Preflight the entire batch before writing any file.
        var recognized = new List<string>();
        foreach (var file in files.Order(StringComparer.Ordinal))
        {
            var json = File.ReadAllText(file);
            if (!GenerationJsonIds.IsGenerationDocument(json)) continue;
            var result = GenerationJsonIds.Normalize(json);
            if (result.AddedCount == 0) continue;
            if (backup && !dryRun && File.Exists(file + ".bak"))
                throw new IOException($"Backup already exists: {file}.bak");
            recognized.Add(file);
        }
        int added = 0;
        foreach (var file in recognized)
        {
            if (backup && !dryRun) File.Copy(file, file + ".bak");
            added += dryRun ? GenerationJsonIds.Normalize(File.ReadAllText(file)).AddedCount
                : GenerationJsonIds.NormalizeFile(file).AddedCount;
            AnsiConsole.WriteLine(file);
        }
        AnsiConsole.WriteLine($"{(dryRun ? "Would normalize" : "Normalized")} {recognized.Count} generation JSON file(s); added {added} IDs.");
        return true;
    }

    static bool HandleGenerate(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Input JSON file path is required.[/]");
            AnsiConsole.MarkupLine("Usage: officeeditor generate <input.json> [--output <path>] [--theme <name>] [--no-write-ids] [--backup]");
            return false;
        }

        var inputPath = args[1];
        string? outputPath = null;
        string? theme = null;
        var noWriteIds = false;
        var backup = false;

        for (var i = 2; i < args.Length; i++)
        {
            var arg = args[i];
            if (arg.Equals("--output", StringComparison.OrdinalIgnoreCase))
            {
                if (i + 1 >= args.Length)
                {
                    AnsiConsole.MarkupLine("[red]--output requires a value: the output file path (for example out.pptx, out.docx, out.xlsx, out.pdf, out.png, or out.svg).[/]");
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
            else if (arg.Equals("--no-write-ids", StringComparison.OrdinalIgnoreCase))
            {
                noWriteIds = true;
            }
            else if (arg.Equals("--backup", StringComparison.OrdinalIgnoreCase))
            {
                backup = true;
            }
            else
            {
                AnsiConsole.MarkupLine($"[red]Unknown option: {Markup.Escape(arg)}[/]");
                return false;
            }
        }

        if (noWriteIds && backup)
        {
            AnsiConsole.MarkupLine("[red]--backup cannot be combined with --no-write-ids: the source JSON is not rewritten, so there is nothing to back up.[/]");
            return false;
        }

        var resolvedOutputPath = outputPath ?? Path.ChangeExtension(inputPath, ".pptx");
        var outputExtension = Path.GetExtension(resolvedOutputPath);
        if (!outputExtension.Equals(".pptx", StringComparison.OrdinalIgnoreCase) &&
            !outputExtension.Equals(".docx", StringComparison.OrdinalIgnoreCase) &&
            !outputExtension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) &&
            !outputExtension.Equals(".pdf", StringComparison.OrdinalIgnoreCase) &&
            !outputExtension.Equals(".png", StringComparison.OrdinalIgnoreCase) &&
            !outputExtension.Equals(".svg", StringComparison.OrdinalIgnoreCase))
        {
            AnsiConsole.MarkupLine("[red]Unsupported output format. Use a file path ending in .pptx, .docx, .xlsx, .pdf, .png, or .svg.[/]");
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
        if (GenerationJsonIds.IsGenerationDocument(json))
        {
            var normalized = GenerationJsonIds.Normalize(json);
            if (normalized.AddedCount > 0)
            {
                json = normalized.Json;
                if (!noWriteIds)
                {
                    if (backup && File.Exists(inputPath + ".bak"))
                        throw new IOException($"Backup already exists: {inputPath}.bak");
                    if (backup) File.Copy(inputPath, inputPath + ".bak");
                    GenerationJsonIds.NormalizeFile(inputPath);
                    AnsiConsole.MarkupLine($"[grey]Added {normalized.AddedCount} element IDs to {Markup.Escape(inputPath)}.[/]");
                }
            }
        }

        return WriteGeneratedDocument(inputPath, resolvedOutputPath, outputExtension, json, theme);
    }

    /// <summary>
    /// Writes <paramref name="resolvedOutputPath"/> from already-normalized generation JSON.
    /// Shared by <c>generate</c> and <c>create --instructions</c> so both surfaces use the same
    /// validation, error output and rendering pipeline.
    /// </summary>
    static bool WriteGeneratedDocument(string inputPath, string resolvedOutputPath, string outputExtension, string json, string? theme)
    {
        if (outputExtension.Equals(".docx", StringComparison.OrdinalIgnoreCase))
            return GenerateDocx(inputPath, resolvedOutputPath, json, theme);

        if (outputExtension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            return GenerateXlsx(resolvedOutputPath, json);

        if (outputExtension.Equals(".pptx", StringComparison.OrdinalIgnoreCase))
            return GeneratePptx(inputPath, resolvedOutputPath, json);

        var format = outputExtension.Equals(".png", StringComparison.OrdinalIgnoreCase) ? DocumentOutputFormat.Png
            : outputExtension.Equals(".svg", StringComparison.OrdinalIgnoreCase) ? DocumentOutputFormat.Svg
            : DocumentOutputFormat.Pdf;
        return GenerateRender(inputPath, resolvedOutputPath, format, json);
    }

    static bool GeneratePptx(string inputPath, string outputPath, string json)
    {
        var generated = false;

        AnsiConsole.Status()
            .Start("Generating...", ctx =>
            {
                var result = new PptxGenerator().Generate(json, new PptxGeneratorOptions
                {
                    DocumentDirectory = Path.GetDirectoryName(Path.GetFullPath(inputPath))
                });
                if (!result.Success)
                {
                    AnsiConsole.MarkupLine("[red]Validation errors:[/]");
                    foreach (var e in result.Errors)
                        AnsiConsole.MarkupLine($"  [red]{Markup.Escape(e.ToString())}[/]");
                    return;
                }

                File.WriteAllBytes(outputPath, result.PptxBytes!);
                foreach (var warning in result.Warnings.Select(w => w.ToString()).Concat(result.PipelineWarnings))
                    AnsiConsole.MarkupLine($"[yellow]Warning: {Markup.Escape(warning)}[/]");

                var warningCount = result.Warnings.Count + result.PipelineWarnings.Count;
                AnsiConsole.MarkupLine($"[green]{result.SlideCount} slides[/]  " +
                    $"[green]{result.PptxBytes!.Length} bytes[/]  " +
                    (warningCount == 0 ? "[green]0 warnings[/]" : $"[yellow]{warningCount} warnings[/]"));
                generated = true;
            });

        if (!generated)
            return false;

        AnsiConsole.MarkupLine($"[green]Saved: {Markup.Escape(outputPath)}[/]");
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

    /// <summary>
    /// Renders a generation JSON source to PDF/PNG/SVG through the shared render facade
    /// (<see cref="DocumentRenderer"/>), which dispatches .json sources by generation vocabulary.
    /// The source vocabulary is detected and validated first so a broken deck/report fails with
    /// its own format's diagnostics instead of another format's "invalid JSON" error.
    /// </summary>
    static bool GenerateRender(string inputPath, string outputPath, DocumentOutputFormat format, string json)
    {
        var vocabulary = DetectGenerationVocabulary(json);
        if (vocabulary == GenerationVocabulary.Unknown)
        {
            AnsiConsole.MarkupLine($"[red]Cannot render {Markup.Escape(inputPath)}: unrecognized generation JSON.[/]");
            AnsiConsole.MarkupLine("[red]Rendering supports generation JSON with a root 'slides' (PPTX v2.0), 'sections' (DOCX), or 'worksheets' (XLSX v1) collection.[/]");
            return false;
        }

        var errors = ValidateGenerationJson(vocabulary, json);
        if (errors.Count > 0)
        {
            AnsiConsole.MarkupLine($"[red]Cannot render {Markup.Escape(inputPath)}: invalid {GenerationVocabularyName(vocabulary)} generation JSON.[/]");
            foreach (var error in errors)
                AnsiConsole.MarkupLine($"  [red]{Markup.Escape(error)}[/]");
            AnsiConsole.MarkupLine($"[red]Fix the JSON, or write the native file with --output <path{GenerationVocabularyExtension(vocabulary)}>.[/]");
            return false;
        }

        DocumentRenderResult? result = null;
        AnsiConsole.Status()
            .Start($"Rendering {format.ToString().ToUpperInvariant()} via Typst...", _ =>
            {
                result = new DocumentRenderer().Render(new DocumentRenderRequest
                {
                    SourcePath = inputPath,
                    Format = format
                });
            });

        if (result is null || !result.Success)
        {
            AnsiConsole.MarkupLine($"[red]Rendering failed: {Markup.Escape(result?.ErrorMessage ?? "unknown error")}[/]");
            return false;
        }

        foreach (var warning in result.Warnings)
            AnsiConsole.MarkupLine($"[yellow]Warning: {Markup.Escape(warning)}[/]");

        var directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(directory))
            Directory.CreateDirectory(directory);

        if (format == DocumentOutputFormat.Pdf)
        {
            if (result.Pages.Length == 0)
            {
                AnsiConsole.MarkupLine("[red]Rendering produced no output.[/]");
                return false;
            }

            File.WriteAllBytes(outputPath, result.Pages[0]);
            AnsiConsole.MarkupLine($"[green]PDF saved: {Markup.Escape(outputPath)}[/] ({result.Pages.Length} page(s))");
            return true;
        }

        // PNG/SVG produce one buffer per page; write into a directory named by the output path
        // (matching the repository's convert tools).
        var pageExtension = format == DocumentOutputFormat.Png ? "png" : "svg";
        Directory.CreateDirectory(outputPath);
        for (int i = 0; i < result.Pages.Length; i++)
        {
            string pagePath = Path.Combine(outputPath, $"page-{i + 1:D3}.{pageExtension}");
            File.WriteAllBytes(pagePath, result.Pages[i]);
        }

        AnsiConsole.MarkupLine($"[green]{result.Pages.Length} {pageExtension.ToUpperInvariant()} page(s) saved to: {Markup.Escape(outputPath)}[/]");
        return true;
    }

    /// <summary>
    /// Classifies a .json source by its root collection. The order mirrors the shared render
    /// facade's JSON probe (XLSX, then DOCX, then PPTX) so this validation and the facade's own
    /// dispatch agree; the three vocabularies reject each other's root collections outright.
    /// </summary>
    static GenerationVocabulary DetectGenerationVocabulary(string json)
    {
        try
        {
            using var document = JsonDocument.Parse(json);
            if (document.RootElement.ValueKind != JsonValueKind.Object)
                return GenerationVocabulary.Unknown;

            var collections = document.RootElement.EnumerateObject()
                .Where(property => property.Value.ValueKind == JsonValueKind.Array)
                .Select(property => property.Name)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            if (collections.Contains("worksheets")) return GenerationVocabulary.Xlsx;
            if (collections.Contains("sections")) return GenerationVocabulary.Docx;
            if (collections.Contains("slides")) return GenerationVocabulary.Pptx;
        }
        catch (JsonException)
        {
            // Unparseable JSON is reported by the format validators below as a precise error.
        }

        return GenerationVocabulary.Unknown;
    }

    /// <summary>
    /// Validates generation JSON with the detected format's own parser — the same validators the
    /// render facade probes with — and returns their diagnostics.
    /// </summary>
    static IReadOnlyList<string> ValidateGenerationJson(GenerationVocabulary vocabulary, string json)
    {
        switch (vocabulary)
        {
            case GenerationVocabulary.Xlsx:
                return XlsxInstructionParser.ParseAndValidate(json).Validation.Errors
                    .Select(diagnostic => string.IsNullOrEmpty(diagnostic.Path)
                        ? diagnostic.Message
                        : $"{diagnostic.Message} at {diagnostic.Path}")
                    .ToList();
            case GenerationVocabulary.Docx:
                return new DocxGenerationDocumentParser().Validate(json).Errors
                    .Select(error => error.ToString())
                    .ToList();
            case GenerationVocabulary.Pptx:
                return new GenerationDocumentParser().Validate(json).Errors
                    .Select(error => error.ToString())
                    .ToList();
            default:
                return Array.Empty<string>();
        }
    }

    static string GenerationVocabularyName(GenerationVocabulary vocabulary) => vocabulary switch
    {
        GenerationVocabulary.Pptx => "PPTX",
        GenerationVocabulary.Docx => "DOCX",
        GenerationVocabulary.Xlsx => "XLSX",
        _ => "unrecognized"
    };

    static string GenerationVocabularyExtension(GenerationVocabulary vocabulary) => vocabulary switch
    {
        GenerationVocabulary.Pptx => ".pptx",
        GenerationVocabulary.Docx => ".docx",
        GenerationVocabulary.Xlsx => ".xlsx",
        _ => ".json"
    };

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
        AnsiConsole.WriteLine("  officeeditor create <output.docx|output.pptx|output.xlsx> [--text \"content\"] [--style \"StyleId\"] [--title \"title\"] [--sheet \"name\"]");
        AnsiConsole.WriteLine("  officeeditor create <output.docx|output.pptx|output.xlsx> --instructions <input.json>");
        AnsiConsole.WriteLine("  officeeditor inspect <source.json>");
        AnsiConsole.WriteLine("  officeeditor edit <source.json> --instructions <operations.json> --output <edited.json>");
        AnsiConsole.WriteLine("  officeeditor ids <file.json|directory> [--recursive] [--dry-run] [--backup]");
        AnsiConsole.WriteLine("  officeeditor generate <input.json> [--output <path>] [--theme <name>] [--no-write-ids] [--backup]");
        AnsiConsole.WriteLine("  officeeditor detect <template.file>");
        AnsiConsole.WriteLine("  officeeditor merge <template.file> <data.json> <output.file>");
        AnsiConsole.WriteLine();
        AnsiConsole.WriteLine("create options:");
        AnsiConsole.WriteLine("  --instructions <input.json>  Generate the document from generation JSON instead of a blank");
        AnsiConsole.WriteLine("                               document. The output extension selects the vocabulary:");
        AnsiConsole.WriteLine("                               .docx DOCX generation JSON, .pptx PPTX v2.0 deck JSON,");
        AnsiConsole.WriteLine("                               .xlsx XLSX v1 instruction set. Validation and errors are");
        AnsiConsole.WriteLine("                               reported exactly as `generate` reports them; the instruction");
        AnsiConsole.WriteLine("                               file is never rewritten.");
        AnsiConsole.WriteLine("  --text <text>                Blank DOCX only: paragraph text (default \"Hello World\").");
        AnsiConsole.WriteLine("  --style <StyleId>            Blank DOCX only: paragraph style id.");
        AnsiConsole.WriteLine("  --title <title>              Blank PPTX only: title of the single generated slide.");
        AnsiConsole.WriteLine("  --sheet <name>               Blank XLSX only: worksheet name (default \"Sheet1\").");
        AnsiConsole.WriteLine();
        AnsiConsole.WriteLine("generate options:");
        AnsiConsole.WriteLine("  --output <path>  Output file. The extension selects the output:");
        AnsiConsole.WriteLine("                   .pptx  generate PPTX from PPTX v2.0 deck JSON (root \"slides\")");
        AnsiConsole.WriteLine("                   .docx  generate DOCX from DOCX generation JSON (root \"sections\")");
        AnsiConsole.WriteLine("                   .xlsx  generate XLSX from XLSX v1 instruction JSON (root \"worksheets\")");
        AnsiConsole.WriteLine("                   .pdf   render to one PDF file");
        AnsiConsole.WriteLine("                   .png   render one image per page into a directory named by the path");
        AnsiConsole.WriteLine("                   .svg   render one image per page into a directory named by the path");
        AnsiConsole.WriteLine("                          (page-001.png / page-001.svg, ...)");
        AnsiConsole.WriteLine("                   .pdf/.png/.svg accept any of the three generation vocabularies and never");
        AnsiConsole.WriteLine("                   fall back to another format. Defaults to <input.json> with .pptx.");
        AnsiConsole.WriteLine("  --theme <name>   DOCX only: named design theme to apply (editorial|corporate).");
        AnsiConsole.WriteLine("                   Requires a .docx output path.");
        AnsiConsole.WriteLine("  --no-write-ids   Do not write generated element IDs back into <input.json>.");
        AnsiConsole.WriteLine("  --backup         With ID write-back, copy <input.json> to <input.json>.bak first");
        AnsiConsole.WriteLine("                   (fails when that backup already exists).");
        AnsiConsole.WriteLine();
        AnsiConsole.WriteLine("Generation vocabularies: PPTX v2.0 (\"slides\"), DOCX (\"sections\"), XLSX v1 (\"worksheets\").");
        AnsiConsole.WriteLine("Native output (.pptx/.docx/.xlsx) requires the matching vocabulary; .pdf/.png/.svg accept all three.");
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

        string? instructionsPath = null;
        for (var i = 2; i < args.Length; i++)
        {
            if (!args[i].Equals("--instructions", StringComparison.OrdinalIgnoreCase))
                continue;
            if (i + 1 >= args.Length)
            {
                AnsiConsole.MarkupLine("[red]--instructions requires a value: the generation JSON file path.[/]");
                return false;
            }
            instructionsPath = args[++i];
        }

        if (instructionsPath is not null)
            return CreateFromInstructions(format, outputPath, instructionsPath);

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

    /// <summary>
    /// Generates <paramref name="outputPath"/> from generation JSON through the same validated
    /// pipeline <c>generate</c> uses. The instruction file is read-only (element IDs are
    /// allocated in memory only), and every validation/generator error is reported before any
    /// output is written, so a rejected instruction file can never leave a blank document behind.
    /// </summary>
    static bool CreateFromInstructions(DocumentFormat format, string outputPath, string instructionsPath)
    {
        if (!File.Exists(instructionsPath))
        {
            AnsiConsole.MarkupLine($"[red]File not found: {Markup.Escape(instructionsPath)}[/]");
            return false;
        }

        var json = File.ReadAllText(instructionsPath);
        return format switch
        {
            DocumentFormat.Docx => GenerateDocx(instructionsPath, outputPath, json, theme: null),
            DocumentFormat.Pptx => GeneratePptx(instructionsPath, outputPath, json),
            DocumentFormat.Xlsx => GenerateXlsx(outputPath, json),
            _ => throw new ArgumentException($"--instructions does not support '{outputPath}'. Use .docx, .pptx, or .xlsx.")
        };
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
        if (args.Length != 6 || args[2] != "--instructions" || args[4] != "--output")
            throw new ArgumentException("Usage: officeeditor edit <source.json> --instructions <operations.json> --output <edited.json>");
        if (Path.GetFullPath(args[1]) == Path.GetFullPath(args[5]))
            throw new ArgumentException("Use a separate output file to preserve the original source.");
        var edited = GenerationJsonEditor.Apply(File.ReadAllText(args[1]), File.ReadAllText(args[3]));
        // CreateNew prevents accidentally replacing an existing source or output.
        using var output = new FileStream(args[5], FileMode.CreateNew, FileAccess.Write);
        using var writer = new StreamWriter(output);
        writer.Write(edited);
        return true;
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

/// <summary>The generation JSON vocabulary a source document declares by its root collection.</summary>
enum GenerationVocabulary
{
    Unknown,
    Pptx,
    Docx,
    Xlsx
}
