namespace VisualDiff;

/// <summary>
/// Builds the comparison inputs for the built-in suites and arbitrary
/// --ref/--gen pairs, and detects whether a run needs a PDF renderer.
/// </summary>
internal static class SuiteCatalog
{
    // NOTE: the suite now runs on the license-clean REF corpus (self-made decks).
    // The baselines under baselines/pptx/ were purged with the old third-party decks
    // and must be regenerated against this corpus.
    private static readonly (string Name, string Pptx, string RefPdf, string GenPdf)[] PptxDecks =
    [
        ("sales_acceleration_deck", "examples/REF/PPTX/sales_acceleration_deck.pptx", "examples/REF/PPTX/sales_acceleration_deck.pdf", "examples/output/ref/pptx/sales_acceleration_deck.pdf"),
        ("aetherlink-glass-shareholder-overview", "examples/REF/PPTX/AetherLink-Glass-Shareholder-Overview.pptx", "examples/REF/PPTX/AetherLink-Glass-Shareholder-Overview.pdf", "examples/output/ref/pptx/aetherlink-glass-shareholder-overview.pdf"),
        ("northwind-launch-review", "examples/REF/PPTX/northwind-launch-review.pptx", "examples/REF/PPTX/northwind-launch-review.pdf", "examples/output/ref/pptx/northwind-launch-review.pdf"),
        ("northwind-demo", "examples/REF/PPTX/northwind-demo.pptx", "examples/REF/PPTX/northwind-demo.pdf", "examples/output/ref/pptx/northwind-demo.pdf")
    ];

    // XLSX suite: "ours" is rendered by tools/convert-xlsx (Typst pipeline). The reference
    // PDFs are produced by LibreOffice headless — TEST-ONLY oracle, never product code.
    private static readonly (string Name, string Json, string RefPdf, string GenPdf)[] XlsxFixtures =
    [
        ("rich-report", "examples/Xlsx/instructions/rich-report.json", "examples/REF/XLSX/rich-report.pdf", "examples/output/ref/xlsx/rich-report.pdf"),
        ("complex-dashboard", "examples/Xlsx/instructions/complex-dashboard.json", "examples/REF/XLSX/complex-dashboard.pdf", "examples/output/ref/xlsx/complex-dashboard.pdf")
    ];

    public static List<ComparisonInput> BuildInputs(CliOptions options, ToolPaths tools)
    {
        if (options.Suite is not null)
        {
            if (string.Equals(options.Suite, "docx", StringComparison.OrdinalIgnoreCase))
            {
                return
                [
                    new("annual-report", "examples/REF/DOCX/annual-report.pdf", "examples/output/ref/docx/annual-report.pdf", InputKind.Pdf),
                    new("monitoring-report", "examples/REF/DOCX/monitoring-report.pdf", "examples/output/ref/docx/monitoring-report.pdf", InputKind.Pdf)
                ];
            }

            if (string.Equals(options.Suite, "pptx", StringComparison.OrdinalIgnoreCase))
            {
                return BuildPptxSuite(options);
            }

            if (string.Equals(options.Suite, "xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return BuildXlsxSuite(options);
            }

            if (string.Equals(options.Suite, "gen", StringComparison.OrdinalIgnoreCase))
            {
                return GenSuite.BuildInputs(options, tools);
            }

            throw new InvalidOperationException($"Unknown suite '{options.Suite}'. Supported suites: docx, pptx, xlsx, gen.");
        }

        InputKind referenceKind = DetectKind(options.ReferencePath!);
        InputKind generatedKind = DetectKind(options.GeneratedPath!);
        if (referenceKind != generatedKind)
        {
            throw new InvalidOperationException(
                "--ref and --gen must be the same kind of input: both PDFs, both single PNGs, or both directories of PNGs.");
        }

        string name = options.Name ?? Path.GetFileNameWithoutExtension(options.GeneratedPath!);
        if (string.IsNullOrWhiteSpace(name))
        {
            name = "comparison";
        }

        return [new(name, options.ReferencePath!, options.GeneratedPath!, referenceKind)];
    }

    /// <summary>
    /// Cheap, side-effect-free tool requirement detection used to probe external
    /// tools before any expensive work (e.g. --generate conversions) happens.
    /// </summary>
    public static ToolRequirements RequirementsFor(CliOptions options)
    {
        // The gen suite compares PNG-pair directories (PowerPoint ground truth vs Typst
        // preview), so it never needs a PDF renderer; --render additionally needs the
        // typst CLI.
        if (string.Equals(options.Suite, "gen", StringComparison.OrdinalIgnoreCase))
        {
            return ToolRequirements.ImageCompare | (options.RenderTypst ? ToolRequirements.Typst : ToolRequirements.None);
        }

        // The xlsx suite generates its PDFs via tools/convert-xlsx (Typst pipeline in-process),
        // so --generate needs no external typst CLI; only ImageMagick is required.
        if (string.Equals(options.Suite, "xlsx", StringComparison.OrdinalIgnoreCase))
        {
            return ToolRequirements.ImageCompare;
        }

        bool needsRenderer = options.Suite is not null
            || DetectKind(options.ReferencePath!) == InputKind.Pdf;
        return ToolRequirements.ImageCompare | (needsRenderer ? ToolRequirements.PdfRenderer : ToolRequirements.None);
    }

    public static InputKind DetectKind(string path)
    {
        if (Directory.Exists(path))
        {
            return InputKind.Png;
        }

        string extension = Path.GetExtension(path);
        if (extension.Equals(".png", StringComparison.OrdinalIgnoreCase))
        {
            return InputKind.Png;
        }

        if (extension.Equals(".pdf", StringComparison.OrdinalIgnoreCase))
        {
            return InputKind.Pdf;
        }

        throw new InvalidOperationException(
            $"Cannot determine input type of '{path}'. Use a .pdf file, a .png file, or a directory containing PNGs.");
    }

    private static List<ComparisonInput> BuildPptxSuite(CliOptions options)
    {
        List<ComparisonInput> inputs = [];
        List<string> skipped = [];

        foreach ((string name, string pptx, string refPdf, string genPdf) in PptxDecks)
        {
            if (!File.Exists(refPdf))
            {
                throw new FileNotFoundException(
                    $"Committed reference PDF is missing: {Path.GetFullPath(refPdf)}. The suite cannot run without it.", refPdf);
            }

            if (!File.Exists(genPdf))
            {
                if (options.GenerateMissing)
                {
                    if (!TryGeneratePptxPdf(name, pptx, genPdf, options.FontPath))
                    {
                        skipped.Add($"{name} (conversion failed)");
                        continue;
                    }
                }
                else
                {
                    skipped.Add(name);
                    Console.Error.WriteLine($"visual-diff: skipping '{name}': generated PDF not found: {genPdf}");
                    Console.Error.WriteLine($"  produce it with: {GenerateCommand(pptx, genPdf, options.FontPath)}");
                    Console.Error.WriteLine("  or re-run with --generate to let visual-diff build it via tools/convert-pptx.");
                    continue;
                }
            }

            inputs.Add(new(name, refPdf, genPdf, InputKind.Pdf));
        }

        if (inputs.Count == 0)
        {
            string commands = string.Join(Environment.NewLine,
                PptxDecks.Select(d => $"  {GenerateCommand(d.Pptx, d.GenPdf, options.FontPath)}"));
            throw new InvalidOperationException(
                "No generated PPTX PDFs were available under examples/output/ref/pptx/. Produce them first:\n"
                + commands
                + "\nOr re-run with --generate to let visual-diff invoke tools/convert-pptx itself.");
        }

        if (skipped.Count > 0)
        {
            Console.Error.WriteLine($"visual-diff: {skipped.Count} deck(s) skipped: {string.Join(", ", skipped)}");
        }

        return inputs;
    }

    private static string GenerateCommand(string pptx, string genPdf, string? fontPath) =>
        $"dotnet run --project tools/convert-pptx -- {pptx} {genPdf} --format pdf"
        + (fontPath is null ? string.Empty : $" --font-path {fontPath}");

    private static List<ComparisonInput> BuildXlsxSuite(CliOptions options)
    {
        List<ComparisonInput> inputs = [];
        List<string> skipped = [];

        foreach ((string name, string json, string refPdf, string genPdf) in XlsxFixtures)
        {
            if (!File.Exists(refPdf))
            {
                Console.Error.WriteLine($"visual-diff: skipping '{name}': committed reference PDF not found: {refPdf}");
                Console.Error.WriteLine("  Generate it once with LibreOffice (test-only oracle):");
                Console.Error.WriteLine($"  dotnet run --project OfficeEditor.Cli -- generate {json} --output /tmp/{name}.xlsx");
                Console.Error.WriteLine($"  soffice --headless --convert-to pdf --outdir examples/REF/XLSX /tmp/{name}.xlsx");
                skipped.Add(name);
                continue;
            }

            if (!File.Exists(genPdf))
            {
                if (options.GenerateMissing)
                {
                    if (!TryGenerateXlsxPdf(name, json, genPdf))
                    {
                        skipped.Add($"{name} (conversion failed)");
                        continue;
                    }
                }
                else
                {
                    skipped.Add(name);
                    Console.Error.WriteLine($"visual-diff: skipping '{name}': generated PDF not found: {genPdf}");
                    Console.Error.WriteLine($"  produce it with: {XlsxGenerateCommand(json, genPdf)}");
                    Console.Error.WriteLine("  or re-run with --generate to let visual-diff build it via tools/convert-xlsx.");
                    continue;
                }
            }

            inputs.Add(new(name, refPdf, genPdf, InputKind.Pdf));
        }

        if (inputs.Count == 0)
        {
            string commands = string.Join(Environment.NewLine,
                XlsxFixtures.Select(d => $"  {XlsxGenerateCommand(d.Json, d.GenPdf)}"));
            throw new InvalidOperationException(
                "No generated XLSX PDFs were available under examples/output/ref/xlsx/. Produce them first:\n"
                + commands
                + "\nOr re-run with --generate to let visual-diff invoke tools/convert-xlsx itself.");
        }

        if (skipped.Count > 0)
        {
            Console.Error.WriteLine($"visual-diff: {skipped.Count} fixture(s) skipped: {string.Join(", ", skipped)}");
        }

        return inputs;
    }

    private static string XlsxGenerateCommand(string json, string genPdf) =>
        $"dotnet run --project tools/convert-xlsx -- {json} {genPdf} --format pdf";

    private static bool TryGenerateXlsxPdf(string name, string json, string genPdf)
    {
        if (!File.Exists(json))
        {
            Console.Error.WriteLine($"visual-diff: cannot generate '{name}': source JSON missing: {json}");
            return false;
        }

        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(genPdf))!);
        List<string> arguments = ["run", "--project", "tools/convert-xlsx", "--", json, genPdf, "--format", "pdf"];

        Console.WriteLine($"Generating '{name}' PDF via tools/convert-xlsx...");
        try
        {
            ProcessResult result = ComparisonRunner.RunProcess("dotnet", arguments);
            if (result.ExitCode != 0 || !File.Exists(genPdf))
            {
                Console.Error.WriteLine($"visual-diff: convert-xlsx failed for '{name}' (exit {result.ExitCode}).");
                Console.Error.WriteLine(result.ErrorOrOutput);
                return false;
            }

            return true;
        }
        catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or InvalidOperationException)
        {
            Console.Error.WriteLine($"visual-diff: could not start 'dotnet' to generate '{name}': {ex.Message}");
            return false;
        }
    }

    private static bool TryGeneratePptxPdf(string name, string pptx, string genPdf, string? fontPath)
    {
        if (!File.Exists(pptx))
        {
            Console.Error.WriteLine($"visual-diff: cannot generate '{name}': source deck missing: {pptx}");
            return false;
        }

        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(genPdf))!);
        List<string> arguments = ["run", "--project", "tools/convert-pptx", "--", pptx, genPdf, "--format", "pdf"];
        if (fontPath is not null)
        {
            arguments.Add("--font-path");
            arguments.Add(fontPath);
        }

        Console.WriteLine($"Generating '{name}' PDF via tools/convert-pptx...");
        try
        {
            ProcessResult result = ComparisonRunner.RunProcess("dotnet", arguments);
            if (result.ExitCode != 0 || !File.Exists(genPdf))
            {
                Console.Error.WriteLine($"visual-diff: convert-pptx failed for '{name}' (exit {result.ExitCode}).");
                Console.Error.WriteLine(result.ErrorOrOutput);
                return false;
            }

            return true;
        }
        catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or InvalidOperationException)
        {
            Console.Error.WriteLine($"visual-diff: could not start 'dotnet' to generate '{name}': {ex.Message}");
            return false;
        }
    }
}
