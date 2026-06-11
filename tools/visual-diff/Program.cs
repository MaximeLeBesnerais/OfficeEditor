using System.Diagnostics;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

int exitCode = 0;
try
{
    CliOptions options = CliOptions.Parse(args);
    if (options.ShowHelp)
    {
        PrintUsage();
        return 0;
    }

    ToolPaths tools = ToolPaths.Resolve();
    List<ComparisonInput> inputs = BuildInputs(options);
    if (inputs.Count == 0)
    {
        throw new InvalidOperationException("No comparisons were selected.");
    }

    Directory.CreateDirectory(options.OutputDirectory);
    List<DocumentMetrics> documents = [];
    foreach (ComparisonInput input in inputs)
    {
        documents.Add(RunComparison(input, options.OutputDirectory, options.Dpi, tools));
    }

    MetricsReport report = new(
        GeneratedAt: DateTimeOffset.UtcNow,
        Dpi: options.Dpi,
        OutputDirectory: Path.GetFullPath(options.OutputDirectory),
        Documents: documents);

    WriteJson(Path.Combine(options.OutputDirectory, "metrics.json"), report);
    WriteHtml(Path.Combine(options.OutputDirectory, "index.html"), report);

    Console.WriteLine($"Visual diff report: {Path.GetFullPath(Path.Combine(options.OutputDirectory, "index.html"))}");
    Console.WriteLine($"Metrics JSON:       {Path.GetFullPath(Path.Combine(options.OutputDirectory, "metrics.json"))}");
}
catch (Exception ex)
{
    Console.Error.WriteLine($"visual-diff: {ex.Message}");
    Console.Error.WriteLine("Run 'visual-diff --help' for usage.");
    exitCode = 1;
}

return exitCode;

static List<ComparisonInput> BuildInputs(CliOptions options)
{
    if (options.Suite is not null)
    {
        if (!string.Equals(options.Suite, "docx", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException($"Unknown suite '{options.Suite}'. Supported suites: docx.");
        }

        return
        [
            new("Annual reporting template ENGLISH_0", "examples/REF/DOCX/Annual reporting template ENGLISH_0.pdf", "examples/output/ref/docx/Annual reporting template ENGLISH_0.pdf"),
            new("Monitoring Report Template", "examples/REF/DOCX/Monitoring Report Template.pdf", "examples/output/ref/docx/Monitoring Report Template.pdf"),
            new("-entreprise-bcp-pme", "examples/REF/DOCX/-entreprise-bcp-pme.pdf", "examples/output/ref/docx/-entreprise-bcp-pme.pdf")
        ];
    }

    if (options.ReferencePdf is null || options.GeneratedPdf is null)
    {
        throw new InvalidOperationException("Provide either '--suite docx' or both '--ref <reference.pdf>' and '--gen <generated.pdf>'.");
    }

    string name = options.Name ?? Path.GetFileNameWithoutExtension(options.GeneratedPdf);
    return [new(name, options.ReferencePdf, options.GeneratedPdf)];
}

static DocumentMetrics RunComparison(ComparisonInput input, string outputRoot, int dpi, ToolPaths tools)
{
    string referencePdf = Path.GetFullPath(input.ReferencePdf);
    string generatedPdf = Path.GetFullPath(input.GeneratedPdf);
    if (!File.Exists(referencePdf))
    {
        throw new FileNotFoundException($"Reference PDF is missing: {referencePdf}", referencePdf);
    }

    if (!File.Exists(generatedPdf))
    {
        throw new FileNotFoundException($"Generated PDF is missing: {generatedPdf}", generatedPdf);
    }

    string safeName = MakeSafeName(input.Name);
    string documentDir = Path.Combine(outputRoot, safeName);
    string referenceDir = Path.Combine(documentDir, "reference");
    string generatedDir = Path.Combine(documentDir, "generated");
    string diffDir = Path.Combine(documentDir, "diff");

    ResetDirectory(referenceDir);
    ResetDirectory(generatedDir);
    ResetDirectory(diffDir);

    Console.WriteLine($"Rendering '{input.Name}' at {dpi} DPI...");
    List<string> referencePages = RenderPdf(referencePdf, Path.Combine(referenceDir, "page"), dpi, tools);
    List<string> generatedPages = RenderPdf(generatedPdf, Path.Combine(generatedDir, "page"), dpi, tools);
    int comparedPageCount = Math.Min(referencePages.Count, generatedPages.Count);

    List<PageMetrics> pages = [];
    for (int i = 0; i < comparedPageCount; i++)
    {
        string diffImage = Path.Combine(diffDir, $"page-{i + 1:000}.png");
        pages.Add(ComparePage(i + 1, referencePages[i], generatedPages[i], diffImage, tools));
    }

    bool mismatch = referencePages.Count != generatedPages.Count;
    if (mismatch)
    {
        Console.WriteLine($"Page count mismatch for '{input.Name}': reference={referencePages.Count}, generated={generatedPages.Count}");
    }

    DocumentMetrics metrics = new(
        Name: input.Name,
        ReferencePdf: referencePdf,
        GeneratedPdf: generatedPdf,
        OutputDirectory: Path.GetFullPath(documentDir),
        ReferencePageCount: referencePages.Count,
        GeneratedPageCount: generatedPages.Count,
        PageCountMismatch: mismatch,
        Pages: pages);

    WriteJson(Path.Combine(documentDir, "metrics.json"), metrics);
    return metrics;
}

static List<string> RenderPdf(string pdfPath, string outputPrefix, int dpi, ToolPaths tools)
{
    Directory.CreateDirectory(Path.GetDirectoryName(outputPrefix)!);
    string renderer = tools.PdfToCairo ?? tools.PdfToPpm ?? throw new InvalidOperationException("No PDF renderer was found.");
    string prefixDirectory = Path.GetDirectoryName(outputPrefix)!;
    string prefixName = Path.GetFileName(outputPrefix);

    ProcessResult result = tools.PdfToCairo is not null
        ? RunProcess(renderer, ["-png", "-r", dpi.ToString(CultureInfo.InvariantCulture), pdfPath, outputPrefix])
        : RunProcess(renderer, ["-png", "-r", dpi.ToString(CultureInfo.InvariantCulture), pdfPath, outputPrefix]);

    if (result.ExitCode != 0)
    {
        throw new InvalidOperationException($"Failed to render PDF '{pdfPath}' with {Path.GetFileName(renderer)}: {result.ErrorOrOutput}");
    }

    return Directory.EnumerateFiles(prefixDirectory, $"{prefixName}-*.png")
        .OrderBy(NaturalPageKey, StringComparer.Ordinal)
        .ToList();
}

static PageMetrics ComparePage(int pageNumber, string referenceImage, string generatedImage, string diffImage, ToolPaths tools)
{
    ProcessResult result = RunProcess(tools.Compare, ["-metric", "RMSE", referenceImage, generatedImage, diffImage]);
    string metricText = result.ErrorOrOutput.Trim();
    (double? rmse, double? normalized) = ParseRmse(metricText);

    // ImageMagick compare returns 1 when images differ, so only 2+ is treated as an operational error.
    string? error = result.ExitCode > 1 ? metricText : null;
    if (error is not null && !File.Exists(diffImage))
    {
        File.Copy(referenceImage, diffImage, overwrite: true);
    }

    return new PageMetrics(
        PageNumber: pageNumber,
        ReferenceImage: referenceImage,
        GeneratedImage: generatedImage,
        DiffImage: diffImage,
        Rmse: rmse,
        NormalizedRmse: normalized,
        PercentRmse: normalized is null ? null : normalized * 100.0,
        CompareExitCode: result.ExitCode,
        RawMetric: metricText,
        Error: error);
}

static (double? Rmse, double? Normalized) ParseRmse(string text)
{
    Match match = Regex.Match(text, @"(?<rmse>[0-9]+(?:\.[0-9]+)?)\s*\((?<norm>[0-9]+(?:\.[0-9]+)?)\)");
    if (!match.Success)
    {
        return (null, null);
    }

    double rmse = double.Parse(match.Groups["rmse"].Value, CultureInfo.InvariantCulture);
    double normalized = double.Parse(match.Groups["norm"].Value, CultureInfo.InvariantCulture);
    return (rmse, normalized);
}

static void WriteJson<T>(string path, T value)
{
    JsonSerializerOptions options = new() { WriteIndented = true };
    File.WriteAllText(path, JsonSerializer.Serialize(value, options));
}

static void WriteHtml(string path, MetricsReport report)
{
    StringBuilder html = new();
    html.AppendLine("<!doctype html>");
    html.AppendLine("<html lang=\"en\"><head><meta charset=\"utf-8\"><title>PDF visual diff</title>");
    html.AppendLine("<style>body{font-family:sans-serif;margin:2rem}table{border-collapse:collapse;width:100%;margin-bottom:2rem}th,td{border:1px solid #ddd;padding:.5rem;vertical-align:top}img{max-width:220px;border:1px solid #ccc}.warn{color:#9a5b00;font-weight:bold}.err{color:#b00020;font-weight:bold}</style>");
    html.AppendLine("</head><body>");
    html.AppendLine("<h1>PDF visual diff</h1>");
    html.AppendLine($"<p>Generated: {Escape(report.GeneratedAt.ToString("u", CultureInfo.InvariantCulture))} · DPI: {report.Dpi}</p>");

    foreach (DocumentMetrics document in report.Documents)
    {
        html.AppendLine($"<h2>{Escape(document.Name)}</h2>");
        html.AppendLine($"<p>Reference pages: {document.ReferencePageCount}; Generated pages: {document.GeneratedPageCount}</p>");
        if (document.PageCountMismatch)
        {
            html.AppendLine("<p class=\"warn\">Page count mismatch: only common pages were compared.</p>");
        }

        html.AppendLine("<table><thead><tr><th>Page</th><th>Reference</th><th>Generated</th><th>Diff</th><th>Metrics</th></tr></thead><tbody>");
        foreach (PageMetrics page in document.Pages)
        {
            html.AppendLine("<tr>");
            html.AppendLine($"<td>{page.PageNumber}</td>");
            html.AppendLine(ImageCell(path, page.ReferenceImage));
            html.AppendLine(ImageCell(path, page.GeneratedImage));
            html.AppendLine(ImageCell(path, page.DiffImage));
            html.AppendLine("<td>");
            html.AppendLine($"RMSE: {FormatNullable(page.Rmse)}<br>");
            html.AppendLine($"Normalized: {FormatNullable(page.NormalizedRmse)}<br>");
            html.AppendLine($"Percent: {FormatNullable(page.PercentRmse)}%<br>");
            html.AppendLine($"Raw: {Escape(page.RawMetric)}");
            if (page.Error is not null)
            {
                html.AppendLine($"<div class=\"err\">{Escape(page.Error)}</div>");
            }
            html.AppendLine("</td></tr>");
        }

        html.AppendLine("</tbody></table>");
    }

    html.AppendLine("</body></html>");
    File.WriteAllText(path, html.ToString());
}

static string ImageCell(string htmlPath, string imagePath)
{
    string relative = Path.GetRelativePath(Path.GetDirectoryName(htmlPath)!, imagePath).Replace(Path.DirectorySeparatorChar, '/');
    string escaped = Escape(relative);
    return $"<td><a href=\"{escaped}\"><img src=\"{escaped}\" alt=\"{escaped}\"></a></td>";
}

static string FormatNullable(double? value) => value?.ToString("0.######", CultureInfo.InvariantCulture) ?? "n/a";

static string NaturalPageKey(string path)
{
    Match match = Regex.Match(Path.GetFileNameWithoutExtension(path), @"-(\d+)$");
    return match.Success ? int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture).ToString("D8", CultureInfo.InvariantCulture) : path;
}

static string MakeSafeName(string name)
{
    string normalized = Regex.Replace(name.ToLowerInvariant(), @"[^a-z0-9._-]+", "-").Trim('-');
    if (string.IsNullOrWhiteSpace(normalized))
    {
        normalized = "document";
    }

    string hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(name)))[..8].ToLowerInvariant();
    return $"{normalized}-{hash}";
}

static string Escape(string value) => System.Net.WebUtility.HtmlEncode(value);

static void ResetDirectory(string path)
{
    if (Directory.Exists(path))
    {
        Directory.Delete(path, recursive: true);
    }

    Directory.CreateDirectory(path);
}

static ProcessResult RunProcess(string executable, IReadOnlyList<string> arguments)
{
    ProcessStartInfo startInfo = new(executable)
    {
        RedirectStandardOutput = true,
        RedirectStandardError = true,
        UseShellExecute = false
    };

    foreach (string argument in arguments)
    {
        startInfo.ArgumentList.Add(argument);
    }

    using Process process = Process.Start(startInfo) ?? throw new InvalidOperationException($"Failed to start {executable}.");
    string stdout = process.StandardOutput.ReadToEnd();
    string stderr = process.StandardError.ReadToEnd();
    process.WaitForExit();
    return new ProcessResult(process.ExitCode, stdout, stderr);
}

static void PrintUsage()
{
    Console.WriteLine("""
visual-diff - render and compare PDF pages as PNG images.

Usage:
  visual-diff --suite docx [--out examples/output/visual-diff/docx] [--dpi 150]
  visual-diff --ref <reference.pdf> --gen <generated.pdf> --out <output-dir> [--name <name>] [--dpi 150]

Options:
  --suite <docx>       Compare the built-in DOCX reference suite.
  --ref <path>         Reference PDF for an arbitrary pair.
  --gen <path>         Generated PDF for an arbitrary pair.
  --out <path>         Output directory. Defaults to examples/output/visual-diff/docx for --suite docx.
  --name <name>        Display/report name for an arbitrary pair.
  --dpi <number>       Render DPI. Default: 150.
  --help               Show this help text.
""");
}

record ComparisonInput(string Name, string ReferencePdf, string GeneratedPdf);
record ProcessResult(int ExitCode, string Stdout, string Stderr)
{
    public string ErrorOrOutput => string.IsNullOrWhiteSpace(Stderr) ? Stdout : Stderr;
}

record MetricsReport(DateTimeOffset GeneratedAt, int Dpi, string OutputDirectory, List<DocumentMetrics> Documents);
record DocumentMetrics(string Name, string ReferencePdf, string GeneratedPdf, string OutputDirectory, int ReferencePageCount, int GeneratedPageCount, bool PageCountMismatch, List<PageMetrics> Pages);
record PageMetrics(int PageNumber, string ReferenceImage, string GeneratedImage, string DiffImage, double? Rmse, double? NormalizedRmse, double? PercentRmse, int CompareExitCode, string RawMetric, string? Error);

sealed record ToolPaths(string? PdfToCairo, string? PdfToPpm, string Compare)
{
    public static ToolPaths Resolve()
    {
        string? cairo = FindExecutable("pdftocairo");
        string? ppm = FindExecutable("pdftoppm");
        string? compare = FindExecutable("compare");

        if (cairo is null && ppm is null)
        {
            throw new InvalidOperationException("Missing PDF renderer: install pdftocairo or pdftoppm (Poppler). Tried PATH and /usr/bin.");
        }

        if (compare is null)
        {
            throw new InvalidOperationException("Missing ImageMagick compare executable. Install ImageMagick or ensure 'compare' is on PATH.");
        }

        return new ToolPaths(cairo, ppm, compare);
    }

    private static string? FindExecutable(string name)
    {
        string usrBin = Path.Combine("/usr/bin", name);
        if (File.Exists(usrBin))
        {
            return usrBin;
        }

        string? path = Environment.GetEnvironmentVariable("PATH");
        if (path is null)
        {
            return null;
        }

        foreach (string directory in path.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
        {
            string candidate = Path.Combine(directory, name);
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        return null;
    }
}

sealed record CliOptions(string? Suite, string? ReferencePdf, string? GeneratedPdf, string OutputDirectory, string? Name, int Dpi, bool ShowHelp)
{
    public static CliOptions Parse(string[] args)
    {
        string? suite = null;
        string? reference = null;
        string? generated = null;
        string? output = null;
        string? name = null;
        int dpi = 150;
        bool help = args.Length == 0;

        for (int i = 0; i < args.Length; i++)
        {
            string arg = args[i];
            switch (arg)
            {
                case "--help" or "-h":
                    help = true;
                    break;
                case "--suite":
                    suite = RequireValue(args, ref i, arg);
                    break;
                case "--ref":
                    reference = RequireValue(args, ref i, arg);
                    break;
                case "--gen":
                    generated = RequireValue(args, ref i, arg);
                    break;
                case "--out":
                    output = RequireValue(args, ref i, arg);
                    break;
                case "--name":
                    name = RequireValue(args, ref i, arg);
                    break;
                case "--dpi":
                    string dpiText = RequireValue(args, ref i, arg);
                    if (!int.TryParse(dpiText, NumberStyles.None, CultureInfo.InvariantCulture, out dpi) || dpi <= 0)
                    {
                        throw new InvalidOperationException("--dpi must be a positive integer.");
                    }
                    break;
                default:
                    throw new InvalidOperationException($"Unknown argument '{arg}'.");
            }
        }

        if (help)
        {
            return new CliOptions(suite, reference, generated, output ?? "examples/output/visual-diff/docx", name, dpi, help);
        }

        output ??= string.Equals(suite, "docx", StringComparison.OrdinalIgnoreCase)
            ? "examples/output/visual-diff/docx"
            : throw new InvalidOperationException("--out is required when comparing an arbitrary PDF pair.");

        return new CliOptions(suite, reference, generated, output, name, dpi, help);
    }

    private static string RequireValue(string[] args, ref int index, string option)
    {
        if (index + 1 >= args.Length || args[index + 1].StartsWith("--", StringComparison.Ordinal))
        {
            throw new InvalidOperationException($"{option} requires a value.");
        }

        index++;
        return args[index];
    }
}
