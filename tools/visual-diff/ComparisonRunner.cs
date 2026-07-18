using System.Diagnostics;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace VisualDiff;

internal enum InputKind
{
    Pdf,
    Png
}

internal sealed record ComparisonInput(string Name, string ReferencePath, string GeneratedPath, InputKind Kind);

internal sealed record ProcessResult(int ExitCode, string Stdout, string Stderr)
{
    public string ErrorOrOutput => string.IsNullOrWhiteSpace(Stderr) ? Stdout : Stderr;
}

/// <summary>
/// Runs one document comparison: obtains page PNGs (by rendering PDFs with
/// poppler, or by collecting pre-rendered PNGs in PNG-pair mode), compares each
/// page with ImageMagick, and writes per-document metrics.
/// </summary>
internal static partial class ComparisonRunner
{
    public static DocumentMetrics Run(ComparisonInput input, string outputRoot, int dpi, ToolPaths tools)
    {
        string referencePath = Path.GetFullPath(input.ReferencePath);
        string generatedPath = Path.GetFullPath(input.GeneratedPath);
        ValidateInput(referencePath, "Reference", input.Kind);
        ValidateInput(generatedPath, "Generated", input.Kind);

        string safeName = MakeSafeName(input.Name);
        string documentDir = Path.Combine(outputRoot, safeName);
        string referenceDir = Path.Combine(documentDir, "reference");
        string generatedDir = Path.Combine(documentDir, "generated");
        string diffDir = Path.Combine(documentDir, "diff");

        ResetDirectory(referenceDir);
        ResetDirectory(generatedDir);
        ResetDirectory(diffDir);

        List<string> referencePages;
        List<string> generatedPages;
        if (input.Kind == InputKind.Png)
        {
            // PNG-pair mode: no poppler involved; inputs are copied into the
            // report tree under normalized page-NNN.png names so the diff and
            // HTML stages are identical to the PDF path.
            Console.WriteLine($"Collecting PNG pages for '{input.Name}'...");
            referencePages = CollectPngPages(referencePath, referenceDir);
            generatedPages = CollectPngPages(generatedPath, generatedDir);
        }
        else
        {
            Console.WriteLine($"Rendering '{input.Name}' at {dpi} DPI...");
            referencePages = RenderPdf(referencePath, Path.Combine(referenceDir, "page"), dpi, tools);
            generatedPages = RenderPdf(generatedPath, Path.Combine(generatedDir, "page"), dpi, tools);
        }

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
            ReferencePdf: referencePath,
            GeneratedPdf: generatedPath,
            OutputDirectory: Path.GetFullPath(documentDir),
            ReferencePageCount: referencePages.Count,
            GeneratedPageCount: generatedPages.Count,
            PageCountMismatch: mismatch,
            Pages: pages);

        ReportWriter.WriteJson(Path.Combine(documentDir, "metrics.json"), metrics);
        return metrics;
    }

    private static void ValidateInput(string path, string label, InputKind kind)
    {
        bool exists = kind == InputKind.Png
            ? File.Exists(path) || Directory.Exists(path)
            : File.Exists(path);
        if (!exists)
        {
            string what = kind == InputKind.Png ? "PNG file or directory" : "PDF";
            throw new FileNotFoundException($"{label} {what} is missing: {path}", path);
        }
    }

    private static List<string> CollectPngPages(string source, string destinationDir)
    {
        List<string> sources;
        if (File.Exists(source))
        {
            sources = [source];
        }
        else
        {
            sources = Directory.EnumerateFiles(source, "*.png", SearchOption.TopDirectoryOnly)
                .OrderBy(NaturalPageKey, StringComparer.Ordinal)
                .ToList();
            if (sources.Count == 0)
            {
                throw new InvalidOperationException($"No PNG files found in directory: {source}");
            }
        }

        List<string> pages = new(sources.Count);
        for (int i = 0; i < sources.Count; i++)
        {
            string destination = Path.Combine(destinationDir, $"page-{i + 1:000}.png");
            File.Copy(sources[i], destination, overwrite: true);
            pages.Add(destination);
        }

        return pages;
    }

    private static List<string> RenderPdf(string pdfPath, string outputPrefix, int dpi, ToolPaths tools)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(outputPrefix)!);
        string renderer = tools.PdfToCairo ?? tools.PdfToPpm
            ?? throw new InvalidOperationException("No PDF renderer was found. Install poppler or use PNG-pair inputs.");
        string prefixDirectory = Path.GetDirectoryName(outputPrefix)!;
        string prefixName = Path.GetFileName(outputPrefix);

        // pdftocairo and pdftoppm share the same -png -r <dpi> <pdf> <prefix> CLI shape.
        ProcessResult result = RunProcess(renderer, ["-png", "-r", dpi.ToString(CultureInfo.InvariantCulture), pdfPath, outputPrefix]);
        if (result.ExitCode != 0)
        {
            throw new InvalidOperationException($"Failed to render PDF '{pdfPath}' with {Path.GetFileName(renderer)}: {result.ErrorOrOutput}");
        }

        return Directory.EnumerateFiles(prefixDirectory, $"{prefixName}-*.png")
            .OrderBy(NaturalPageKey, StringComparer.Ordinal)
            .ToList();
    }

    private static PageMetrics ComparePage(int pageNumber, string referenceImage, string generatedImage, string diffImage, ToolPaths tools)
    {
        if (tools.Compare is null)
        {
            throw new InvalidOperationException("ImageMagick compare is required but was not found. Install ImageMagick (see --probe).");
        }

        string[] arguments = [.. tools.ComparePrefixArgs, "-metric", "RMSE", referenceImage, generatedImage, diffImage];
        ProcessResult result = RunProcess(tools.Compare, arguments);
        string metricText = result.ErrorOrOutput.Trim();
        (double? rmse, double? normalized) = RmseParser.Parse(metricText);

        // ImageMagick compare returns 1 when images differ, so only 2+ is treated as an operational error.
        string? error = result.ExitCode > 1 ? metricText : null;
        if (error is null && rmse is null)
        {
            // compare "succeeded" but produced no parseable metric — surface it instead of silently reporting nulls.
            error = $"Unparseable compare output: '{metricText}'";
        }

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

    internal static ProcessResult RunProcess(string executable, IReadOnlyList<string> arguments)
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

    private static string NaturalPageKey(string path)
    {
        // Zero-pad a trailing page number ("page-2" -> "00000002") so page-10
        // sorts after page-2; names without a trailing number sort by path.
        Match match = TrailingNumberPattern().Match(Path.GetFileNameWithoutExtension(path));
        return match.Success ? int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture).ToString("D8", CultureInfo.InvariantCulture) : path;
    }

    private static string MakeSafeName(string name)
    {
        string normalized = Regex.Replace(name.ToLowerInvariant(), @"[^a-z0-9._-]+", "-").Trim('-');
        if (string.IsNullOrWhiteSpace(normalized))
        {
            normalized = "document";
        }

        string hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(name)))[..8].ToLowerInvariant();
        return $"{normalized}-{hash}";
    }

    private static void ResetDirectory(string path)
    {
        if (Directory.Exists(path))
        {
            Directory.Delete(path, recursive: true);
        }

        Directory.CreateDirectory(path);
    }

    [GeneratedRegex(@"-(\d+)$")]
    private static partial Regex TrailingNumberPattern();
}
