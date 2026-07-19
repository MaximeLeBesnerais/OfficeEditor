using VisualDiff;

try
{
    CliOptions options = CliOptions.Parse(args);
    switch (options.Mode)
    {
        case CliMode.Help:
            CliOptions.PrintUsage();
            return 0;
        case CliMode.Probe:
            Console.Write(ToolPaths.ProbeReport());
            return 0;
        case CliMode.Check:
            return options.ThresholdsPath is not null
                ? ThresholdCheck.Run(options.CheckMetricsPath!, options.ThresholdsPath)
                : BaselineCheck.Run(options.CheckMetricsPath!, options.BaselinePath!, options.Margin);
    }

    // Run mode. The gen suite gates on per-primitive thresholds by default: the
    // committed thresholds file applies unless the caller overrode the check.
    if (string.Equals(options.Suite, "gen", StringComparison.OrdinalIgnoreCase)
        && options.ThresholdsPath is null
        && options.BaselinePath is null
        && File.Exists(GenSuite.DefaultThresholdsPath))
    {
        options = options with { ThresholdsPath = GenSuite.DefaultThresholdsPath };
    }

    // Probe tools before any expensive work (e.g. --generate conversions) so missing
    // ImageMagick/poppler fails fast with a clear message. PNG-pair inputs (which the
    // gen suite always uses) do not require poppler.
    ToolRequirements requirements = SuiteCatalog.RequirementsFor(options);
    ToolPaths tools = ToolPaths.Resolve(requirements);

    List<ComparisonInput> inputs = SuiteCatalog.BuildInputs(options, tools);
    if (inputs.Count == 0)
    {
        throw new InvalidOperationException("No comparisons were selected.");
    }

    Directory.CreateDirectory(options.OutputDirectory);
    List<DocumentMetrics> documents = [];
    foreach (ComparisonInput input in inputs)
    {
        documents.Add(ComparisonRunner.Run(input, options.OutputDirectory, options.Dpi, tools));
    }

    MetricsReport report = new(
        GeneratedAt: DateTimeOffset.UtcNow,
        Dpi: options.Dpi,
        OutputDirectory: Path.GetFullPath(options.OutputDirectory),
        Documents: documents);

    string metricsPath = Path.Combine(options.OutputDirectory, "metrics.json");
    ReportWriter.WriteJson(metricsPath, report);
    ReportWriter.WriteHtml(Path.Combine(options.OutputDirectory, "index.html"), report);

    Console.WriteLine($"Visual diff report: {Path.GetFullPath(Path.Combine(options.OutputDirectory, "index.html"))}");
    Console.WriteLine($"Metrics JSON:       {Path.GetFullPath(metricsPath)}");

    // Report-only by default (exit 0 regardless of RMSE). The optional
    // threshold wrappers only engage when --baseline or --thresholds is passed.
    if (options.ThresholdsPath is not null)
    {
        return ThresholdCheck.Run(metricsPath, options.ThresholdsPath);
    }

    if (options.BaselinePath is not null)
    {
        return BaselineCheck.Run(metricsPath, options.BaselinePath, options.Margin);
    }

    return 0;
}
catch (Exception ex)
{
    Console.Error.WriteLine($"visual-diff: {ex.Message}");
    Console.Error.WriteLine("Run 'visual-diff --help' for usage.");
    return 1;
}
