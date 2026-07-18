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
            return BaselineCheck.Run(options.CheckMetricsPath!, options.BaselinePath!, options.Margin);
    }

    // Run mode. Probe tools before any expensive work (e.g. --generate
    // conversions) so missing ImageMagick/poppler fails fast with a clear
    // message. PNG-pair inputs do not require poppler.
    ToolRequirements requirements = SuiteCatalog.RequirementsFor(options);
    ToolPaths tools = ToolPaths.Resolve(requirements);

    List<ComparisonInput> inputs = SuiteCatalog.BuildInputs(options);
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
    // threshold wrapper only engages when --baseline is passed.
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
