namespace TypstBridge.Managed.Models;

/// <summary>
/// Managed compile request passed to the native Typst bridge.
/// </summary>
public sealed record TypstCompileRequest
{
    public TypstCompileRequest(
        string source,
        string workingDirectory,
        string rootFileName = "main.typ",
        IReadOnlyList<string>? fontPaths = null,
        TypstOutputFormat outputFormat = TypstOutputFormat.Pdf,
        double ppi = 144.0)
    {
        Source = source ?? throw new ArgumentNullException(nameof(source));
        WorkingDirectory = workingDirectory ?? throw new ArgumentNullException(nameof(workingDirectory));
        RootFileName = rootFileName ?? throw new ArgumentNullException(nameof(rootFileName));
        FontPaths = fontPaths is null ? [] : [.. fontPaths];
        OutputFormat = outputFormat;
        Ppi = ppi;
    }

    public string Source { get; }

    /// <summary>
    /// Directory used by the native bridge to resolve relative asset paths.
    /// </summary>
    public string WorkingDirectory { get; }

    public string RootFileName { get; }

    public IReadOnlyList<string> FontPaths { get; }

    public TypstOutputFormat OutputFormat { get; }

    /// <summary>
    /// Rasterization density for PNG output. Ignored by PDF and SVG output.
    /// </summary>
    public double Ppi { get; }
}
