namespace VisualDiff;

/// <summary>
/// Which external tools a run needs. PNG-pair comparisons skip the PDF
/// renderer entirely, so poppler is only required when PDF inputs exist.
/// </summary>
[Flags]
internal enum ToolRequirements
{
    None = 0,
    PdfRenderer = 1,
    ImageCompare = 2,
    Typst = 4
}

/// <summary>
/// Resolved external tool paths. Probing order matches the original resolver
/// (fixed locations first, then PATH) and additionally covers the Homebrew
/// prefix on Apple Silicon (<c>/opt/homebrew/bin</c>). When ImageMagick 7's
/// legacy <c>compare</c> binary is absent but <c>magick</c> exists,
/// <see cref="ComparePrefixArgs"/> carries <c>["compare"]</c> so callers can
/// invoke <c>magick compare ...</c> transparently. <see cref="Typst"/> is the
/// typst CLI, only required by the gen suite's opt-in <c>--render</c> step.
/// </summary>
internal sealed record ToolPaths(
    string? PdfToCairo,
    string? PdfToPpm,
    string? Compare,
    IReadOnlyList<string> ComparePrefixArgs,
    string? Typst)
{
    private static readonly string[] FixedProbeDirectories = ["/usr/bin", "/opt/homebrew/bin"];

    public static ToolPaths Resolve(ToolRequirements requirements)
    {
        string? cairo = FindExecutable("pdftocairo");
        string? ppm = FindExecutable("pdftoppm");
        (string? compare, IReadOnlyList<string> prefix) = FindCompare();
        string? typst = FindExecutable("typst");

        if (requirements.HasFlag(ToolRequirements.PdfRenderer) && cairo is null && ppm is null)
        {
            throw new InvalidOperationException(
                "Missing PDF renderer: install poppler (macOS: 'brew install poppler'; "
                + "Debian/Ubuntu: 'apt install poppler-utils'), or use PNG-pair mode "
                + "('--ref/--gen' with .png files or directories), which does not need poppler. "
                + ProbedNote());
        }

        if (requirements.HasFlag(ToolRequirements.ImageCompare) && compare is null)
        {
            throw new InvalidOperationException(
                "Missing ImageMagick: install ImageMagick (macOS: 'brew install imagemagick'; "
                + "Debian/Ubuntu: 'apt install imagemagick') so that 'compare' (or 'magick') is available. "
                + ProbedNote());
        }

        if (requirements.HasFlag(ToolRequirements.Typst) && typst is null)
        {
            throw new InvalidOperationException(
                "--render needs the typst CLI: install it (macOS: 'brew install typst'; "
                + "see https://github.com/typst/typst for other platforms). "
                + "Without --render the gen suite still diffs any previously rendered PNG pages. "
                + ProbedNote());
        }

        return new ToolPaths(cairo, ppm, compare, prefix, typst);
    }

    /// <summary>Human-readable availability report for <c>--probe</c>.</summary>
    public static string ProbeReport()
    {
        string? cairo = FindExecutable("pdftocairo");
        string? ppm = FindExecutable("pdftoppm");
        (string? compare, IReadOnlyList<string> prefix) = FindCompare();

        string renderer = cairo is not null
            ? $"available (pdftocairo: {cairo})"
            : ppm is not null ? $"available (pdftoppm: {ppm})" : "UNAVAILABLE - PDF inputs cannot be rendered";
        string comparer = compare is not null
            ? $"available ({(prefix.Count > 0 ? "magick compare" : "compare")}: {compare})"
            : "UNAVAILABLE - no comparisons can run";

        return $"""
            visual-diff tool probe
              pdftocairo: {cairo ?? "MISSING"}
              pdftoppm:   {ppm ?? "MISSING"}
              compare:    {compare ?? "MISSING"}
              magick:     {FindExecutable("magick") ?? "MISSING"}
              typst:      {FindExecutable("typst") ?? "MISSING (only needed for --suite gen --render)"}
            Probed locations: fixed directories [{string.Join(", ", FixedProbeDirectories)}], then PATH.
            PDF rendering: {renderer}
            Image compare: {comparer}

            """;
    }

    private static (string? Exe, IReadOnlyList<string> Prefix) FindCompare()
    {
        string? compare = FindExecutable("compare");
        if (compare is not null)
        {
            return (compare, []);
        }

        // ImageMagick 7 distributions that only ship the 'magick' dispatcher.
        string? magick = FindExecutable("magick");
        return magick is not null ? (magick, ["compare"]) : (null, []);
    }

    private static string ProbedNote() =>
        $"Probed: {string.Join(", ", FixedProbeDirectories)}, then PATH.";

    private static string? FindExecutable(string name)
    {
        foreach (string directory in FixedProbeDirectories)
        {
            string candidate = Path.Combine(directory, name);
            if (File.Exists(candidate))
            {
                return candidate;
            }
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
