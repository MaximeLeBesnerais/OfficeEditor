using System.Diagnostics;
using System.Text.Json;
using Xunit.Abstractions;

namespace DocxEditor.Tests.Unit;

public sealed class VisualDiffCliTests : IDisposable
{
    private readonly ITestOutputHelper output;
    private readonly string tempDirectory = Path.Combine(Path.GetTempPath(), "VisualDiffCliTests", Guid.NewGuid().ToString("N"));

    public VisualDiffCliTests(ITestOutputHelper output)
    {
        this.output = output;
        Directory.CreateDirectory(tempDirectory);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(tempDirectory))
            {
                Directory.Delete(tempDirectory, recursive: true);
            }
        }
        catch
        {
            // Best effort cleanup; test output is already written to the temp directory.
        }
    }

    [Fact]
    public void VisualDiff_WithIdenticalPdf_WritesReportAndZeroMetrics()
    {
        string repositoryRoot = GetRepositoryRoot();
        string referencePdf = Path.Combine(repositoryRoot, "examples", "REF", "DOCX", "Monitoring Report Template.pdf");
        if (!File.Exists(referencePdf))
        {
            output.WriteLine($"Skipping visual-diff CLI regression: reference PDF is missing: {referencePdf}");
            return;
        }

        if (FindExecutable("pdftocairo") is null && FindExecutable("pdftoppm") is null)
        {
            output.WriteLine("Skipping visual-diff CLI regression: missing Poppler renderer (pdftocairo or pdftoppm).");
            return;
        }

        if (FindExecutable("compare") is null)
        {
            output.WriteLine("Skipping visual-diff CLI regression: missing ImageMagick compare executable.");
            return;
        }

        string outputDirectory = Path.Combine(tempDirectory, "visual-diff-output");
        ProcessResult result = RunDotnet(
            repositoryRoot,
            "run",
            "--project", Path.Combine(repositoryRoot, "tools", "visual-diff"),
            "--",
            "--ref", referencePdf,
            "--gen", referencePdf,
            "--out", outputDirectory,
            "--name", "identical",
            "--dpi", "72");

        Assert.True(result.ExitCode == 0, $"visual-diff failed with exit code {result.ExitCode}.{Environment.NewLine}STDOUT:{Environment.NewLine}{result.Stdout}{Environment.NewLine}STDERR:{Environment.NewLine}{result.Stderr}");

        string indexPath = Path.Combine(outputDirectory, "index.html");
        string metricsPath = Path.Combine(outputDirectory, "metrics.json");
        Assert.True(File.Exists(indexPath), $"Expected report index at {indexPath}.");
        Assert.True(File.Exists(metricsPath), $"Expected metrics JSON at {metricsPath}.");

        using JsonDocument metrics = JsonDocument.Parse(File.ReadAllText(metricsPath));
        JsonElement document = Assert.Single(metrics.RootElement.GetProperty("Documents").EnumerateArray());
        Assert.Equal("identical", document.GetProperty("Name").GetString());
        Assert.False(document.GetProperty("PageCountMismatch").GetBoolean());

        JsonElement.ArrayEnumerator pages = document.GetProperty("Pages").EnumerateArray();
        Assert.NotEmpty(pages);
        foreach (JsonElement page in pages)
        {
            AssertMetricIsZero(page.GetProperty("Rmse"));
            AssertMetricIsZero(page.GetProperty("NormalizedRmse"));
        }

        string[] pngFiles = Directory.EnumerateFiles(outputDirectory, "*.png", SearchOption.AllDirectories).ToArray();
        Assert.Contains(pngFiles, path => path.Contains($"{Path.DirectorySeparatorChar}reference{Path.DirectorySeparatorChar}", StringComparison.Ordinal));
        Assert.Contains(pngFiles, path => path.Contains($"{Path.DirectorySeparatorChar}generated{Path.DirectorySeparatorChar}", StringComparison.Ordinal));
        Assert.Contains(pngFiles, path => path.Contains($"{Path.DirectorySeparatorChar}diff{Path.DirectorySeparatorChar}", StringComparison.Ordinal));
    }

    private static void AssertMetricIsZero(JsonElement metric)
    {
        Assert.Equal(JsonValueKind.Number, metric.ValueKind);
        Assert.InRange(Math.Abs(metric.GetDouble()), 0.0, 0.000001);
    }

    private static ProcessResult RunDotnet(string workingDirectory, params string[] arguments)
    {
        ProcessStartInfo startInfo = new("dotnet")
        {
            WorkingDirectory = workingDirectory,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false
        };

        foreach (string argument in arguments)
        {
            startInfo.ArgumentList.Add(argument);
        }

        using Process process = Process.Start(startInfo) ?? throw new InvalidOperationException("Failed to start dotnet.");
        string stdout = process.StandardOutput.ReadToEnd();
        string stderr = process.StandardError.ReadToEnd();
        if (!process.WaitForExit(TimeSpan.FromMinutes(2)))
        {
            process.Kill(entireProcessTree: true);
            throw new TimeoutException($"visual-diff CLI timed out.{Environment.NewLine}STDOUT:{Environment.NewLine}{stdout}{Environment.NewLine}STDERR:{Environment.NewLine}{stderr}");
        }

        return new ProcessResult(process.ExitCode, stdout, stderr);
    }

    private static string GetRepositoryRoot()
    {
        DirectoryInfo? directory = new(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "DocxEditor.sln")))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new InvalidOperationException("Could not locate repository root containing DocxEditor.sln.");
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

    private sealed record ProcessResult(int ExitCode, string Stdout, string Stderr);
}
