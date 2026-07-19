using System.Diagnostics;

namespace DocxEditor.Tests.Generation.Fixtures;

/// <summary>
/// End-to-end parity gate: generates all fixtures, renders the Typst preview via the
/// typst CLI and diffs every fixture against the PowerPoint ground truth, enforcing the
/// per-primitive thresholds through the real visual-diff entry point.
/// OPT-IN: set OE_RUN_PARITY_RENDER_TESTS=1 to enable. Skipped by default because
/// sandboxed environments have neither the typst CLI nor PowerPoint (the ground-truth
/// renders must exist — see tools/visual-diff/baselines/gen/README.md); the
/// deterministic generation + threshold plumbing above carries the acceptance weight
/// there, exactly like P5's compile tests.
/// </summary>
public sealed class ParityRenderTests
{
    private const string EnableEnvVar = "OE_RUN_PARITY_RENDER_TESTS";

    [Fact]
    public void GenSuite_AllFixtures_WithinPerPrimitiveThresholds()
    {
        if (Environment.GetEnvironmentVariable(EnableEnvVar) != "1")
        {
            return; // no typst CLI / PowerPoint ground truth in this environment (see class summary)
        }

        var repoRoot = FixtureCatalogTests.ResolveRepositoryRoot();
        var result = Run("dotnet",
            ["run", "--project", "tools/visual-diff", "--", "--suite", "gen", "--generate", "--render",
             "--thresholds", "tools/visual-diff/baselines/gen/thresholds.json"],
            repoRoot);

        Assert.True(result.ExitCode == 0,
            $"gen parity gate failed (exit {result.ExitCode}).\n{result.Stdout}\n{result.Stderr}");
    }

    private static (int ExitCode, string Stdout, string Stderr) Run(string executable, IReadOnlyList<string> arguments, string workingDirectory)
    {
        var startInfo = new ProcessStartInfo(executable)
        {
            WorkingDirectory = workingDirectory,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false
        };
        foreach (var argument in arguments)
        {
            startInfo.ArgumentList.Add(argument);
        }

        using var process = Process.Start(startInfo) ?? throw new InvalidOperationException($"Failed to start {executable}.");
        string stdout = process.StandardOutput.ReadToEnd();
        string stderr = process.StandardError.ReadToEnd();
        process.WaitForExit();
        return (process.ExitCode, stdout, stderr);
    }
}
