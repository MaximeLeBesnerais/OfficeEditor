using System.Diagnostics;
using System.Runtime.InteropServices;
using TypstBridge.Managed;

namespace PptxBenchmark;

/// <summary>Environment facts recorded in the benchmark report.</summary>
public sealed record EnvironmentInfo(
    string Os,
    string Architecture,
    string Runtime,
    string Sdk,
    string TypstBridgeVersion)
{
    public static EnvironmentInfo Collect()
    {
        return new EnvironmentInfo(
            RuntimeInformation.OSDescription,
            RuntimeInformation.OSArchitecture.ToString(),
            RuntimeInformation.FrameworkDescription,
            QueryDotnetSdkVersion(),
            ProbeTypstBridgeVersion());
    }

    private static string QueryDotnetSdkVersion()
    {
        try
        {
            var psi = new ProcessStartInfo("dotnet", "--version")
            {
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            using var process = Process.Start(psi)!;
            var stdoutTask = process.StandardOutput.ReadToEndAsync();
            var stderrTask = process.StandardError.ReadToEndAsync();
            if (!process.WaitForExitAsync().Wait(TimeSpan.FromSeconds(30)) || process.ExitCode != 0)
            {
                return "unknown";
            }
            _ = stderrTask;
            return stdoutTask.GetAwaiter().GetResult().Trim();
        }
        catch
        {
            return "unknown";
        }
    }

    private static string ProbeTypstBridgeVersion()
    {
        try
        {
            var compiler = new TypstBridgeCompiler();
            return compiler.Probe() ? compiler.Version : "unavailable (native probe failed)";
        }
        catch (Exception ex)
        {
            return $"unavailable ({ex.Message})";
        }
    }
}
