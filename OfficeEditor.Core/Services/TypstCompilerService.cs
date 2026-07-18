using System.Collections.Concurrent;
using System.Diagnostics;
using System.Globalization;
using TypstBridge.Managed;
using TypstBridge.Managed.Models;

namespace OfficeEditor.Core.Services;

/// <summary>
/// Timing data for a single <see cref="TypstCompilerService.Compile"/> call.
/// Only recorded when <see cref="TypstCompilerService.TimingEnabled"/> is set
/// (or the OFFICEEDITOR_TIMING environment variable is truthy at process start).
/// </summary>
public sealed record TypstCompileTiming
{
    public required DateTimeOffset TimestampUtc { get; init; }
    public required OutputFormat Format { get; init; }
    /// <summary>Backend that produced the result: "bridge", "legacy-typstsharp", or "cli".</summary>
    public required string Backend { get; init; }
    public required double TotalMilliseconds { get; init; }
    /// <summary>Time spent in the TypstBridge attempt (always tried first).</summary>
    public required double BridgeMilliseconds { get; init; }
    /// <summary>Time spent in the fallback backend, when one was used.</summary>
    public double? FallbackMilliseconds { get; init; }
    public required int PageCount { get; init; }
    public required bool Success { get; init; }
}

public enum OutputFormat
{
    Pdf,
    Png,
    Svg
}

public sealed record CompileOptions
{
    public OutputFormat Format { get; init; } = OutputFormat.Pdf;
    public float Ppi { get; init; } = 150;
    public string? FontDirectory { get; init; }
    public string? WorkingDirectory { get; init; }
    public TimeSpan ProcessTimeout { get; init; } = TimeSpan.FromMinutes(2);
}

public sealed record CompileResult
{
    public required byte[][] Pages { get; init; }
    public required bool Success { get; init; }
    public string? ErrorMessage { get; init; }
}

public sealed class TypstCompilerService : IDisposable
{
    private bool _disposed;
    private static readonly bool _legacyTypstSharpAvailable;
    private dynamic? _legacyTypstSharpCompiler; // legacy typstsharp PDF fallback compiler

    // Opt-in compile timing (OFFICEEDITOR_TIMING env var or TimingEnabled property).
    // When off there is no allocation and no clock read on the Compile path.
    private static readonly ConcurrentQueue<TypstCompileTiming> _timings = new();
    private static volatile bool _timingEnabled = IsTimingEnvironmentVariableSet();

    /// <summary>
    /// Enables per-compile timing records. Defaults to the OFFICEEDITOR_TIMING
    /// environment variable ("1"/"true"/"yes", case-insensitive).
    /// </summary>
    public static bool TimingEnabled
    {
        get => _timingEnabled;
        set => _timingEnabled = value;
    }

    /// <summary>
    /// Removes and returns all timing records collected so far. Thread-safe.
    /// </summary>
    public static IReadOnlyList<TypstCompileTiming> DrainTimings()
    {
        var drained = new List<TypstCompileTiming>();
        while (_timings.TryDequeue(out var timing))
        {
            drained.Add(timing);
        }
        return drained;
    }

    private static bool IsTimingEnvironmentVariableSet()
    {
        var value = Environment.GetEnvironmentVariable("OFFICEEDITOR_TIMING");
        return string.Equals(value, "1", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "yes", StringComparison.OrdinalIgnoreCase);
    }

    static TypstCompilerService()
    {
        // Probe whether the legacy typstsharp fallback can be loaded for PDF output.
        try
        {
            var compilerType = Type.GetType("typstsharp.TypstCompiler, typstsharp");
            if (compilerType != null)
            {
                var method = compilerType.GetMethod("FromSource", new[] { typeof(string) });
                if (method != null)
                {
                    using var probe = (IDisposable?)method.Invoke(null, new object[] { "#text[probe]" });
                    _legacyTypstSharpAvailable = probe != null;
                }
            }
        }
        catch
        {
            _legacyTypstSharpAvailable = false;
        }
    }

    public CompileResult Compile(string source, CompileOptions? options = null)
    {
        options ??= new CompileOptions();

        // Timers stay null when timing is disabled — zero overhead on the hot path.
        var totalTimer = _timingEnabled ? Stopwatch.StartNew() : null;
        var bridgeTimer = _timingEnabled ? Stopwatch.StartNew() : null;

        var bridgeResult = CompileBridge(source, options);
        bridgeTimer?.Stop();
        if (bridgeResult.Success)
        {
            RecordTiming(totalTimer, bridgeTimer, fallbackTimer: null, "bridge", options.Format, bridgeResult);
            return bridgeResult;
        }

        var useLegacyFallback = _legacyTypstSharpAvailable && options.Format == OutputFormat.Pdf && string.IsNullOrEmpty(options.FontDirectory);
        var fallbackTimer = _timingEnabled ? Stopwatch.StartNew() : null;
        var fallbackResult = useLegacyFallback
            ? CompileLegacyTypstSharp(source, options)
            : CompileCli(source, options);
        fallbackTimer?.Stop();

        if (!fallbackResult.Success)
        {
            fallbackResult = fallbackResult with
            {
                ErrorMessage = CombineErrors(bridgeResult.ErrorMessage, fallbackResult.ErrorMessage)
            };
        }

        // Note: CompileLegacyTypstSharp may internally degrade to the CLI on error,
        // so "legacy-typstsharp" marks the entry path, not necessarily the final backend.
        RecordTiming(totalTimer, bridgeTimer, fallbackTimer, useLegacyFallback ? "legacy-typstsharp" : "cli", options.Format, fallbackResult);
        return fallbackResult;
    }

    private static void RecordTiming(Stopwatch? totalTimer, Stopwatch? bridgeTimer, Stopwatch? fallbackTimer, string backend, OutputFormat format, CompileResult result)
    {
        if (totalTimer is null)
        {
            return;
        }

        totalTimer.Stop();
        _timings.Enqueue(new TypstCompileTiming
        {
            TimestampUtc = DateTimeOffset.UtcNow,
            Format = format,
            Backend = backend,
            TotalMilliseconds = totalTimer.Elapsed.TotalMilliseconds,
            BridgeMilliseconds = bridgeTimer?.Elapsed.TotalMilliseconds ?? 0,
            FallbackMilliseconds = fallbackTimer?.Elapsed.TotalMilliseconds,
            PageCount = result.Pages.Length,
            Success = result.Success
        });
    }

    private static CompileResult CompileBridge(string source, CompileOptions options)
    {
        try
        {
            var compiler = new TypstBridgeCompiler();
            if (!compiler.Probe())
            {
                return Failure("TypstBridge native library is unavailable.");
            }

            var request = new TypstCompileRequest(
                source,
                GetBridgeWorkingDirectory(options),
                fontPaths: GetBridgeFontPaths(options),
                outputFormat: MapOutputFormat(options.Format),
                ppi: options.Ppi);

            var result = compiler.Compile(request);
            if (!result.Success)
            {
                return Failure(CreateBridgeErrorMessage(result));
            }

            var pages = result.Outputs
                .OrderBy(output => output.PageIndex)
                .Select(output => output.Data)
                .ToArray();

            if (pages.Length == 0)
            {
                return Failure("TypstBridge produced no output.");
            }

            return new CompileResult
            {
                Pages = pages,
                Success = true
            };
        }
        catch (Exception ex)
        {
            return Failure($"TypstBridge compilation failed: {ex.Message}");
        }
    }

    private static TypstOutputFormat MapOutputFormat(OutputFormat format) => format switch
    {
        OutputFormat.Png => TypstOutputFormat.Png,
        OutputFormat.Svg => TypstOutputFormat.Svg,
        OutputFormat.Pdf => TypstOutputFormat.Pdf,
        _ => TypstOutputFormat.Pdf
    };

    private static string GetBridgeWorkingDirectory(CompileOptions options)
    {
        if (!string.IsNullOrWhiteSpace(options.WorkingDirectory))
        {
            return options.WorkingDirectory;
        }

        var currentDirectory = Directory.GetCurrentDirectory();
        return string.IsNullOrWhiteSpace(currentDirectory) ? Path.GetTempPath() : currentDirectory;
    }

    private static IReadOnlyList<string> GetBridgeFontPaths(CompileOptions options)
    {
        return string.IsNullOrWhiteSpace(options.FontDirectory)
            ? Array.Empty<string>()
            : options.FontDirectory
                .Split(Path.PathSeparator)
                .Select(path => path.Trim())
                .Where(path => path.Length > 0)
                .ToArray();
    }

    private static CompileResult Failure(string errorMessage) => new()
    {
        Pages = Array.Empty<byte[]>(),
        Success = false,
        ErrorMessage = errorMessage
    };

    private static string CreateBridgeErrorMessage(TypstCompileResult result)
    {
        var diagnostics = result.Diagnostics.Select(diagnostic =>
            $"{diagnostic.Severity}: {diagnostic.Message} ({diagnostic.File ?? "<source>"}:{diagnostic.Line}:{diagnostic.Column})");
        var diagnosticText = string.Join(Environment.NewLine, diagnostics);

        if (string.IsNullOrWhiteSpace(diagnosticText))
        {
            return $"TypstBridge exited with status {result.Status}. {result.Message}".Trim();
        }

        return $"TypstBridge exited with status {result.Status}. {result.Message}{Environment.NewLine}{diagnosticText}".Trim();
    }

    private static string CombineErrors(string? bridgeError, string? fallbackError)
    {
        if (string.IsNullOrWhiteSpace(bridgeError))
        {
            return fallbackError ?? string.Empty;
        }

        if (string.IsNullOrWhiteSpace(fallbackError))
        {
            return bridgeError;
        }

        return $"{fallbackError}{Environment.NewLine}TypstBridge error: {bridgeError}";
    }

    private CompileResult CompileLegacyTypstSharp(string source, CompileOptions options)
    {
        try
        {
            var compilerType = Type.GetType("typstsharp.TypstCompiler, typstsharp")
                ?? throw new InvalidOperationException("typstsharp type not found");

            var fromSource = compilerType.GetMethod("FromSource", new[] { typeof(string) })
                ?? throw new InvalidOperationException("FromSource method not found");

            if (_legacyTypstSharpCompiler is IDisposable previousCompiler)
            {
                previousCompiler.Dispose();
            }

            _legacyTypstSharpCompiler = fromSource.Invoke(null, new object[] { source });

            var compileMethod = compilerType.GetMethod("Compile")
                ?? throw new InvalidOperationException("Compile method not found");

            var result = compileMethod.Invoke(_legacyTypstSharpCompiler, null)!;
            var buffersProperty = result.GetType().GetProperty("Buffers")
                ?? throw new InvalidOperationException("Buffers property not found");

            var buffers = (System.Collections.IList)buffersProperty.GetValue(result)!;
            var pages = new byte[buffers.Count][];
            for (int i = 0; i < buffers.Count; i++)
            {
                pages[i] = (byte[])buffers[i]!;
            }

            return new CompileResult
            {
                Pages = pages,
                Success = true
            };
        }
        catch (Exception)
        {
            // Fall back to the Typst CLI safety net on any legacy typstsharp error.
            return CompileCli(source, options);
        }
    }

    private static CompileResult CompileCli(string source, CompileOptions options)
    {
        var useProvidedDir = !string.IsNullOrEmpty(options.WorkingDirectory) && Directory.Exists(options.WorkingDirectory);
        var compileId = Guid.NewGuid().ToString("N");
        var tempDir = useProvidedDir
            ? options.WorkingDirectory!
            : Path.Combine(Path.GetTempPath(), $"typst_compile_{compileId}");

        if (!useProvidedDir)
        {
            Directory.CreateDirectory(tempDir);
        }

        try
        {
            var typstFile = Path.Combine(tempDir, $"input_{compileId}.typ");
            File.WriteAllText(typstFile, source);

            string outputPattern;
            switch (options.Format)
            {
                case OutputFormat.Png:
                    outputPattern = Path.Combine(tempDir, $"page_{compileId}_{{p}}.png");
                    break;
                case OutputFormat.Svg:
                    outputPattern = Path.Combine(tempDir, $"page_{compileId}_{{p}}.svg");
                    break;
                case OutputFormat.Pdf:
                default:
                    outputPattern = Path.Combine(tempDir, $"output_{compileId}.pdf");
                    break;
            }

            var psi = new ProcessStartInfo("typst")
            {
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                WorkingDirectory = tempDir
            };
            psi.ArgumentList.Add("compile");
            psi.ArgumentList.Add(typstFile);
            psi.ArgumentList.Add(outputPattern);

            if (options.Format == OutputFormat.Png)
            {
                psi.ArgumentList.Add("--ppi");
                psi.ArgumentList.Add(options.Ppi.ToString(CultureInfo.InvariantCulture));
            }

            if (!string.IsNullOrEmpty(options.FontDirectory))
            {
                psi.ArgumentList.Add("--font-path");
                psi.ArgumentList.Add(options.FontDirectory);
            }

            using var process = Process.Start(psi);
            if (process == null)
            {
                return new CompileResult
                {
                    Pages = Array.Empty<byte[]>(),
                    Success = false,
                    ErrorMessage = "Failed to start typst process. Is 'typst' installed and in PATH?"
                };
            }

            var stderrTask = process.StandardError.ReadToEndAsync();
            var stdoutTask = process.StandardOutput.ReadToEndAsync();
            var waitTask = process.WaitForExitAsync();

            if (!waitTask.Wait(options.ProcessTimeout))
            {
                try
                {
                    process.Kill(entireProcessTree: true);
                }
                catch
                {
                    // Best effort termination before reporting timeout.
                }

                return new CompileResult
                {
                    Pages = Array.Empty<byte[]>(),
                    Success = false,
                    ErrorMessage = $"Typst CLI timed out after {options.ProcessTimeout.TotalSeconds.ToString("F0", CultureInfo.InvariantCulture)} seconds."
                };
            }

            var stderr = stderrTask.GetAwaiter().GetResult();
            var stdout = stdoutTask.GetAwaiter().GetResult();

            if (process.ExitCode != 0)
            {
                return new CompileResult
                {
                    Pages = Array.Empty<byte[]>(),
                    Success = false,
                    ErrorMessage = $"Typst CLI exited with code {process.ExitCode}. stderr: {stderr}"
                };
            }

            // Collect output files
            var pages = new List<byte[]>();

            if (options.Format == OutputFormat.Pdf)
            {
                var pdfPath = Path.Combine(tempDir, $"output_{compileId}.pdf");
                if (File.Exists(pdfPath))
                {
                    pages.Add(File.ReadAllBytes(pdfPath));
                }
            }
            else
            {
                // PNG or SVG: collect page files
                var ext = options.Format == OutputFormat.Png ? "png" : "svg";
                var pageFiles = Directory.GetFiles(tempDir, $"page_{compileId}_*.{ext}")
                    .OrderBy(GetTypstPageNumber);

                foreach (var pageFile in pageFiles)
                {
                    pages.Add(File.ReadAllBytes(pageFile));
                }
            }

            if (pages.Count == 0)
            {
                return new CompileResult
                {
                    Pages = Array.Empty<byte[]>(),
                    Success = false,
                    ErrorMessage = $"Typst CLI produced no output. stdout: {stdout} stderr: {stderr}"
                };
            }

            return new CompileResult
            {
                Pages = pages.ToArray(),
                Success = true
            };
        }
        catch (Exception ex)
        {
            return new CompileResult
            {
                Pages = Array.Empty<byte[]>(),
                Success = false,
                ErrorMessage = $"CLI compilation failed: {ex.Message}"
            };
        }
        finally
        {
            try
            {
                if (useProvidedDir)
                {
                    foreach (var file in Directory.GetFiles(tempDir, $"*_{compileId}*"))
                    {
                        File.Delete(file);
                    }
                }
                else
                {
                    Directory.Delete(tempDir, true);
                }
            }
            catch
            {
                // Best effort cleanup
            }
        }
    }

    private static int GetTypstPageNumber(string path)
    {
        var fileName = Path.GetFileNameWithoutExtension(path);
        var markerIndex = fileName.LastIndexOf('_');
        if (markerIndex >= 0 && int.TryParse(fileName[(markerIndex + 1)..], NumberStyles.Integer, CultureInfo.InvariantCulture, out var pageNumber))
        {
            return pageNumber;
        }

        return int.MaxValue;
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            if (_legacyTypstSharpCompiler is IDisposable d)
            {
                d.Dispose();
            }
            _disposed = true;
        }
    }
}
