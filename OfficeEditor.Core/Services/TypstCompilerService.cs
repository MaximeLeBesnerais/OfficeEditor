using System.Diagnostics;
using System.Globalization;

namespace OfficeEditor.Core.Services;

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
    private static readonly bool _nativeAvailable;
    private dynamic? _compiler; // typstsharp compiler when native is available

    static TypstCompilerService()
    {
        // Probe whether the native typstsharp library can be loaded
        try
        {
            var compilerType = Type.GetType("typstsharp.TypstCompiler, typstsharp");
            if (compilerType != null)
            {
                var method = compilerType.GetMethod("FromSource", new[] { typeof(string) });
                if (method != null)
                {
                    using var probe = (IDisposable?)method.Invoke(null, new object[] { "#text[probe]" });
                    _nativeAvailable = probe != null;
                }
            }
        }
        catch
        {
            _nativeAvailable = false;
        }
    }

    public CompileResult Compile(string source, CompileOptions? options = null)
    {
        options ??= new CompileOptions();

        if (_nativeAvailable && options.Format == OutputFormat.Pdf && string.IsNullOrEmpty(options.FontDirectory))
        {
            return CompileNative(source, options);
        }

        return CompileCli(source, options);
    }

    private CompileResult CompileNative(string source, CompileOptions options)
    {
        try
        {
            var compilerType = Type.GetType("typstsharp.TypstCompiler, typstsharp")
                ?? throw new InvalidOperationException("typstsharp type not found");

            var fromSource = compilerType.GetMethod("FromSource", new[] { typeof(string) })
                ?? throw new InvalidOperationException("FromSource method not found");

            if (_compiler is IDisposable previousCompiler)
            {
                previousCompiler.Dispose();
            }

            _compiler = fromSource.Invoke(null, new object[] { source });

            var compileMethod = compilerType.GetMethod("Compile")
                ?? throw new InvalidOperationException("Compile method not found");

            var result = compileMethod.Invoke(_compiler, null)!;
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
            // Fall back to CLI on any native error
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
            if (_compiler is IDisposable d)
            {
                d.Dispose();
            }
            _disposed = true;
        }
    }
}
