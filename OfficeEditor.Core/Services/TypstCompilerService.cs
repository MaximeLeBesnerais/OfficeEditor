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

        if (_nativeAvailable)
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
        catch (Exception ex)
        {
            // Fall back to CLI on any native error
            return CompileCli(source, options);
        }
    }

    private static CompileResult CompileCli(string source, CompileOptions options)
    {
        var useProvidedDir = !string.IsNullOrEmpty(options.WorkingDirectory) && Directory.Exists(options.WorkingDirectory);
        var tempDir = useProvidedDir
            ? options.WorkingDirectory!
            : Path.Combine(Path.GetTempPath(), $"typst_compile_{Guid.NewGuid():N}");
        
        if (!useProvidedDir)
        {
            Directory.CreateDirectory(tempDir);
        }

        try
        {
            var typstFile = Path.Combine(tempDir, "input.typ");
            File.WriteAllText(typstFile, source);

            string outputPattern;
            switch (options.Format)
            {
                case OutputFormat.Png:
                    outputPattern = Path.Combine(tempDir, "page_{p}.png");
                    break;
                case OutputFormat.Svg:
                    outputPattern = Path.Combine(tempDir, "page_{p}.svg");
                    break;
                case OutputFormat.Pdf:
                default:
                    outputPattern = Path.Combine(tempDir, "output.pdf");
                    break;
            }

            var args = new List<string>
            {
                "compile",
                $"\"{typstFile}\"",
                $"\"{outputPattern}\""
            };

            if (options.Format == OutputFormat.Png)
            {
                args.Add($"--ppi {options.Ppi.ToString(CultureInfo.InvariantCulture)}");
            }

            if (!string.IsNullOrEmpty(options.FontDirectory) && Directory.Exists(options.FontDirectory))
            {
                args.Add($"--font-path \"{options.FontDirectory}\"");
            }

            var psi = new ProcessStartInfo("typst", string.Join(" ", args))
            {
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                WorkingDirectory = tempDir
            };

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

            process.WaitForExit();
            var stderr = process.StandardError.ReadToEnd();
            var stdout = process.StandardOutput.ReadToEnd();

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
                var pdfPath = Path.Combine(tempDir, "output.pdf");
                if (File.Exists(pdfPath))
                {
                    pages.Add(File.ReadAllBytes(pdfPath));
                }
            }
            else
            {
                // PNG or SVG: collect page files
                var ext = options.Format == OutputFormat.Png ? "png" : "svg";
                var pageFiles = Directory.GetFiles(tempDir, $"page_*.{ext}")
                    .OrderBy(f => f, StringComparer.Ordinal);

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
            if (!useProvidedDir)
            {
                try
                {
                    Directory.Delete(tempDir, true);
                }
                catch
                {
                    // Best effort cleanup
                }
            }
        }
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
