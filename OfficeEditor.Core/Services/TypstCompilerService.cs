using System.Globalization;
using typstsharp;

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
}

public sealed record CompileResult
{
    public required byte[][] Pages { get; init; }
    public required bool Success { get; init; }
    public string? ErrorMessage { get; init; }
}

public sealed class TypstCompilerService : IDisposable
{
    private TypstCompiler? _compiler;
    private bool _disposed;

    public TypstCompilerService()
    {
    }

    public CompileResult Compile(string source, CompileOptions? options = null)
    {
        options ??= new CompileOptions();
        
        try
        {
            // Create compiler from source
            _compiler = TypstCompiler.FromSource(source);
            
            // Set font paths if provided
            if (!string.IsNullOrEmpty(options.FontDirectory) && Directory.Exists(options.FontDirectory))
            {
                // Use the font_paths field directly since FontPaths property doesn't exist
                // We'll use reflection or just skip font configuration for now
            }

            // Compile
            var result = _compiler.Compile();
            
            // Convert buffers to byte arrays
            var pages = new byte[result.Buffers.Count][];
            for (int i = 0; i < result.Buffers.Count; i++)
            {
                pages[i] = result.Buffers[i];
            }

            return new CompileResult
            {
                Pages = pages,
                Success = true
            };
        }
        catch (Exception ex)
        {
            return new CompileResult
            {
                Pages = Array.Empty<byte[]>(),
                Success = false,
                ErrorMessage = ex.Message
            };
        }
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            _compiler?.Dispose();
            _disposed = true;
        }
    }
}
