using System.Diagnostics;
using OfficeEditor.Core.Services;

namespace OfficeEditor.Api.Services;

/// <summary>
/// Fire-and-forget startup warmup for the Typst render paths. Compiles a trivial
/// one-slide document once without a font directory (the REF-render path) and once
/// with the configured Demo:FontDirectory (the generation-preview path) — the native
/// font cache is keyed by the font-path list, so each path needs its own warm compile.
/// Never blocks startup and never crashes the host: per-step failures are logged as
/// warnings (the one place where swallowing is acceptable).
/// </summary>
public sealed class RenderWarmupService : IHostedService
{
    private const string WarmupSource =
        "#set page(width: 960pt, height: 540pt)\n#place(top + left, dx: 40pt, dy: 40pt)[warm]";

    private readonly TypstCompilerService _compiler;
    private readonly string? _fontDirectory;
    private readonly ILogger<RenderWarmupService> _logger;

    public RenderWarmupService(
        TypstCompilerService compiler,
        string? fontDirectory,
        ILogger<RenderWarmupService> logger)
    {
        _compiler = compiler;
        _fontDirectory = string.IsNullOrWhiteSpace(fontDirectory) ? null : fontDirectory;
        _logger = logger;
    }

    public Task StartAsync(CancellationToken cancellationToken)
    {
        _ = Task.Run(Warmup, CancellationToken.None);
        return Task.CompletedTask;
    }

    public Task StopAsync(CancellationToken cancellationToken) => Task.CompletedTask;

    private void Warmup()
    {
        CompileStep("default-fonts", fontDirectory: null);
        if (_fontDirectory is not null)
        {
            CompileStep("demo-fonts", _fontDirectory);
        }
    }

    private void CompileStep(string label, string? fontDirectory)
    {
        try
        {
            _logger.LogInformation("Render warmup ({Label}) starting.", label);
            var timer = Stopwatch.StartNew();
            _compiler.Compile(WarmupSource, new CompileOptions
            {
                Format = OutputFormat.Png,
                FontDirectory = fontDirectory
            });
            timer.Stop();
            _logger.LogInformation(
                "Render warmup ({Label}) completed in {ElapsedMs:F0} ms.",
                label,
                timer.Elapsed.TotalMilliseconds);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Render warmup ({Label}) failed; continuing without a warm cache.", label);
        }
    }
}
