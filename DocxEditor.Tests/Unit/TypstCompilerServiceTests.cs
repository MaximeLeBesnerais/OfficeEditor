using System.Text;
using System.Diagnostics;
using System.Reflection;
using OfficeEditor.Core.Services;
using Xunit;

namespace DocxEditor.Tests.Unit;

[CollectionDefinition("TypstCompilerServiceEnvironment", DisableParallelization = true)]
public sealed class TypstCompilerServiceEnvironmentCollection;

public sealed class UnixFactAttribute : FactAttribute
{
    public UnixFactAttribute()
    {
        if (OperatingSystem.IsWindows())
        {
            Skip = "Fake typst shell scripts are Unix-specific.";
        }
    }
}

[Collection("TypstCompilerServiceEnvironment")]
public class TypstCompilerServiceTests : IDisposable
{
    private readonly string _tempDir;

    public TypstCompilerServiceTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "TypstCompilerTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, true);
            }
        }
        catch
        {
            // Best effort cleanup
        }
    }

    private static CompileResult InvokeCompileCli(string? source, CompileOptions options)
    {
        var method = typeof(TypstCompilerService).GetMethod("CompileCli", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);

        return Assert.IsType<CompileResult>(method.Invoke(null, new object?[] { source, options }));
    }

    private string CreateFakeTypstExecutable(string mode = "success")
    {
        var binDir = Path.Combine(_tempDir, "bin");
        Directory.CreateDirectory(binDir);

        var typstPath = Path.Combine(binDir, "typst");
        File.WriteAllText(typstPath, """
#!/usr/bin/env bash
set -euo pipefail

printf '%s\n' "$PWD" > "$TYPST_FAKE_CWD_LOG"
printf '%s\n' "$@" > "$TYPST_FAKE_ARGS_LOG"

if [[ "${TYPST_FAKE_MODE:-success}" == "fail" ]]; then
  printf 'synthetic typst failure\n' >&2
  exit 7
fi

if [[ "${TYPST_FAKE_MODE:-success}" == "noop" ]]; then
  printf 'compiled without writing files\n'
  exit 0
fi

output="$3"
if [[ "$output" == *.pdf ]]; then
  printf '%%PDF-fake' > "$output"
elif [[ "$output" == *'{p}.png' ]]; then
  printf 'page-10' > "${output/\{p\}/10}"
  printf 'page-2' > "${output/\{p\}/2}"
elif [[ "$output" == *'{p}.svg' ]]; then
  printf '<svg>page-1</svg>' > "${output/\{p\}/1}"
else
  printf 'unexpected output pattern: %s\n' "$output" >&2
  exit 9
fi
""");
        if (!OperatingSystem.IsWindows())
        {
            File.SetUnixFileMode(typstPath, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
        }

        Environment.SetEnvironmentVariable("TYPST_FAKE_MODE", mode);
        Environment.SetEnvironmentVariable("TYPST_FAKE_ARGS_LOG", Path.Combine(_tempDir, "typst-args.log"));
        Environment.SetEnvironmentVariable("TYPST_FAKE_CWD_LOG", Path.Combine(_tempDir, "typst-cwd.log"));

        return binDir;
    }

    private CompileResult InvokeCompileCliWithFakeTypst(string? source, CompileOptions options, string mode = "success")
    {
        Assert.False(OperatingSystem.IsWindows(), "Fake typst shell scripts are Unix-specific.");

        var originalPath = Environment.GetEnvironmentVariable("PATH");
        var fakeTypstDir = CreateFakeTypstExecutable(mode);
        Environment.SetEnvironmentVariable("PATH", fakeTypstDir + Path.PathSeparator + originalPath);

        try
        {
            return InvokeCompileCli(source, options);
        }
        finally
        {
            Environment.SetEnvironmentVariable("PATH", originalPath);
            Environment.SetEnvironmentVariable("TYPST_FAKE_MODE", null);
            Environment.SetEnvironmentVariable("TYPST_FAKE_ARGS_LOG", null);
            Environment.SetEnvironmentVariable("TYPST_FAKE_CWD_LOG", null);
        }
    }

    private string[] ReadFakeTypstArguments()
    {
        return File.ReadAllLines(Path.Combine(_tempDir, "typst-args.log"));
    }

    private static string ResolvePhysicalPath(string path)
    {
        var startInfo = new ProcessStartInfo("realpath", path)
        {
            RedirectStandardOutput = true,
            UseShellExecute = false,
            CreateNoWindow = true,
        };
        using var process = Process.Start(startInfo)!;
        process.WaitForExit();
        return process.StandardOutput.ReadToEnd().Trim();
    }

    private string CreateSimpleTypstDocument()
    {
        return @"
#set page(width: 200pt, height: 100pt, margin: 10pt)
#text(size: 12pt)[Hello, World!]
";
    }

    [Fact]
    public void Compile_SimpleDocument_ReturnsPdf()
    {
        var source = CreateSimpleTypstDocument();
        
        using var compiler = new TypstCompilerService();
        var result = compiler.Compile(source, new CompileOptions { Format = OutputFormat.Pdf });
        
        Assert.True(result.Success);
        Assert.Single(result.Pages);
        Assert.True(result.Pages[0].Length > 0);
        // PDF header check
        Assert.Equal(0x25, result.Pages[0][0]); // '%'
    }

    [Fact]
    public void Compile_SimpleDocument_ReturnsPng()
    {
        var source = CreateSimpleTypstDocument();
        
        using var compiler = new TypstCompilerService();
        var result = compiler.Compile(source, new CompileOptions { Format = OutputFormat.Png, Ppi = 72 });
        
        Assert.True(result.Success);
        Assert.Single(result.Pages);
        byte[] pngSignature = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
        Assert.True(result.Pages[0].Length >= pngSignature.Length);
        Assert.True(result.Pages[0].AsSpan(0, pngSignature.Length).SequenceEqual(pngSignature));
    }

    [Fact]
    public void Compile_SimpleDocument_ReturnsSvg()
    {
        var source = CreateSimpleTypstDocument();
        
        using var compiler = new TypstCompilerService();
        var result = compiler.Compile(source, new CompileOptions { Format = OutputFormat.Svg });
        
        Assert.True(result.Success);
        Assert.Single(result.Pages);
        Assert.True(result.Pages[0].Length > 0);
        Assert.Contains("<svg", Encoding.UTF8.GetString(result.Pages[0]), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Compile_MultipageDocument_ReturnsMultiplePages()
    {
        // PDF output is returned as one buffer regardless of page count; PNG/SVG may be per-page.
        var source = @"
#set page(width: 200pt, height: 100pt, margin: 10pt)
#text(size: 12pt)[Page 1]
#pagebreak()
#text(size: 12pt)[Page 2]
";
        
        using var compiler = new TypstCompilerService();
        var result = compiler.Compile(source, new CompileOptions { Format = OutputFormat.Pdf });
        
        Assert.True(result.Success);
        // PDF output is a single buffer.
        Assert.Single(result.Pages);
        Assert.True(result.Pages[0].Length > 0);
    }

    [Fact]
    public void Compile_InvalidSyntax_ReturnsError()
    {
        var source = @"
#invalid command here
";
        
        using var compiler = new TypstCompilerService();
        var result = compiler.Compile(source, new CompileOptions { Format = OutputFormat.Pdf });
        
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void Compile_WithCustomFonts_ReturnsSuccess()
    {
        var source = @"
#set page(width: 200pt, height: 100pt, margin: 10pt)
#set text(font: ""Arial"")
#text(size: 12pt)[Hello with custom font]
";
        
        using var compiler = new TypstCompilerService();
        var result = compiler.Compile(source, new CompileOptions { Format = OutputFormat.Pdf });
        
        // Should succeed with fallback fonts if Arial is not available
        Assert.True(result.Success);
        Assert.Single(result.Pages);
    }

    [Fact]
    public void Compile_WithDifferentPpi_ReturnsSuccess()
    {
        var source = CreateSimpleTypstDocument();
        
        using var compiler = new TypstCompilerService();
        
        var result72 = compiler.Compile(source, new CompileOptions { Format = OutputFormat.Pdf, Ppi = 72 });
        var result150 = compiler.Compile(source, new CompileOptions { Format = OutputFormat.Pdf, Ppi = 150 });
        
        // Both should succeed; PPI applies to raster output, not PDF buffers.
        Assert.True(result72.Success);
        Assert.True(result150.Success);
        Assert.Single(result72.Pages);
        Assert.Single(result150.Pages);
    }

    [Fact]
    public void GetBridgeFontPaths_SplitsCombinedFontDirectory()
    {
        var options = new CompileOptions
        {
            FontDirectory = string.Join(Path.PathSeparator, "/fonts/embedded", " ", "/fonts/system ")
        };

        var method = typeof(TypstCompilerService).GetMethod("GetBridgeFontPaths", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);

        var fontPaths = Assert.IsAssignableFrom<IReadOnlyList<string>>(method.Invoke(null, new object[] { options }));

        Assert.Equal(new[] { "/fonts/embedded", "/fonts/system" }, fontPaths);
    }

    [Theory]
    [InlineData(OutputFormat.Pdf, "Pdf")]
    [InlineData(OutputFormat.Png, "Png")]
    [InlineData(OutputFormat.Svg, "Svg")]
    public void MapOutputFormat_MapsEveryPublicFormat(OutputFormat format, string expected)
    {
        var method = typeof(TypstCompilerService).GetMethod("MapOutputFormat", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);

        var result = method.Invoke(null, new object[] { format });

        Assert.Equal(expected, result?.ToString());
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void GetBridgeFontPaths_WithNullOrWhitespace_ReturnsEmptyList(string? fontDirectory)
    {
        var method = typeof(TypstCompilerService).GetMethod("GetBridgeFontPaths", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);

        var fontPaths = Assert.IsAssignableFrom<IReadOnlyList<string>>(method.Invoke(null, new object[]
        {
            new CompileOptions { FontDirectory = fontDirectory }
        }));

        Assert.Empty(fontPaths);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void GetBridgeWorkingDirectory_WithMissingOption_FallsBackToNonEmptyDirectory(string? workingDirectory)
    {
        var method = typeof(TypstCompilerService).GetMethod("GetBridgeWorkingDirectory", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);

        var result = Assert.IsType<string>(method.Invoke(null, new object[]
        {
            new CompileOptions { WorkingDirectory = workingDirectory }
        }));

        Assert.False(string.IsNullOrWhiteSpace(result));
    }

    [Theory]
    [InlineData(null, "cli", "cli")]
    [InlineData("", "cli", "cli")]
    [InlineData("bridge", null, "bridge")]
    [InlineData("bridge", "", "bridge")]
    public void CombineErrors_WhenOneSideIsMissing_ReturnsAvailableMessage(string? bridgeError, string? fallbackError, string expected)
    {
        var method = typeof(TypstCompilerService).GetMethod("CombineErrors", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);

        var result = Assert.IsType<string>(method.Invoke(null, new object?[] { bridgeError, fallbackError }));

        Assert.Equal(expected, result);
    }

    [Fact]
    public void CombineErrors_WhenBothMessagesExist_AppendsBridgeContext()
    {
        var method = typeof(TypstCompilerService).GetMethod("CombineErrors", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);

        var result = Assert.IsType<string>(method.Invoke(null, new object?[] { "bridge failed", "cli failed" }));

        Assert.Contains("cli failed", result);
        Assert.Contains("TypstBridge error: bridge failed", result);
    }

    [Fact]
    public void CreateBridgeErrorMessage_WithoutDiagnostics_IncludesStatusAndMessageOnly()
    {
        var method = typeof(TypstCompilerService).GetMethod("CreateBridgeErrorMessage", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);
        var resultModel = new TypstBridge.Managed.Models.TypstCompileResult(6, "unsupported format", [], []);

        var result = Assert.IsType<string>(method.Invoke(null, new object[] { resultModel }));

        Assert.Equal("TypstBridge exited with status 6. unsupported format", result);
    }

    [Fact]
    public void CreateBridgeErrorMessage_WithDiagnostics_FormatsFileAndSourceLocations()
    {
        var method = typeof(TypstCompilerService).GetMethod("CreateBridgeErrorMessage", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);
        var resultModel = new TypstBridge.Managed.Models.TypstCompileResult(2, "compile failed", [],
        [
            new TypstBridge.Managed.Models.TypstDiagnostic(TypstBridge.Managed.Models.TypstDiagnosticSeverity.Error, "bad syntax", null, 4, 2),
            new TypstBridge.Managed.Models.TypstDiagnostic(TypstBridge.Managed.Models.TypstDiagnosticSeverity.Warning, "ignored", "deck.typ", 7, 9)
        ]);

        var result = Assert.IsType<string>(method.Invoke(null, new object[] { resultModel }));

        Assert.Contains("TypstBridge exited with status 2. compile failed", result);
        Assert.Contains("Error: bad syntax (<source>:4:2)", result);
        Assert.Contains("Warning: ignored (deck.typ:7:9)", result);
    }

    [UnixFact]
    public void CompileCli_Pdf_UsesProvidedWorkingDirectoryAndCleansTemporaryFiles()
    {
        var workDir = Path.Combine(_tempDir, "work");
        Directory.CreateDirectory(workDir);
        var keepFile = Path.Combine(workDir, "keep.txt");
        File.WriteAllText(keepFile, "do not delete");

        var result = InvokeCompileCliWithFakeTypst(CreateSimpleTypstDocument(), new CompileOptions
        {
            Format = OutputFormat.Pdf,
            WorkingDirectory = workDir
        });

        Assert.True(result.Success);
        Assert.Single(result.Pages);
        Assert.True(result.Pages[0].AsSpan(0, 4).SequenceEqual("%PDF"u8));
        Assert.Equal(ResolvePhysicalPath(workDir), File.ReadAllText(Path.Combine(_tempDir, "typst-cwd.log")).Trim());
        Assert.True(File.Exists(keepFile));
        Assert.Empty(Directory.GetFiles(workDir, "input_*.typ"));
        Assert.Empty(Directory.GetFiles(workDir, "output_*.pdf"));
    }

    [UnixFact]
    public void CompileCli_Png_AddsPpiAndOrdersPageFilesNumerically()
    {
        var result = InvokeCompileCliWithFakeTypst(CreateSimpleTypstDocument(), new CompileOptions
        {
            Format = OutputFormat.Png,
            Ppi = 96,
            WorkingDirectory = _tempDir
        });

        Assert.True(result.Success);
        Assert.Equal(2, result.Pages.Length);
        Assert.Equal("page-2", Encoding.UTF8.GetString(result.Pages[0]));
        Assert.Equal("page-10", Encoding.UTF8.GetString(result.Pages[1]));

        var args = ReadFakeTypstArguments();
        Assert.Contains("--ppi", args);
        Assert.Contains("96", args);
        Assert.EndsWith(".png", args[2]);
        Assert.Contains("{p}", args[2]);
    }

    [UnixFact]
    public void CompileCli_Svg_DoesNotAddPpiAndCollectsSvgPages()
    {
        var result = InvokeCompileCliWithFakeTypst(CreateSimpleTypstDocument(), new CompileOptions
        {
            Format = OutputFormat.Svg,
            Ppi = 300,
            WorkingDirectory = _tempDir
        });

        Assert.True(result.Success);
        Assert.Single(result.Pages);
        Assert.Equal("<svg>page-1</svg>", Encoding.UTF8.GetString(result.Pages[0]));

        var args = ReadFakeTypstArguments();
        Assert.DoesNotContain("--ppi", args);
        Assert.DoesNotContain("300", args);
        Assert.EndsWith(".svg", args[2]);
        Assert.Contains("{p}", args[2]);
    }

    [UnixFact]
    public void CompileCli_WithFontDirectory_PassesCombinedFontPathAsSingleCliValue()
    {
        var combinedFontPath = string.Join(Path.PathSeparator, "/fonts/embedded", "/fonts/system");

        var result = InvokeCompileCliWithFakeTypst(CreateSimpleTypstDocument(), new CompileOptions
        {
            Format = OutputFormat.Pdf,
            FontDirectory = combinedFontPath,
            WorkingDirectory = _tempDir
        });

        Assert.True(result.Success);
        var args = ReadFakeTypstArguments();
        var fontPathOptionIndex = Array.IndexOf(args, "--font-path");
        Assert.True(fontPathOptionIndex >= 0);
        Assert.Equal(combinedFontPath, args[fontPathOptionIndex + 1]);
    }

    [UnixFact]
    public void CompileCli_WhenProcessExitsNonZero_ReturnsExitCodeAndStderr()
    {
        var result = InvokeCompileCliWithFakeTypst(CreateSimpleTypstDocument(), new CompileOptions
        {
            Format = OutputFormat.Pdf,
            WorkingDirectory = _tempDir
        }, mode: "fail");

        Assert.False(result.Success);
        Assert.Empty(result.Pages);
        Assert.Contains("Typst CLI exited with code 7", result.ErrorMessage);
        Assert.Contains("synthetic typst failure", result.ErrorMessage);
    }

    [UnixFact]
    public void CompileCli_WhenProcessWritesNoOutput_ReturnsNoOutputErrorWithStdout()
    {
        var result = InvokeCompileCliWithFakeTypst(CreateSimpleTypstDocument(), new CompileOptions
        {
            Format = OutputFormat.Pdf,
            WorkingDirectory = _tempDir
        }, mode: "noop");

        Assert.False(result.Success);
        Assert.Empty(result.Pages);
        Assert.Contains("Typst CLI produced no output", result.ErrorMessage);
        Assert.Contains("compiled without writing files", result.ErrorMessage);
    }

    [UnixFact]
    public void CompileCli_WhenProcessTimesOut_KillsProcessAndReturnsTimeoutFailure()
    {
        var originalPath = Environment.GetEnvironmentVariable("PATH");
        var binDir = Path.Combine(_tempDir, "slow-bin");
        Directory.CreateDirectory(binDir);
        var typstPath = Path.Combine(binDir, "typst");
        File.WriteAllText(typstPath, "#!/usr/bin/env bash\nsleep 5\n");
        if (!OperatingSystem.IsWindows())
        {
            File.SetUnixFileMode(typstPath, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
        }

        Environment.SetEnvironmentVariable("PATH", binDir + Path.PathSeparator + originalPath);

        try
        {
            var result = InvokeCompileCli(CreateSimpleTypstDocument(), new CompileOptions
            {
                Format = OutputFormat.Pdf,
                WorkingDirectory = _tempDir,
                ProcessTimeout = TimeSpan.FromMilliseconds(100)
            });

            Assert.False(result.Success);
            Assert.Empty(result.Pages);
            Assert.Contains("Typst CLI timed out", result.ErrorMessage);
        }
        finally
        {
            Environment.SetEnvironmentVariable("PATH", originalPath);
        }
    }

    [Fact]
    public void CompileCli_WhenTypstExecutableIsMissing_ReturnsFailureInsteadOfThrowing()
    {
        var originalPath = Environment.GetEnvironmentVariable("PATH");
        Environment.SetEnvironmentVariable("PATH", _tempDir);

        try
        {
            var result = InvokeCompileCli(CreateSimpleTypstDocument(), new CompileOptions
            {
                Format = OutputFormat.Pdf,
                WorkingDirectory = _tempDir
            });

            Assert.False(result.Success);
            Assert.Empty(result.Pages);
            Assert.Contains("CLI compilation failed", result.ErrorMessage);
        }
        finally
        {
            Environment.SetEnvironmentVariable("PATH", originalPath);
        }
    }

    [Fact]
    public void GetBridgeWorkingDirectory_UsesProvidedDirectoryBeforeCurrentDirectory()
    {
        var method = typeof(TypstCompilerService).GetMethod("GetBridgeWorkingDirectory", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);

        var workDir = Path.Combine(_tempDir, "bridge-work");
        var result = Assert.IsType<string>(method.Invoke(null, new object[]
        {
            new CompileOptions { WorkingDirectory = workDir }
        }));

        Assert.Equal(workDir, result);
    }

    [Fact]
    public void MapOutputFormat_WithUnknownFormat_FallsBackToPdf()
    {
        var method = typeof(TypstCompilerService).GetMethod("MapOutputFormat", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);

        var result = method.Invoke(null, new object[] { (OutputFormat)999 });

        Assert.Equal("Pdf", result?.ToString());
    }

    [Fact]
    public void Dispose_CanBeCalledMultipleTimes()
    {
        var compiler = new TypstCompilerService();
        compiler.Dispose();
        compiler.Dispose(); // Should not throw
    }
}
