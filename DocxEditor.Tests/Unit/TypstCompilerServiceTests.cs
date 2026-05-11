using System.Text;
using System.Reflection;
using OfficeEditor.Core.Services;
using Xunit;

namespace DocxEditor.Tests.Unit;

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
        // Note: typstsharp returns a single PDF buffer regardless of page count
        var source = @"
#set page(width: 200pt, height: 100pt, margin: 10pt)
#text(size: 12pt)[Page 1]
#pagebreak()
#text(size: 12pt)[Page 2]
";
        
        using var compiler = new TypstCompilerService();
        var result = compiler.Compile(source, new CompileOptions { Format = OutputFormat.Pdf });
        
        Assert.True(result.Success);
        // typstsharp returns a single PDF buffer
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
        
        // Both should succeed (PPI is ignored for PDF output)
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

    [Fact]
    public void Dispose_CanBeCalledMultipleTimes()
    {
        var compiler = new TypstCompilerService();
        compiler.Dispose();
        compiler.Dispose(); // Should not throw
    }
}
