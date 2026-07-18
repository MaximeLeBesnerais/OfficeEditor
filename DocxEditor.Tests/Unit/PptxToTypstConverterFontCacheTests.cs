using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using PptxEditor.Core.Builders;
using PptxEditor.Core.Converters;
using Xunit;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Tests for the process-wide system-font cache in <see cref="PptxToTypstConverter"/>.
/// The expensive system-font scan (directory walk + fc-list) must run at most once per
/// process, be shared across converter instances, and be explicitly resettable via
/// <see cref="PptxToTypstConverter.InvalidateSystemFontCache"/>.
/// </summary>
public sealed class PptxToTypstConverterFontCacheTests : IDisposable
{
    private static readonly FieldInfo s_cacheField = typeof(PptxToTypstConverter).GetField(
        "s_systemFontCache", BindingFlags.Static | BindingFlags.NonPublic)
        ?? throw new InvalidOperationException("Expected private static field 's_systemFontCache' on PptxToTypstConverter.");

    private readonly string _tempDir;

    public PptxToTypstConverterFontCacheTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), nameof(PptxToTypstConverterFontCacheTests), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        SlideOpsTestHelpers.BestEffortDelete(_tempDir);
    }

    [Fact]
    public void SystemFontScan_IsSharedAcrossConverterInstances()
    {
        PptxToTypstConverter.InvalidateSystemFontCache();
        Assert.Null(s_cacheField.GetValue(null));

        var firstSource = ConvertOnce(out var cacheAfterFirst);
        Assert.NotNull(cacheAfterFirst);

        var secondSource = ConvertOnce(out var cacheAfterSecond);

        // Same cache instance reused: the scan did not run again for the second converter.
        Assert.Same(cacheAfterFirst, cacheAfterSecond);
        // Cache sharing must not change conversion output.
        Assert.Equal(firstSource, secondSource);
    }

    [Fact]
    public void InvalidateSystemFontCache_ClearsCacheAndNextConversionRescans()
    {
        ConvertOnce(out var populated);
        Assert.NotNull(populated);

        PptxToTypstConverter.InvalidateSystemFontCache();
        Assert.Null(s_cacheField.GetValue(null));

        ConvertOnce(out var repopulated);
        Assert.NotNull(repopulated);
    }

    /// <summary>Runs a full Convert + source emission on a fresh document/converter pair.</summary>
    private string ConvertOnce(out object? cacheSnapshot)
    {
        var path = CreateDeck();
        using var document = PresentationDocument.Open(path, false);
        using var converter = new PptxToTypstConverter(document);

        var source = converter.GenerateTypstSource(converter.Convert());
        cacheSnapshot = s_cacheField.GetValue(null);
        return source;
    }

    private string CreateDeck()
    {
        var path = Path.Combine(_tempDir, $"{Guid.NewGuid():N}.pptx");
        using (var builder = PresentationBuilder.Create(path))
        {
            builder.AddSlide();
            builder.CurrentSlide
                .AddTitle("Font cache test")
                .AddText("Some body text");
            builder.Save();
        }

        return path;
    }
}
