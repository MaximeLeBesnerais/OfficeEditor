using System.Text;
using PptxEditor.Core.Builders;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Tests for <see cref="PresentationBuilder.ExportThumbnail"/> (per-slide render over the
/// whole-deck fallback) and for <see cref="ThumbnailOptions.Format"/> flowing through to
/// the compiler (png default, svg honored). Follows the compile-dependent smoke-test
/// pattern: real Typst compilation on a minimal in-memory deck.
/// </summary>
public sealed class ExportThumbnailTests
{
    private static IPresentationBuilder CreateDeckWithTitles(params string[] titles)
    {
        var builder = PresentationBuilder.Create();
        foreach (var title in titles)
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTitle(title);
        }

        return builder;
    }

    [Fact]
    public void ExportThumbnails_UnsupportedFormat_ThrowsArgumentException()
    {
        using var builder = CreateDeckWithTitles("Only slide");

        // Validation happens before any compilation.
        Assert.Throws<ArgumentException>(() =>
            builder.ExportThumbnails(new ThumbnailOptions { Format = "bmp" }));
    }

    [Fact]
    public void ExportThumbnail_InvalidSlideIndex_Throws()
    {
        using var builder = CreateDeckWithTitles("Only slide");

        Assert.Throws<ArgumentOutOfRangeException>(() => builder.ExportThumbnail(-1));
        Assert.Throws<ArgumentOutOfRangeException>(() => builder.ExportThumbnail(1));
    }

    [Fact]
    public void ExportThumbnails_SvgFormat_ReturnsSvgPages()
    {
        using var builder = CreateDeckWithTitles("Slide one", "Slide two");

        var pages = builder.ExportThumbnails(new ThumbnailOptions { Format = "svg" });

        Assert.Equal(2, pages.Length);
        foreach (var page in pages)
        {
            // SVG signature: the document text starts with (or quickly contains) "<svg".
            var head = Encoding.ASCII.GetString(page, 0, Math.Min(page.Length, 512));
            Assert.Contains("<svg", head);
        }
    }

    [Fact]
    public void ExportThumbnails_DefaultFormat_ReturnsPngPages()
    {
        using var builder = CreateDeckWithTitles("Slide one", "Slide two");

        var pages = builder.ExportThumbnails();

        Assert.Equal(2, pages.Length);
        foreach (var page in pages)
        {
            AssertPngSignature(page);
        }
    }

    [Fact]
    public void ExportThumbnail_ReturnsPageMatchingWholeDeckRender()
    {
        byte[] thumbnail;
        byte[][] allPages;
        using (var builder = CreateDeckWithTitles("Slide one", "Slide two"))
        {
            thumbnail = builder.ExportThumbnail(1);
        }

        using (var builder = CreateDeckWithTitles("Slide one", "Slide two"))
        {
            allPages = builder.ExportThumbnails();
        }

        Assert.Equal(2, allPages.Length);
        AssertPngSignature(thumbnail);

        // Same slide, same defaults → identical bytes as the corresponding page of the
        // whole-deck render (deterministic backend output).
        Assert.Equal(allPages[1], thumbnail);
    }

    [Fact]
    public void ExportThumbnails_CompileFailure_ThrowsInvalidOperationWithCompilerError()
    {
        // Regression: a failed Typst compile used to come back as an EMPTY page array,
        // which the demo upload render serialized as {"success": true, "previews": []}
        // — the UI showed nothing with no error. Compile failures must throw.
        using var builder = CreateDeckWithUndecodableImage();

        var ex = Assert.Throws<InvalidOperationException>(() => builder.ExportThumbnails());

        Assert.Contains("Typst", ex.Message);
    }

    /// <summary>
    /// A deck whose only content is an image Typst cannot decode (garbage bytes saved
    /// as .jpg), so the whole-deck compile fails deterministically.
    /// </summary>
    private static IPresentationBuilder CreateDeckWithUndecodableImage()
    {
        var imagePath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.jpg");
        File.WriteAllBytes(imagePath, [0xDE, 0xAD, 0xBE, 0xEF, 0x00, 0x01, 0x02, 0x03]);
        try
        {
            var builder = PresentationBuilder.Create();
            builder.AddSlide();
            builder.CurrentSlide.AddImage(imagePath);
            return builder;
        }
        finally
        {
            if (File.Exists(imagePath))
            {
                File.Delete(imagePath);
            }
        }
    }

    private static void AssertPngSignature(byte[] bytes)
    {
        Assert.True(bytes.Length > 8, "PNG page is suspiciously small.");
        Assert.Equal(0x89, bytes[0]);
        Assert.Equal((byte)'P', bytes[1]);
        Assert.Equal((byte)'N', bytes[2]);
        Assert.Equal((byte)'G', bytes[3]);
    }
}
