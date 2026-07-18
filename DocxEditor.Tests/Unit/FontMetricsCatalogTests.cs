using DocumentFormat.OpenXml.Packaging;
using PptxEditor.Core.Builders;
using PptxEditor.Core.Models;
using PptxEditor.Core.Services;
using P = DocumentFormat.OpenXml.Presentation;

namespace DocxEditor.Tests.Unit;

public sealed class FontMetricsCatalogTests : IDisposable
{
    private readonly string _tempDir = SlideOpsTestHelpers.CreateTempDirectory(nameof(FontMetricsCatalogTests));

    public void Dispose() => SlideOpsTestHelpers.BestEffortDelete(_tempDir);

    private static FontMetricsCatalogOptions Hermetic(params (string Family, string Path)[] paths)
        => new() { AdditionalFontPaths = paths, IncludeSystemFonts = false };

    [Fact]
    public void Resolve_RegisteredPath_ReturnsMetrics()
    {
        var font = TextFitTestFont.WriteFont(_tempDir, "testsans.ttf", "TestSans");
        var catalog = new FontMetricsCatalog(options: Hermetic(("TestSans", font)));

        var resolved = catalog.Resolve("TestSans", bold: false, italic: false);

        Assert.NotNull(resolved.Metrics);
        Assert.Equal(TextFitTestFont.UnitsPerEm, resolved.Metrics.UnitsPerEm);
        Assert.False(resolved.WasSubstituted);
        Assert.Empty(resolved.Warnings);
    }

    [Fact]
    public void Resolve_CachesPerFamilyBoldItalicKey()
    {
        var font = TextFitTestFont.WriteFont(_tempDir, "testsans.ttf", "TestSans");
        var catalog = new FontMetricsCatalog(options: Hermetic(("TestSans", font)));

        var first = catalog.Resolve("TestSans", false, false);
        var second = catalog.Resolve("TestSans", false, false);
        var otherVariant = catalog.Resolve("TestSans", true, false);

        Assert.Same(first, second);
        Assert.NotSame(first, otherVariant);
    }

    [Fact]
    public void Resolve_BoldVariantName_PreferredOverRegular()
    {
        var regular = TextFitTestFont.WriteFont(_tempDir, "regular.ttf", "TestSans", advance: 500);
        var bold = TextFitTestFont.WriteFont(_tempDir, "bold.ttf", "TestSans Bold", advance: 600);
        var catalog = new FontMetricsCatalog(options: Hermetic(("TestSans", regular), ("TestSans Bold", bold)));

        var resolved = catalog.Resolve("TestSans", bold: true, italic: false);

        Assert.NotNull(resolved.Metrics);
        Assert.Equal((ushort)600, resolved.Metrics.AdvanceWidths['a']);
    }

    [Fact]
    public void Resolve_AptosWithOnlyCarlitoAvailable_SubstitutesWithWarning()
    {
        var carlito = TextFitTestFont.WriteFont(_tempDir, "carlito.ttf", "Carlito");
        var catalog = new FontMetricsCatalog(options: Hermetic(("Carlito", carlito)));

        var resolved = catalog.Resolve("Aptos", false, false);

        Assert.NotNull(resolved.Metrics);
        Assert.True(resolved.WasSubstituted);
        Assert.Equal("Carlito", resolved.ResolvedFamily);
        Assert.Contains(resolved.Warnings, w => w.Contains("substituted") && w.Contains("not metric-compatible"));
    }

    [Fact]
    public void Resolve_TrueTypeCollection_MarkedUnparseableWithoutCrashing()
    {
        var ttc = TextFitTestFont.WriteTrueTypeCollection(_tempDir);
        var catalog = new FontMetricsCatalog(options: Hermetic(("TtcFont", ttc)));

        var resolved = catalog.Resolve("TtcFont", false, false);

        Assert.Null(resolved.Metrics);
        Assert.Contains(resolved.Warnings, w => w.Contains("TrueType Collection"));
        Assert.Contains(resolved.Warnings, w => w.Contains("unmeasurable"));
    }

    [Fact]
    public void Resolve_UnknownFamily_UnmeasurableWithWarning()
    {
        var catalog = new FontMetricsCatalog(options: new FontMetricsCatalogOptions { IncludeSystemFonts = false });

        var resolved = catalog.Resolve("NoSuchFont", false, false);

        Assert.Null(resolved.Metrics);
        Assert.False(resolved.WasSubstituted);
        Assert.Contains(resolved.Warnings, w => w.Contains("unmeasurable"));
    }

    [Fact]
    public void Resolve_DirectoryScan_ReadsFamilyFromNameTable()
    {
        var fontDir = Path.Combine(_tempDir, "scanme");
        Directory.CreateDirectory(fontDir);
        TextFitTestFont.WriteFont(fontDir, "whatever-filename.ttf", "ScanSans");
        var catalog = new FontMetricsCatalog(options: new FontMetricsCatalogOptions
        {
            AdditionalFontDirectories = [fontDir],
            IncludeSystemFonts = false
        });

        var resolved = catalog.Resolve("ScanSans", false, false);

        Assert.NotNull(resolved.Metrics);
        Assert.Equal(TextFitTestFont.UnitsPerEm, resolved.Metrics.UnitsPerEm);
    }

    [Fact]
    public void Resolve_EmbeddedFonts_KeyedByDeclaredVariant()
    {
        var path = CreateDeckWithEmbeddedFont();

        using var doc = PresentationDocument.Open(path, false);
        var catalog = new FontMetricsCatalog(doc, new FontMetricsCatalogOptions { IncludeSystemFonts = false });

        var regular = catalog.Resolve("EmbedSans", false, false);
        var bold = catalog.Resolve("EmbedSans", true, false);

        Assert.NotNull(regular.Metrics);
        Assert.Equal((ushort)500, regular.Metrics.AdvanceWidths['a']);
        Assert.NotNull(bold.Metrics);
        Assert.Equal((ushort)600, bold.Metrics.AdvanceWidths['a']);
    }

    private string CreateDeckWithEmbeddedFont()
    {
        var path = Path.Combine(_tempDir, $"{Guid.NewGuid():N}.pptx");
        using (var builder = PresentationBuilder.Create(path))
        {
            builder.AddSlide();
            builder.Save();
        }

        using (var doc = PresentationDocument.Open(path, true))
        {
            var presentationPart = doc.PresentationPart!;
            var presentation = presentationPart.Presentation!;

            var regularPart = presentationPart.AddNewPart<FontPart>("application/x-fontdata", "rIdTestFontRegular");
            using (var stream = new MemoryStream(TextFitTestFont.BuildFont("EmbedSans", advance: 500)))
                regularPart.FeedData(stream);

            var boldPart = presentationPart.AddNewPart<FontPart>("application/x-fontdata", "rIdTestFontBold");
            using (var stream = new MemoryStream(TextFitTestFont.BuildFont("EmbedSans Bold", advance: 600)))
                boldPart.FeedData(stream);

            const string ns = "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" " +
                              "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                              "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";
            var listXml = $"<p:embeddedFontLst {ns}>" +
                          "<p:embeddedFont>" +
                          "<p:font typeface=\"EmbedSans\"/>" +
                          "<p:regular r:id=\"rIdTestFontRegular\"/>" +
                          "<p:bold r:id=\"rIdTestFontBold\"/>" +
                          "</p:embeddedFont>" +
                          "</p:embeddedFontLst>";
            presentation.AppendChild(new P.EmbeddedFontList(listXml));
            presentation.Save();
        }

        return path;
    }
}
