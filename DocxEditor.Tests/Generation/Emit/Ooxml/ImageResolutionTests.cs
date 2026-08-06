using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocxEditor.Tests.Generation.Docx;
using OfficeEditor.Core.Rendering;
using PptxEditor.Core.Generation.Archetypes;
using PptxEditor.Core.Generation.Components;
using PptxEditor.Core.Generation.Emit.Ooxml;
using PptxEditor.Core.Generation.Layout;
using PptxEditor.Core.Generation.Model;
using PptxEditor.Core.Generation.Schema;
using PptxEditor.Core.Models;

namespace DocxEditor.Tests.Generation.Emit.Ooxml;

/// <summary>
/// Asset-resolution acceptance for the ratified doc-relative design: relative image
/// sources resolve against the JSON document's directory ONLY (no repository-root or CWD
/// fallback), with canonical/symlink-aware containment; absolute paths and data URIs pass
/// through unchanged; string-only surfaces (no document directory) get a loud, actionable
/// error. Also covers the DOCX facade threading the document directory as the allowed root.
/// </summary>
public sealed class ImageResolutionTests
{
    #region PPTX: document-directory resolution

    [Fact]
    public void RelativeSource_ResolvesAgainstDocumentDirectory_EndToEnd()
    {
        using var temp = new TempDirectory();
        var imageBytes = TestImages.Png(64, 32);
        File.WriteAllBytes(temp.File("hero.png"), imageBytes);

        var deckPath = temp.File("deck.json");
        File.WriteAllText(deckPath, DeckJsonWithImage("hero.png"));

        var layout = ResolveDeck(File.ReadAllText(deckPath));
        var emission = new OoxmlEmitter(new OoxmlEmitOptions
        {
            DocumentDirectory = Path.GetDirectoryName(Path.GetFullPath(deckPath))
        }).Emit(layout);

        Assert.Single(layout.Slides);
        AssertEmbeddedImageBytes(emission.Bytes, imageBytes);
    }

    [Fact]
    public void RelativeSource_InNestedAssetsDirectory_ResolvesAgainstDocumentDirectory()
    {
        // Mirrors the demo decks: demo/deck.json uses "src": "assets/…" relative to the
        // JSON file, never a repo-root-relative path.
        using var temp = new TempDirectory();
        var assetsDir = Path.Combine(temp.Path, "assets");
        Directory.CreateDirectory(assetsDir);
        var imageBytes = TestImages.Png(32, 32);
        File.WriteAllBytes(Path.Combine(assetsDir, "dashboard.png"), imageBytes);

        var deckPath = temp.File("deck.json");
        File.WriteAllText(deckPath, DeckJsonWithImage("assets/dashboard.png"));

        var layout = ResolveDeck(File.ReadAllText(deckPath));
        var emission = new OoxmlEmitter(new OoxmlEmitOptions
        {
            DocumentDirectory = Path.GetDirectoryName(Path.GetFullPath(deckPath))
        }).Emit(layout);

        AssertEmbeddedImageBytes(emission.Bytes, imageBytes);
    }

    [Fact]
    public void RelativeSource_WithoutDocumentDirectory_ThrowsLoudActionableError()
    {
        var layout = LayoutWithImageSource("hero.png");

        var ex = Assert.Throws<ArgumentException>(() => new OoxmlEmitter().Emit(layout));

        Assert.Contains("relative path", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("document's directory", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("data URI", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RelativeSource_EscapingDocumentDirectory_Throws()
    {
        using var temp = new TempDirectory();
        var deckDir = Path.Combine(temp.Path, "deckdir");
        Directory.CreateDirectory(deckDir);
        File.WriteAllBytes(Path.Combine(temp.Path, "outside.png"), TestImages.Png(16, 16));

        var layout = LayoutWithImageSource("../outside.png");

        var ex = Assert.Throws<ArgumentException>(() =>
            new OoxmlEmitter(new OoxmlEmitOptions { DocumentDirectory = deckDir }).Emit(layout));

        Assert.Contains("outside the document directory", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RelativeSource_SymlinkEscape_Throws()
    {
        using var temp = new TempDirectory();
        var deckDir = Path.Combine(temp.Path, "deckdir");
        Directory.CreateDirectory(deckDir);
        var outside = Path.Combine(temp.Path, "outside.png");
        File.WriteAllBytes(outside, TestImages.Png(16, 16));

        // A link INSIDE the document directory pointing OUTSIDE it must be rejected by
        // the canonical (symlink-aware) containment check, even though the lexical path
        // stays inside.
        var link = Path.Combine(deckDir, "link.png");
        try
        {
            File.CreateSymbolicLink(link, outside);
        }
        catch (Exception linkEx) when (linkEx is IOException or UnauthorizedAccessException or PlatformNotSupportedException)
        {
            // Environments without symlink privileges can't exercise this path; the test
            // still runs on macOS/Linux CI where links are creatable.
            return;
        }

        var layout = LayoutWithImageSource("link.png");

        var ex = Assert.Throws<ArgumentException>(() =>
            new OoxmlEmitter(new OoxmlEmitOptions { DocumentDirectory = deckDir }).Emit(layout));

        Assert.Contains("symbolic link", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RelativeSource_MissingFileInsideDocumentDirectory_ThrowsFileNotFound()
    {
        using var temp = new TempDirectory();
        var layout = LayoutWithImageSource("missing.png");

        Assert.Throws<FileNotFoundException>(() =>
            new OoxmlEmitter(new OoxmlEmitOptions { DocumentDirectory = temp.Path }).Emit(layout));
    }

    [Fact]
    public void AbsoluteSource_IsReadAsIs_RegardlessOfDocumentDirectory()
    {
        using var temp = new TempDirectory();
        var imageBytes = TestImages.Png(24, 24);
        var imagePath = temp.File("absolute.png");
        File.WriteAllBytes(imagePath, imageBytes);

        // DocumentDirectory points somewhere the image does NOT live — absolute paths must
        // be unaffected by it.
        var layout = LayoutWithImageSource(imagePath);
        var emission = new OoxmlEmitter(new OoxmlEmitOptions
        {
            DocumentDirectory = Path.Combine(temp.Path, "elsewhere")
        }).Emit(layout);

        AssertEmbeddedImageBytes(emission.Bytes, imageBytes);
    }

    [Fact]
    public void DataUriSource_IsReadAsIs()
    {
        var imageBytes = TestImages.Png(20, 20);
        var layout = LayoutWithImageSource(TestImages.DataUriBase64(imageBytes));

        var emission = new OoxmlEmitter(new OoxmlEmitOptions
        {
            DocumentDirectory = Path.GetTempPath()
        }).Emit(layout);

        AssertEmbeddedImageBytes(emission.Bytes, imageBytes);
    }

    #endregion

    #region DOCX: facade threads the document directory as the allowed root

    [Fact]
    public void DocxFacade_RelativeImage_ResolvesAgainstDocumentDirectory()
    {
        using var temp = new TempDirectory();
        File.WriteAllBytes(temp.File("photo.png"), TestImages.Png(64, 32));

        var jsonPath = temp.File("report.doc.json");
        File.WriteAllText(jsonPath, """
            {
              "version": "1.0",
              "design": { "palette": { "ink": "#1F2937" } },
              "sections": [ { "blocks": [
                { "type": "image", "src": "photo.png", "width": 50, "height": 50 }
              ] } ]
            }
            """);

        // Typ = Typst source export, which needs no Typst backend, so this runs without
        // the OE_RUN_TYPST_COMPILE_TESTS gate. The relative "src" resolves only if the
        // facade threads the JSON's directory as the AllowedRoot (the test runner's CWD
        // does not contain photo.png).
        var result = new DocumentRenderer().Render(new DocumentRenderRequest
        {
            SourcePath = jsonPath,
            Format = DocumentOutputFormat.Typ
        });

        Assert.True(result.Success, result.ErrorMessage);
        var source = Encoding.UTF8.GetString(Assert.Single(result.Pages));
        Assert.Contains("image(", source, StringComparison.Ordinal);
    }

    [Fact]
    public void DocxFacade_RelativeImageOutsideDocumentDirectory_Fails()
    {
        using var temp = new TempDirectory();
        var deckDir = Path.Combine(temp.Path, "deckdir");
        Directory.CreateDirectory(deckDir);
        // The image EXISTS, but outside the JSON document's directory: hard-root
        // containment must reject it even though a CWD fallback would have found it.
        File.WriteAllBytes(temp.File("photo.png"), TestImages.Png(64, 32));

        var jsonPath = Path.Combine(deckDir, "report.doc.json");
        File.WriteAllText(jsonPath, """
            {
              "version": "1.0",
              "design": { "palette": { "ink": "#1F2937" } },
              "sections": [ { "blocks": [
                { "type": "image", "src": "../photo.png", "width": 50, "height": 50 }
              ] } ]
            }
            """);

        var result = new DocumentRenderer().Render(new DocumentRenderRequest
        {
            SourcePath = jsonPath,
            Format = DocumentOutputFormat.Typ
        });

        Assert.False(result.Success);
        Assert.Contains("outside", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Helpers

    private static string DeckJsonWithImage(string src) => $$"""
        {
          "version": "2.0",
          "design": {
            "palette": { "primary": "#0B3D91", "accent": "#FF6B00", "ink": "#1A1A1A", "paper": "#FFFFFF" },
            "fonts": { "display": "Aptos Display", "body": "Aptos" },
            "metrics": { "marginPt": 40, "gutterPt": 16, "titleSizePt": 30, "bodySizePt": 14 }
          },
          "slides": [
            {
              "type": "container",
              "fill": "paper",
              "padding": 40,
              "layout": { "mode": "column", "gap": 16 },
              "children": [
                { "type": "image", "src": "{{src}}", "fit": "fill", "size": { "w": 200, "h": 100 } }
              ]
            }
          ]
        }
        """;

    private static LayoutResult ResolveDeck(string json)
    {
        var validation = new GenerationDocumentParser().Validate(json);
        Assert.True(validation.IsValid,
            string.Join(Environment.NewLine, validation.Errors.Select(error => error.ToString())));
        var archetyped = ArchetypeExpander.Expand(validation.Document!);
        var componentized = ComponentExpander.Expand(archetyped);
        return new LayoutResolver().Resolve(componentized);
    }

    private static LayoutResult LayoutWithImageSource(string source) => new()
    {
        Slides =
        [
            new ResolvedSlide
            {
                WidthPt = 960,
                HeightPt = 540,
                Root = new ResolvedContainer
                {
                    X = 0,
                    Y = 0,
                    Width = 960,
                    Height = 540,
                    Overflow = OverflowPolicy.Error,
                    Children =
                    [
                        new ResolvedImage
                        {
                            X = 0, Y = 0, Width = 100, Height = 100,
                            Source = source, Fit = ImageFitMode.Fill
                        }
                    ]
                }
            }
        ],
        Warnings = []
    };

    private static void AssertEmbeddedImageBytes(byte[] pptxBytes, byte[] expectedImageBytes)
    {
        using var document = PresentationDocument.Open(new MemoryStream(pptxBytes), false);
        var slidePart = document.PresentationPart!.SlideParts.First();
        var imagePart = Assert.Single(slidePart.ImageParts);
        using var stream = imagePart.GetStream();
        using var payload = new MemoryStream();
        stream.CopyTo(payload);
        Assert.Equal(expectedImageBytes, payload.ToArray());
    }

    #endregion
}
