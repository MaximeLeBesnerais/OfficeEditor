using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Generation;
using DocxEditor.Core.Generation.Contracts;

namespace DocxEditor.Tests.Generation.Docx;

/// <summary>
/// Style-preservation coverage: template styles are referenced by ID and never mutated, and the
/// generator lazily creates theme-derived baseline styles (title, subtitle, eyebrow, headings,
/// muted body, labels, metrics, table header/body, callout, footer) with collision-free IDs.
/// </summary>
public class DocxStylePreservationTests
{
    private static readonly string TemplateDir = Path.Combine(Path.GetTempPath(), "officeeditor-tests");

    [Fact]
    public void TemplateStyles_AreNeverMutated_AndReferencedById()
    {
        var templatePath = CreateTemplateWithCustomStyle();

        var json = """
            {"version":"1.0","sections":[{"blocks":[
              {"type":"paragraph","style":"MyStyle","text":"Styled by template"},
              {"type":"paragraph","role":"subtitle","text":"Under the title"}
            ]}]}
            """;

        var templateStyleXml = ReadStyleOuterXml(templatePath, "MyStyle");

        var generated = new DocxGenerator().GenerateToBytes(json, new DocxGeneratorOptions { TemplatePath = templatePath });
        var doc = OpenDocument(generated.Content);
        var stylesPart = doc.MainDocumentPart!.StyleDefinitionsPart!;

        // The template's definition is untouched.
        var preserved = stylesPart.Styles!.Elements<Style>().Single(s => s.StyleId?.Value == "MyStyle");
        Assert.Equal(templateStyleXml, preserved.OuterXml);

        // The body paragraph references the template style by ID (no redefinition).
        var styledParagraph = doc.MainDocumentPart!.Document!.Body!
            .Elements<Paragraph>()
            .First(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "MyStyle");
        Assert.NotNull(styledParagraph);

        // A semantic role that the template lacks is generated with a theme-derived definition.
        var subtitle = stylesPart.Styles.Elements<Style>().Single(s => s.StyleId?.Value == "Subtitle");
        Assert.Equal(StyleValues.Paragraph, subtitle.Type?.Value);
        Assert.Equal("Arial", subtitle.StyleRunProperties!.RunFonts!.Ascii!.Value);
        Assert.Equal("5A6B7B", subtitle.StyleRunProperties.Color!.Val!.Value);

        Assert.Empty(generated.Result.Warnings);

        Cleanup(templatePath);
    }

    [Fact]
    public void TemplateNormalDefault_IsNeverOverridden()
    {
        var templatePath = CreateTemplateWithCustomStyle();

        var json = """
            {"version":"1.0","sections":[{"blocks":[{"type":"paragraph","text":"Plain body"}]}]}
            """;

        var templateNormalXml = ReadStyleOuterXml(templatePath, "Normal");
        var generated = new DocxGenerator().GenerateToBytes(json, new DocxGeneratorOptions { TemplatePath = templatePath });
        var doc = OpenDocument(generated.Content);

        var normal = doc.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<Style>()
            .Single(s => s.StyleId?.Value == "Normal");
        Assert.Equal(templateNormalXml, normal.OuterXml);

        Cleanup(templatePath);
    }

    [Fact]
    public void BlankDocument_GeneratesThemeDefaultsAndRoleStyles()
    {
        var json = """
            {"version":"1.0","sections":[{"blocks":[
              {"type":"heading","level":1,"text":"Operations Brief"},
              {"type":"paragraph","text":"Body copy"},
              {"type":"paragraph","role":"eyebrow","text":"QUARTERLY"},
              {"type":"paragraph","role":"metric","text":"+34%"},
              {"type":"paragraph","role":"footer","text":"page footer"}
            ]}]}
            """;

        var generated = new DocxGenerator().GenerateToBytes(json);
        Assert.Empty(generated.Result.Warnings);
        var doc = OpenDocument(generated.Content);
        var styles = doc.MainDocumentPart!.StyleDefinitionsPart!.Styles!;

        // The document defaults are the editorial body font.
        var defaults = styles.DocDefaults!.RunPropertiesDefault!.RunPropertiesBaseStyle!;
        Assert.Equal("Arial", defaults.RunFonts!.Ascii!.Value);

        Assert.Contains(styles.Elements<Style>(), s => s.StyleId?.Value == "Heading1");
        Assert.Contains(styles.Elements<Style>(), s => s.StyleId?.Value == "Eyebrow");
        Assert.Contains(styles.Elements<Style>(), s => s.StyleId?.Value == "Metric");
        Assert.Contains(styles.Elements<Style>(), s => s.StyleId?.Value == "Footer");

        var eyebrow = styles.Elements<Style>().Single(s => s.StyleId?.Value == "Eyebrow");
        Assert.Equal(StyleValues.Paragraph, eyebrow.Type?.Value);
        Assert.Contains(eyebrow.StyleRunProperties!.ChildElements, e => e is Caps);

        // Headings keep with the next paragraph and carry the theme color.
        var heading1 = styles.Elements<Style>().Single(s => s.StyleId?.Value == "Heading1");
        Assert.Contains(heading1.StyleParagraphProperties!.ChildElements, e => e is KeepNext);
        Assert.Equal("1F3A5F", heading1.StyleRunProperties!.Color!.Val!.Value);
    }

    [Fact]
    public void GeneratedRoleStyleIds_AreCollisionFreeWhenTemplateCollides()
    {
        var templatePath = CreateTemplateWithCollidingStyleIds();

        var json = """
            {"version":"1.0","sections":[{"blocks":[
              {"type":"paragraph","role":"subtitle","text":"Sub"},
              {"type":"paragraph","role":"label","text":"Label"}
            ]}]}
            """;

        var generated = new DocxGenerator().GenerateToBytes(json, new DocxGeneratorOptions { TemplatePath = templatePath });
        var doc = OpenDocument(generated.Content);
        var styles = doc.MainDocumentPart!.StyleDefinitionsPart!.Styles!;

        var ids = styles.Elements<Style>().Select(s => s.StyleId!.Value!).ToList();
        Assert.Equal(ids.Count, ids.Distinct().Count());

        Cleanup(templatePath);
    }

    // ---- helpers ----

    private static string CreateTemplateWithCustomStyle()
    {
        var path = NextTemplatePath();
        using (var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(new Run(new Text("Template body")))));
            var styles = main.AddNewPart<StyleDefinitionsPart>();
            styles.Styles = new Styles(
                new Style
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "Normal",
                    Default = true,
                    StyleName = new StyleName { Val = "Normal" },
                    StyleRunProperties = new StyleRunProperties(new RunFonts { Ascii = "Courier New", HighAnsi = "Courier New" })
                },
                new Style
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "MyStyle",
                    StyleName = new StyleName { Val = "My Style" },
                    BasedOn = new BasedOn { Val = "Normal" },
                    StyleParagraphProperties = new StyleParagraphProperties(new SpacingBetweenLines { After = "480" }),
                    StyleRunProperties = new StyleRunProperties(
                        new RunFonts { Ascii = "Georgia", HighAnsi = "Georgia" },
                        new Color { Val = "00FF00" })
                });
            document.Save();
        }
        return path;
    }

    /// <summary>A template whose style IDs collide with the conventional baseline IDs.</summary>
    private static string CreateTemplateWithCollidingStyleIds()
    {
        var path = NextTemplatePath();
        using (var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(new Run(new Text("Template body")))));
            var styles = main.AddNewPart<StyleDefinitionsPart>();
            styles.Styles = new Styles(
                new Style
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "Subtitle",
                    StyleName = new StyleName { Val = "Not the baseline" },
                    StyleRunProperties = new StyleRunProperties(new Color { Val = "123456" })
                },
                new Style
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "Label",
                    StyleName = new StyleName { Val = "Not the baseline" },
                    StyleRunProperties = new StyleRunProperties(new Color { Val = "654321" })
                });
            document.Save();
        }
        return path;
    }

    private static string NextTemplatePath()
    {
        Directory.CreateDirectory(TemplateDir);
        return Path.Combine(TemplateDir, $"template-{Guid.NewGuid():N}.docx");
    }

    private static string ReadStyleOuterXml(string path, string styleId)
    {
        using var document = WordprocessingDocument.Open(path, false);
        return document.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<Style>()
            .Single(s => s.StyleId?.Value == styleId)
            .OuterXml;
    }

    private static WordprocessingDocument OpenDocument(byte[] bytes)
    {
        var stream = new MemoryStream(bytes);
        return WordprocessingDocument.Open(stream, false);
    }

    private static void Cleanup(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
            // Best-effort cleanup must never mask the primary assertion.
        }
    }
}
