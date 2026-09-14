using DocumentFormat.OpenXml.Packaging;
using PptxEditor.Core.Converters;
using PptxEditor.Core.Generation.Archetypes;
using PptxEditor.Core.Generation.Components;
using PptxEditor.Core.Generation.Emit.Ooxml;
using PptxEditor.Core.Generation.Layout;
using PptxEditor.Core.Generation.Schema;

namespace DocxEditor.Tests.Generation;

/// <summary>
/// End-to-end speaker-notes flow: generation JSON with slide-level "notes" → parse →
/// (archetype/component expansion) → resolve → OOXML emit → the emitted package carries a
/// NotesSlidePart whose text is readable back through the markdown read path.
/// </summary>
public class NotesPipelineTests
{
    [Fact]
    public void ContainerSlideNotes_SurvivePipeline_AndReadBackViaMarkdown()
    {
        const string json = """
            {
              "version": "2.0",
              "design": { "palette": { "primary": "#0B3D91", "accent": "#FF6B00", "ink": "#1A1A1A", "paper": "#FFFFFF" } },
              "slides": [
                {
                  "type": "container",
                  "layout": { "mode": "column" },
                  "children": [
                    { "type": "text", "text": "Quarterly Results", "fontSize": 30, "color": "primary", "bold": true,
                      "size": { "h": 44 } }
                  ],
                  "notes": "Pause after the title; emphasise the 34% revenue growth."
                }
              ]
            }
            """;

        var document = new GenerationDocumentParser().Parse(json);
        var layout = new LayoutResolver().Resolve(document);
        var emitted = new OoxmlEmitter().Emit(layout);

        using var package = PresentationDocument.Open(new MemoryStream(emitted.Bytes), false);

        // The emitted package carries the notes in a NotesSlidePart.
        var slidePart = package.PresentationPart!.SlideParts.Single();
        var notesSlide = slidePart.NotesSlidePart?.NotesSlide;
        Assert.NotNull(notesSlide);
        Assert.Contains("Pause after the title", notesSlide!.CommonSlideData!.ShapeTree!.InnerText);

        // And the read path surfaces them as a [!note] callout.
        var markdown = PptxToMarkdownConverter.Convert(package);
        Assert.Contains("> [!note]", markdown);
        Assert.Contains("> Pause after the title; emphasise the 34% revenue growth.", markdown);
    }

    [Fact]
    public void ArchetypeSlideNotes_SurviveExpansion_AndEmit()
    {
        const string json = """
            {
              "version": "2.0",
              "design": { "palette": { "primary": "#0B3D91", "accent": "#FF6B00", "ink": "#1A1A1A", "paper": "#FFFFFF" } },
              "slides": [
                {
                  "type": "cover",
                  "content": { "title": "Northwind Labs", "kicker": "Q3 2026" },
                  "notes": "Open with the quarter's headline number."
                }
              ]
            }
            """;

        var document = new GenerationDocumentParser().Parse(json);
        var expanded = ArchetypeExpander.Expand(document);
        var componentized = ComponentExpander.Expand(expanded);
        var layout = new LayoutResolver().Resolve(componentized);
        var emitted = new OoxmlEmitter().Emit(layout);

        using var package = PresentationDocument.Open(new MemoryStream(emitted.Bytes), false);

        var slidePart = package.PresentationPart!.SlideParts.Single();
        var notesSlide = slidePart.NotesSlidePart?.NotesSlide;
        Assert.NotNull(notesSlide);
        Assert.Contains("Open with the quarter's headline number.", notesSlide!.CommonSlideData!.ShapeTree!.InnerText);

        var markdown = PptxToMarkdownConverter.Convert(package);
        Assert.Contains("> [!note]", markdown);
        Assert.Contains("> Open with the quarter's headline number.", markdown);
    }

    [Fact]
    public void DeckWithoutNotes_ProducesNoNotesPartAndNoNoteCallout()
    {
        const string json = """
            {
              "version": "2.0",
              "design": { "palette": { "primary": "#0B3D91" } },
              "slides": [ { "type": "container", "children": [] } ]
            }
            """;

        var document = new GenerationDocumentParser().Parse(json);
        var layout = new LayoutResolver().Resolve(document);
        var emitted = new OoxmlEmitter().Emit(layout);

        using var package = PresentationDocument.Open(new MemoryStream(emitted.Bytes), false);

        Assert.Null(package.PresentationPart!.SlideParts.Single().NotesSlidePart);
        Assert.DoesNotContain("[!note]", PptxToMarkdownConverter.Convert(package));
    }
}
