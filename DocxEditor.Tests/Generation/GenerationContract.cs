using DocumentFormat.OpenXml.Packaging;

namespace OfficeEditor.Testing;

internal static class GenerationContract
{
    internal const string Json = """
        {
          "version": "2.0",
          "slideSize": { "width": 800, "height": 600 },
          "design": {
            "palette": { "primary": "#0B3D91", "accent": "#FF6B00", "ink": "#1A1A1A", "paper": "#FFFFFF", "muted": "#8A94A6" }
          },
          "slides": [
            { "type": "cover", "id": "intro", "notes": "Speaker notes survive generation.",
              "content": { "title": "Shared generation", "subtitle": "One pipeline" } },
            { "type": "container", "id": "detail", "children": [
              { "type": "text", "id": "fitted", "text": "Long text that needs to fit into a small box",
                "fontSize": 30, "color": "#123456", "overflow": "shrink",
                "at": { "x": 10, "y": 10 }, "size": { "w": 40, "h": 15 } }
            ] }
          ]
        }
        """;

    // Compare delivery semantics, excluding ZIP metadata, timestamps and package relationship ids.
    internal static string[] Snapshot(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = PresentationDocument.Open(stream, false);
        var part = document.PresentationPart!;
        var snapshot = new List<string> { part.Presentation!.SlideSize!.OuterXml };
        foreach (var slideId in part.Presentation!.SlideIdList!.Elements<DocumentFormat.OpenXml.Presentation.SlideId>())
        {
            var slide = (SlidePart)part.GetPartById(slideId.RelationshipId!);
            snapshot.Add(slide.Slide!.OuterXml);
            snapshot.Add(slide.NotesSlidePart?.NotesSlide!.OuterXml ?? "");
        }
        return snapshot.ToArray();
    }
}
