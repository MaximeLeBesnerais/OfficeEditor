using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PptxEditor.Core.Builders;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Shared helpers for slide-operation tests (reorder, duplicate, remove).
/// </summary>
internal static class SlideOpsTestHelpers
{
    public static string CreateTempDirectory(string testClassName)
    {
        var directory = Path.Combine(Path.GetTempPath(), testClassName, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        return directory;
    }

    public static void BestEffortDelete(string directory)
    {
        try
        {
            if (Directory.Exists(directory))
            {
                Directory.Delete(directory, true);
            }
        }
        catch
        {
            // Best-effort cleanup.
        }
    }

    /// <summary>Creates a deck where each slide carries a single identifying title.</summary>
    public static string CreateDeckWithTitledSlides(string directory, params string[] titles)
    {
        var path = Path.Combine(directory, $"{Guid.NewGuid():N}.pptx");
        using (var builder = PresentationBuilder.Create(path))
        {
            foreach (var title in titles)
            {
                builder.AddSlide();
                builder.CurrentSlide.AddTitle(title);
            }

            builder.Save();
        }

        return path;
    }

    /// <summary>Returns the InnerText of each slide in SlideIdList order.</summary>
    public static List<string> ReadSlideTextsInOrder(string path)
    {
        using var doc = PresentationDocument.Open(path, false);
        return ReadSlideTextsInOrder(doc);
    }

    public static List<string> ReadSlideTextsInOrder(PresentationDocument doc)
    {
        var texts = new List<string>();
        var slideIdList = doc.PresentationPart!.Presentation!.SlideIdList!;

        foreach (var slideId in slideIdList.ChildElements.OfType<SlideId>())
        {
            var slidePart = (SlidePart)doc.PresentationPart.GetPartById(slideId.RelationshipId!);
            texts.Add(slidePart.Slide!.InnerText);
        }

        return texts;
    }

    public static List<SlidePart> GetSlidePartsInOrder(PresentationDocument doc)
    {
        var parts = new List<SlidePart>();
        var slideIdList = doc.PresentationPart!.Presentation!.SlideIdList!;

        foreach (var slideId in slideIdList.ChildElements.OfType<SlideId>())
        {
            parts.Add((SlidePart)doc.PresentationPart.GetPartById(slideId.RelationshipId!));
        }

        return parts;
    }

    /// <summary>Depth-first enumeration of every OpenXmlPart reachable from the root.</summary>
    public static List<OpenXmlPart> EnumerateAllParts(OpenXmlPartContainer root)
    {
        var parts = new List<OpenXmlPart>();
        var visited = new HashSet<OpenXmlPartContainer>();
        var stack = new Stack<OpenXmlPartContainer>();
        stack.Push(root);

        while (stack.Count > 0)
        {
            var container = stack.Pop();
            if (!visited.Add(container))
            {
                continue;
            }

            foreach (var pair in container.Parts)
            {
                parts.Add(pair.OpenXmlPart);
                stack.Push(pair.OpenXmlPart);
            }
        }

        return parts;
    }

    /// <summary>Writes a 1x1 transparent PNG and returns its path.</summary>
    public static string WriteMinimalPng(string directory)
    {
        var bytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");
        var path = Path.Combine(directory, $"{Guid.NewGuid():N}.png");
        File.WriteAllBytes(path, bytes);
        return path;
    }
}
