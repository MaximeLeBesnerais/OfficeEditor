using PptxEditor.Core.Builders;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace DocxEditor.Tests.Unit;

public class PresentationBuilderTests : IDisposable
{
    private readonly string _testFilePath = Path.Combine(Path.GetTempPath(), $"test_pptx_{Guid.NewGuid()}.pptx");

    [Fact]
    public void Create_ShouldCreateNewPresentation()
    {
        // Act
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.Save();
        }

        // Assert
        Assert.True(File.Exists(_testFilePath));
        using var doc = PresentationDocument.Open(_testFilePath, false);
        Assert.NotNull(doc.PresentationPart);
    }

    [Fact]
    public void AddSlide_ShouldAddSlideToPresentation()
    {
        // Act
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.Save();
        }

        // Assert
        using var doc = PresentationDocument.Open(_testFilePath, false);
        var presentationPart = doc.PresentationPart;
        Assert.NotNull(presentationPart);
        var presentation = presentationPart.Presentation;
        Assert.NotNull(presentation);
        var slideIdList = presentation.SlideIdList;
        Assert.NotNull(slideIdList);
        Assert.Single(slideIdList.ChildElements.OfType<SlideId>());
    }

    [Fact]
    public void AddSlide_WithTitle_ShouldAddTitle()
    {
        // Act
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("Test Title");
            builder.Save();
        }

        // Assert
        using var doc = PresentationDocument.Open(_testFilePath, false);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var text = slidePart.Slide!.InnerText;
        Assert.Contains("Test Title", text);
    }

    [Fact]
    public void AddSlide_WithContent_ShouldAddContent()
    {
        // Act
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("Title");
            builder.CurrentSlide.AddText("Body text");
            builder.Save();
        }

        // Assert
        using var doc = PresentationDocument.Open(_testFilePath, false);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var text = slidePart.Slide!.InnerText;
        Assert.Contains("Title", text);
        Assert.Contains("Body text", text);
    }

    [Fact]
    public void AddSlide_WithBulletList_ShouldAddList()
    {
        // Act
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("List");
            builder.CurrentSlide.AddBulletList(new[] { "Item 1", "Item 2", "Item 3" });
            builder.Save();
        }

        // Assert
        using var doc = PresentationDocument.Open(_testFilePath, false);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var text = slidePart.Slide!.InnerText;
        Assert.Contains("Item 1", text);
        Assert.Contains("Item 2", text);
        Assert.Contains("Item 3", text);
    }

    [Fact]
    public void CurrentSlide_BeforeAddingSlide_ShouldThrow()
    {
        using var builder = PresentationBuilder.Create(_testFilePath);

        Assert.Throws<InvalidOperationException>(() => builder.CurrentSlide);
    }

    [Fact]
    public void GetSlide_AndRemoveSlide_ShouldValidateIndexes()
    {
        using var builder = PresentationBuilder.Create(_testFilePath);
        builder.AddSlide().AddSlide();

        Assert.Throws<ArgumentOutOfRangeException>(() => builder.GetSlide(-1));
        Assert.Throws<ArgumentOutOfRangeException>(() => builder.GetSlide(2));
        Assert.Throws<ArgumentOutOfRangeException>(() => builder.RemoveSlide(2));

        builder.RemoveSlide(1);

        Assert.Equal(1, builder.SlideCount);
    }

    [Fact]
    public void ReorderSlide_ShouldMoveSlideIdsAndValidateBounds()
    {
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("First");
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("Second");

            Assert.Throws<ArgumentOutOfRangeException>(() => builder.ReorderSlide(0, 2));

            builder.ReorderSlide(0, 1);
            builder.Save();
        }

        using var doc = PresentationDocument.Open(_testFilePath, false);
        var slideIds = doc.PresentationPart!.Presentation!.SlideIdList!.ChildElements.OfType<SlideId>().ToList();
        var firstSlide = (SlidePart)doc.PresentationPart.GetPartById(slideIds[0].RelationshipId!);

        Assert.Contains("Second", firstSlide.Slide!.InnerText);
    }

    [Fact]
    public void AddSlide_WithMissingLayoutName_ShouldFallBackToFirstLayout()
    {
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide("does-not-exist");
            builder.CurrentSlide.AddTitle("Fallback layout");
            builder.Save();
        }

        using var doc = PresentationDocument.Open(_testFilePath, false);
        var slidePart = doc.PresentationPart!.SlideParts.Single();
        Assert.NotNull(slidePart.SlideLayoutPart);
        Assert.Contains("Fallback layout", slidePart.Slide!.InnerText);
    }

    [Fact]
    public void SlideBuilder_EasyErrorAndContentBranches_ShouldBeDeterministic()
    {
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            var slide = builder.CurrentSlide;

            Assert.Throws<FileNotFoundException>(() => slide.AddImage(Path.Combine(Path.GetTempPath(), "missing-image.png")));
            Assert.Throws<ArgumentException>(() => slide.AddTable([]));

            slide.AddSubtitle("Subtitle")
                .AddNumberedList(["One", "Two"])
                .AddChart(ChartType.Pie, new Dictionary<string, int> { ["A"] = 1, ["B"] = 2 });
            builder.Save();
        }

        using var doc = PresentationDocument.Open(_testFilePath, false);
        var text = doc.PresentationPart!.SlideParts.Single().Slide!.InnerText;
        Assert.Contains("Subtitle", text);
        Assert.Contains("One", text);
        Assert.Contains("[Pie Chart: A=1, B=2]", text);
    }

    [Fact]
    public void AddSlide_WithTable_ShouldAddTable()
    {
        // Act
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("Table");
            builder.CurrentSlide.AddTable(new List<List<string>>
            {
                new() { "Header 1", "Header 2" },
                new() { "Cell 1", "Cell 2" }
            });
            builder.Save();
        }

        // Assert
        using var doc = PresentationDocument.Open(_testFilePath, false);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var text = slidePart.Slide!.InnerText;
        Assert.Contains("Header 1", text);
        Assert.Contains("Cell 2", text);
    }

    [Fact]
    public void DetectVariables_ShouldFindVariables()
    {
        // Arrange
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("Hello {{name}}");
            builder.Save();
        }

        // Act
        List<OfficeEditor.Core.Models.VariableInfo> variables;
        using (var builder = PresentationBuilder.Open(_testFilePath))
        {
            variables = builder.DetectVariables();
        }

        // Assert
        Assert.Single(variables);
        Assert.Equal("name", variables[0].Name);
    }

    [Fact]
    public void DetectVariables_ShouldTrimDefaultsAndDeduplicateWithinShapeLocation()
    {
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("Hello {{ name |Guest}} and {{name|Ignored}}");
            builder.Save();
        }

        using var opened = PresentationBuilder.Open(_testFilePath);
        var variables = opened.DetectVariables();

        var variable = Assert.Single(variables);
        Assert.Equal("name", variable.Name);
        Assert.Equal("Guest", variable.DefaultValue);
        Assert.Contains("slide:1:shape:Title", variable.Location);
    }

    [Fact]
    public void MergeVariables_ShouldReplaceVariables()
    {
        // Arrange
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("Hello {{name}}");
            builder.Save();
        }

        // Act
        using (var builder = PresentationBuilder.Open(_testFilePath))
        {
            builder.MergeVariables(new Dictionary<string, string>
            {
                ["name"] = "World"
            });
            builder.Save();
        }

        // Assert
        using var doc = PresentationDocument.Open(_testFilePath, false);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var text = slidePart.Slide!.InnerText;
        Assert.Contains("Hello World", text);
        Assert.DoesNotContain("{{name}}", text);
    }

    [Fact]
    public void MergeVariables_ShouldUseDefaultsAndLeaveUnprovidedVariables()
    {
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("{{known}} {{missing|Fallback}} {{unprovided}}");
            builder.Save();
        }

        using (var builder = PresentationBuilder.Open(_testFilePath))
        {
            builder.MergeVariables(new Dictionary<string, string> { ["known"] = "Value" });
            builder.Save();
        }

        using var doc = PresentationDocument.Open(_testFilePath, false);
        var text = doc.PresentationPart!.SlideParts.First().Slide!.InnerText;
        Assert.Contains("Value Fallback {{unprovided}}", text);
    }

    [Fact]
    public void Create_SaveToBytes_ReturnsValidPptx()
    {
        // Arrange
        byte[] bytes;

        // Act
        using (var builder = PresentationBuilder.Create())
        {
            builder.AddSlide();
            bytes = builder.SaveToBytes();
        }

        // Assert
        Assert.NotEmpty(bytes);
        using var stream = new MemoryStream(bytes);
        using var doc = PresentationDocument.Open(stream, false);
        Assert.NotNull(doc.PresentationPart);
        Assert.Single(doc.PresentationPart!.Presentation!.SlideIdList!.ChildElements.OfType<SlideId>());
    }

    [Fact]
    public void Create_Save_Stream_ReturnsValidPptx()
    {
        // Arrange
        using var outputStream = new MemoryStream();

        // Act
        using (var builder = PresentationBuilder.Create())
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("Stream title");
            builder.Save(outputStream);
        }

        // Assert
        Assert.True(outputStream.CanRead);
        Assert.True(outputStream.CanWrite);
        Assert.True(outputStream.Length > 0);
        outputStream.Position = 0;
        using var doc = PresentationDocument.Open(outputStream, false);
        Assert.NotNull(doc.PresentationPart);
        var slide = doc.PresentationPart!.SlideParts.First();
        Assert.Contains("Stream title", slide.Slide!.InnerText);
    }

    [Fact]
    public void Open_FromBytes_AndSaveToBytes_RoundtripsSlide()
    {
        // Arrange
        byte[] originalBytes;
        using (var builder = PresentationBuilder.Create())
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("Original title");
            originalBytes = builder.SaveToBytes();
        }

        // Act
        byte[] modifiedBytes;
        using (var builder = PresentationBuilder.Open(originalBytes))
        {
            builder.GetSlide(0).AddText("Added text");
            modifiedBytes = builder.SaveToBytes();
        }

        // Assert
        using var stream = new MemoryStream(modifiedBytes);
        using var doc = PresentationDocument.Open(stream, false);
        var slide = doc.PresentationPart!.SlideParts.First();
        Assert.Contains("Original title", slide.Slide!.InnerText);
        Assert.Contains("Added text", slide.Slide.InnerText);
    }

    [Fact]
    public void Save_WithNoPath_OnPathlessDocument_ThrowsInvalidOperationException()
    {
        // Arrange
        using var builder = PresentationBuilder.Create();
        builder.AddSlide();

        // Act & Assert
        Assert.Throws<InvalidOperationException>(() => builder.Save());
    }

    public void Dispose()
    {
        if (File.Exists(_testFilePath))
        {
            File.Delete(_testFilePath);
        }
    }
}
