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
        var slideIdList = doc.PresentationPart!.Presentation.SlideIdList;
        Assert.NotNull(slideIdList);
        Assert.Single(slideIdList!.ChildElements.OfType<SlideId>());
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

    public void Dispose()
    {
        if (File.Exists(_testFilePath))
        {
            File.Delete(_testFilePath);
        }
    }
}
