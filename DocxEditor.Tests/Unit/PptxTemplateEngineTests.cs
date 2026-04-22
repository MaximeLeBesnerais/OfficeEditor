using PptxEditor.Core.Builders;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace DocxEditor.Tests.Unit;

public class PptxTemplateEngineTests : IDisposable
{
    private readonly string _testFilePath = Path.Combine(Path.GetTempPath(), $"test_pptx_template_{Guid.NewGuid()}.pptx");

    [Fact]
    public void ProcessTemplate_ShouldReplaceIfCondition_WhenTrue()
    {
        // Arrange
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddText("{{#if showContent}}Visible Content{{/if}}");
            builder.Save();
        }

        // Act
        using (var builder = PresentationBuilder.Open(_testFilePath))
        {
            builder.ProcessTemplate(new Dictionary<string, object>
            {
                ["showContent"] = true
            });
            builder.Save();
        }

        // Assert
        using var doc = PresentationDocument.Open(_testFilePath, false);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var text = slidePart.Slide!.InnerText;
        Assert.Contains("Visible Content", text);
        Assert.DoesNotContain("{{#if", text);
    }

    [Fact]
    public void ProcessTemplate_ShouldRemoveIfCondition_WhenFalse()
    {
        // Arrange
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddText("{{#if showContent}}Visible Content{{/if}}");
            builder.Save();
        }

        // Act
        using (var builder = PresentationBuilder.Open(_testFilePath))
        {
            builder.ProcessTemplate(new Dictionary<string, object>
            {
                ["showContent"] = false
            });
            builder.Save();
        }

        // Assert
        using var doc = PresentationDocument.Open(_testFilePath, false);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var text = slidePart.Slide!.InnerText;
        Assert.DoesNotContain("Visible Content", text);
        Assert.DoesNotContain("{{#if", text);
    }

    [Fact]
    public void ProcessTemplate_ShouldReplaceIfNotCondition_WhenFalse()
    {
        // Arrange
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddText("{{#ifnot hiddenContent}}Shown Content{{/ifnot}}");
            builder.Save();
        }

        // Act
        using (var builder = PresentationBuilder.Open(_testFilePath))
        {
            builder.ProcessTemplate(new Dictionary<string, object>
            {
                ["hiddenContent"] = false
            });
            builder.Save();
        }

        // Assert
        using var doc = PresentationDocument.Open(_testFilePath, false);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var text = slidePart.Slide!.InnerText;
        Assert.Contains("Shown Content", text);
        Assert.DoesNotContain("{{#ifnot", text);
    }

    [Fact]
    public void ProcessTemplate_ShouldRemoveIfNotCondition_WhenTrue()
    {
        // Arrange
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddText("{{#ifnot hiddenContent}}Shown Content{{/ifnot}}");
            builder.Save();
        }

        // Act
        using (var builder = PresentationBuilder.Open(_testFilePath))
        {
            builder.ProcessTemplate(new Dictionary<string, object>
            {
                ["hiddenContent"] = true
            });
            builder.Save();
        }

        // Assert
        using var doc = PresentationDocument.Open(_testFilePath, false);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var text = slidePart.Slide!.InnerText;
        Assert.DoesNotContain("Shown Content", text);
        Assert.DoesNotContain("{{#ifnot", text);
    }

    [Fact]
    public void ProcessTemplate_ShouldProcessEachLoop()
    {
        // Arrange
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddText("{{#each items}}{name}}-{{/each}}");
            builder.Save();
        }

        // Act
        using (var builder = PresentationBuilder.Open(_testFilePath))
        {
            builder.ProcessTemplate(new Dictionary<string, object>
            {
                ["items"] = new List<Dictionary<string, object>>
                {
                    new() { ["name"] = "Item1" },
                    new() { ["name"] = "Item2" },
                    new() { ["name"] = "Item3" }
                }
            });
            builder.Save();
        }

        // Assert
        using var doc = PresentationDocument.Open(_testFilePath, false);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var text = slidePart.Slide!.InnerText;
        Assert.Contains("Item1", text);
        Assert.Contains("Item2", text);
        Assert.Contains("Item3", text);
        Assert.DoesNotContain("{{#each", text);
    }

    [Fact]
    public void ProcessTemplate_ShouldHandleEmptyEachLoop()
    {
        // Arrange
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddText("{{#each items}}{name}}{{/each}}");
            builder.Save();
        }

        // Act
        using (var builder = PresentationBuilder.Open(_testFilePath))
        {
            builder.ProcessTemplate(new Dictionary<string, object>
            {
                ["items"] = new List<Dictionary<string, object>>()
            });
            builder.Save();
        }

        // Assert
        using var doc = PresentationDocument.Open(_testFilePath, false);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var text = slidePart.Slide!.InnerText;
        Assert.DoesNotContain("{{#each", text);
    }

    [Fact]
    public void ProcessTemplate_ShouldEvaluateNumericComparison()
    {
        // Arrange
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddText("{{#if count > 5}}High{{/if}}");
            builder.Save();
        }

        // Act
        using (var builder = PresentationBuilder.Open(_testFilePath))
        {
            builder.ProcessTemplate(new Dictionary<string, object>
            {
                ["count"] = 10
            });
            builder.Save();
        }

        // Assert
        using var doc = PresentationDocument.Open(_testFilePath, false);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var text = slidePart.Slide!.InnerText;
        Assert.Contains("High", text);
    }

    [Fact]
    public void ProcessTemplate_ShouldEvaluateStringComparison()
    {
        // Arrange
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddText("{{#if status == 'active'}}Active{{/if}}");
            builder.Save();
        }

        // Act
        using (var builder = PresentationBuilder.Open(_testFilePath))
        {
            builder.ProcessTemplate(new Dictionary<string, object>
            {
                ["status"] = "active"
            });
            builder.Save();
        }

        // Assert
        using var doc = PresentationDocument.Open(_testFilePath, false);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var text = slidePart.Slide!.InnerText;
        Assert.Contains("Active", text);
    }

    [Fact]
    public void ProcessTemplate_ShouldEvaluateNotEqualComparison()
    {
        // Arrange
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddText("{{#if status != 'inactive'}}Not Inactive{{/if}}");
            builder.Save();
        }

        // Act
        using (var builder = PresentationBuilder.Open(_testFilePath))
        {
            builder.ProcessTemplate(new Dictionary<string, object>
            {
                ["status"] = "active"
            });
            builder.Save();
        }

        // Assert
        using var doc = PresentationDocument.Open(_testFilePath, false);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var text = slidePart.Slide!.InnerText;
        Assert.Contains("Not Inactive", text);
    }

    [Fact]
    public void ProcessTemplate_ShouldHandleTruthyValues()
    {
        // Arrange
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddText("{{#if hasData}}Has Data{{/if}}");
            builder.Save();
        }

        // Act - test with non-empty string
        using (var builder = PresentationBuilder.Open(_testFilePath))
        {
            builder.ProcessTemplate(new Dictionary<string, object>
            {
                ["hasData"] = "some data"
            });
            builder.Save();
        }

        // Assert
        using var doc = PresentationDocument.Open(_testFilePath, false);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var text = slidePart.Slide!.InnerText;
        Assert.Contains("Has Data", text);
    }

    [Fact]
    public void ProcessTemplate_ShouldHandleFalsyStringValues()
    {
        // Arrange
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddText("{{#if hasData}}Has Data{{/if}}");
            builder.Save();
        }

        // Act - test with empty string
        using (var builder = PresentationBuilder.Open(_testFilePath))
        {
            builder.ProcessTemplate(new Dictionary<string, object>
            {
                ["hasData"] = ""
            });
            builder.Save();
        }

        // Assert
        using var doc = PresentationDocument.Open(_testFilePath, false);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var text = slidePart.Slide!.InnerText;
        Assert.DoesNotContain("Has Data", text);
    }

    [Fact]
    public void ProcessTemplate_ShouldHandleFalsyFalseString()
    {
        // Arrange
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddText("{{#if enabled}}Enabled{{/if}}");
            builder.Save();
        }

        // Act - test with "false" string
        using (var builder = PresentationBuilder.Open(_testFilePath))
        {
            builder.ProcessTemplate(new Dictionary<string, object>
            {
                ["enabled"] = "false"
            });
            builder.Save();
        }

        // Assert
        using var doc = PresentationDocument.Open(_testFilePath, false);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var text = slidePart.Slide!.InnerText;
        Assert.DoesNotContain("Enabled", text);
    }

    [Fact]
    public void ProcessTemplate_ShouldHandleFalsyZeroString()
    {
        // Arrange
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddText("{{#if count}}Has Count{{/if}}");
            builder.Save();
        }

        // Act - test with "0" string
        using (var builder = PresentationBuilder.Open(_testFilePath))
        {
            builder.ProcessTemplate(new Dictionary<string, object>
            {
                ["count"] = "0"
            });
            builder.Save();
        }

        // Assert
        using var doc = PresentationDocument.Open(_testFilePath, false);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var text = slidePart.Slide!.InnerText;
        Assert.DoesNotContain("Has Count", text);
    }

    [Fact]
    public void ProcessTemplate_ShouldHandleMissingVariable()
    {
        // Arrange
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddText("{{#if missing}}Should Not Show{{/if}}");
            builder.Save();
        }

        // Act - no data provided
        using (var builder = PresentationBuilder.Open(_testFilePath))
        {
            builder.ProcessTemplate(new Dictionary<string, object>());
            builder.Save();
        }

        // Assert
        using var doc = PresentationDocument.Open(_testFilePath, false);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var text = slidePart.Slide!.InnerText;
        Assert.DoesNotContain("Should Not Show", text);
    }

    [Fact]
    public void ProcessTemplate_ShouldHandleBooleanTrue()
    {
        // Arrange
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddText("{{#if isActive}}Active{{/if}}");
            builder.Save();
        }

        // Act
        using (var builder = PresentationBuilder.Open(_testFilePath))
        {
            builder.ProcessTemplate(new Dictionary<string, object>
            {
                ["isActive"] = true
            });
            builder.Save();
        }

        // Assert
        using var doc = PresentationDocument.Open(_testFilePath, false);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var text = slidePart.Slide!.InnerText;
        Assert.Contains("Active", text);
    }

    [Fact]
    public void ProcessTemplate_ShouldHandleBooleanFalse()
    {
        // Arrange
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddText("{{#if isActive}}Active{{/if}}");
            builder.Save();
        }

        // Act
        using (var builder = PresentationBuilder.Open(_testFilePath))
        {
            builder.ProcessTemplate(new Dictionary<string, object>
            {
                ["isActive"] = false
            });
            builder.Save();
        }

        // Assert
        using var doc = PresentationDocument.Open(_testFilePath, false);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var text = slidePart.Slide!.InnerText;
        Assert.DoesNotContain("Active", text);
    }

    public void Dispose()
    {
        if (File.Exists(_testFilePath))
        {
            File.Delete(_testFilePath);
        }
    }
}
