using DocxEditor.Core.Builders;
using OfficeEditor.Core.Models;
using DocumentFormat.OpenXml.Packaging;

namespace DocxEditor.Tests.Unit;

public class VariableDetectionTests : IDisposable
{
    private readonly string _testFilePath = Path.Combine(Path.GetTempPath(), $"test_vars_{Guid.NewGuid()}.docx");

    [Fact]
    public void DetectVariables_ShouldFindVariablesInDocument()
    {
        // Arrange
        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.AddParagraph("Hello {{clientName}}");
            builder.AddParagraph("Your order {{orderId}} is ready");
            builder.Save();
        }

        // Act
        List<VariableInfo> variables;
        using (var builder = DocumentBuilder.Open(_testFilePath))
        {
            variables = builder.DetectVariables();
        }

        // Assert
        Assert.Equal(2, variables.Count);
        Assert.Contains(variables, v => v.Name == "clientName");
        Assert.Contains(variables, v => v.Name == "orderId");
    }

    [Fact]
    public void DetectVariables_ShouldFindVariablesWithDefaultValues()
    {
        // Arrange
        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.AddParagraph("Hello {{name|Guest}}");
            builder.Save();
        }

        // Act
        List<VariableInfo> variables;
        using (var builder = DocumentBuilder.Open(_testFilePath))
        {
            variables = builder.DetectVariables();
        }

        // Assert
        Assert.Single(variables);
        Assert.Equal("name", variables[0].Name);
        Assert.Equal("Guest", variables[0].DefaultValue);
    }

    [Fact]
    public void MergeVariables_ShouldReplaceVariables()
    {
        // Arrange
        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.AddParagraph("Hello {{clientName}}");
            builder.Save();
        }

        // Act
        using (var builder = DocumentBuilder.Open(_testFilePath))
        {
            builder.MergeVariables(new Dictionary<string, string>
            {
                ["clientName"] = "John Doe"
            });
            builder.Save();
        }

        // Assert
        using (var doc = WordprocessingDocument.Open(_testFilePath, false))
        {
            var mainPart = doc.MainDocumentPart;
            Assert.NotNull(mainPart);
            var document = mainPart.Document;
            Assert.NotNull(document);
            var body = document.Body;
            Assert.NotNull(body);
            var text = body.InnerText;
            Assert.Contains("John Doe", text);
            Assert.DoesNotContain("{{clientName}}", text);
        }
    }

    [Fact]
    public void MergeVariables_ShouldUseDefaultValues()
    {
        // Arrange
        using (var builder = DocumentBuilder.Create(_testFilePath))
        {
            builder.AddParagraph("Hello {{name|Guest}}");
            builder.Save();
        }

        // Act - merge without providing "name"
        using (var builder = DocumentBuilder.Open(_testFilePath))
        {
            builder.MergeVariables(new Dictionary<string, string>
            {
                ["otherVar"] = "value"
            });
            builder.Save();
        }

        // Assert - the variable replacer uses default value when no data provided
        using (var doc = WordprocessingDocument.Open(_testFilePath, false))
        {
            var mainPart = doc.MainDocumentPart;
            Assert.NotNull(mainPart);
            var document = mainPart.Document;
            Assert.NotNull(document);
            var body = document.Body;
            Assert.NotNull(body);
            var text = body.InnerText;
            // The variable replacer uses default value when no data provided
            Assert.Contains("Hello Guest", text);
        }
    }

    public void Dispose()
    {
        if (File.Exists(_testFilePath))
        {
            File.Delete(_testFilePath);
        }
    }
}
