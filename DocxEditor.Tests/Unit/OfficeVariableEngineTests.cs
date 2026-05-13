using OfficeEditor.Core.Variables;

namespace DocxEditor.Tests.Unit;

public class OfficeVariableEngineTests
{
    [Fact]
    public void VariableDetector_Scan_ShouldFindVariablesWithDefaultsAndTrimNames()
    {
        // Arrange
        var detector = new VariableDetector();

        // Act
        var variables = detector.Scan("Hello {{ name |Guest}} from {{city}}", "body");

        // Assert
        Assert.Collection(variables,
            variable =>
            {
                Assert.Equal("name", variable.Name);
                Assert.Equal("{{ name |Guest}}", variable.FullMatch);
                Assert.Equal("body", variable.Location);
                Assert.Equal("Guest", variable.DefaultValue);
            },
            variable =>
            {
                Assert.Equal("city", variable.Name);
                Assert.Equal("{{city}}", variable.FullMatch);
                Assert.Equal("body", variable.Location);
                Assert.Null(variable.DefaultValue);
            });
    }

    [Fact]
    public void VariableDetector_Scan_ShouldSuppressDuplicatesByNameAndLocation()
    {
        // Arrange
        var detector = new VariableDetector();

        // Act
        var variables = detector.Scan("{{name}} {{ name|Guest}} {{other}}", "header");

        // Assert
        Assert.Equal(2, variables.Count);
        Assert.Contains(variables, variable => variable.Name == "name" && variable.Location == "header");
        Assert.Contains(variables, variable => variable.Name == "other" && variable.Location == "header");
    }

    [Fact]
    public void VariableDetector_ScanElements_ShouldRetainSameVariableInDifferentLocationsAndAggregate()
    {
        // Arrange
        var detector = new VariableDetector();
        var elements = new[]
        {
            (Text: "{{name}} {{name}}", Location: "header"),
            (Text: "No variables here", Location: "body"),
            (Text: "{{name}} {{total|0}}", Location: "footer")
        };

        // Act
        var variables = detector.ScanElements(elements);

        // Assert
        Assert.Equal(3, variables.Count);
        Assert.Contains(variables, variable => variable.Name == "name" && variable.Location == "header");
        Assert.Contains(variables, variable => variable.Name == "name" && variable.Location == "footer");
        Assert.Contains(variables, variable => variable.Name == "total" && variable.Location == "footer" && variable.DefaultValue == "0");
    }

    [Fact]
    public void VariableDetector_Scan_ShouldReturnEmptyListWhenNoVariablesExist()
    {
        // Arrange
        var detector = new VariableDetector();

        // Act
        var variables = detector.Scan("Plain text with {single} braces", "body");

        // Assert
        Assert.Empty(variables);
    }

    [Fact]
    public void VariableReplacer_Replace_ShouldPreferProvidedValueOverDefault()
    {
        // Arrange
        var replacer = new VariableReplacer();

        // Act
        var result = replacer.Replace("Hello {{name|Guest}}", new Dictionary<string, string>
        {
            ["name"] = "Ada"
        });

        // Assert
        Assert.Equal("Hello Ada", result);
    }

    [Fact]
    public void VariableReplacer_Replace_ShouldUseDefaultWhenValueIsMissing()
    {
        // Arrange
        var replacer = new VariableReplacer();

        // Act
        var result = replacer.Replace("Hello {{name|Guest}}", new Dictionary<string, string>());

        // Assert
        Assert.Equal("Hello Guest", result);
    }

    [Theory]
    [InlineData("{{name|}}")]
    [InlineData("{{name}}")]
    public void VariableReplacer_Replace_ShouldLeaveUnknownVariableUnchangedWhenNoUsableDefaultExists(string variable)
    {
        // Arrange
        var replacer = new VariableReplacer();

        // Act
        var result = replacer.Replace($"Hello {variable}", new Dictionary<string, string>());

        // Assert
        Assert.Equal($"Hello {variable}", result);
    }

    [Fact]
    public void VariableReplacer_Replace_ShouldReplaceMultipleVariablesAndTrimNames()
    {
        // Arrange
        var replacer = new VariableReplacer();

        // Act
        var result = replacer.Replace("{{ first }} {{last}} from {{city|Paris}}", new Dictionary<string, string>
        {
            ["first"] = "Grace",
            ["last"] = "Hopper"
        });

        // Assert
        Assert.Equal("Grace Hopper from Paris", result);
    }

    [Fact]
    public void TemplateEngine_Process_ShouldHandleIfAndIfNotForBooleanValues()
    {
        // Arrange
        var engine = new TemplateEngine();
        var template = "{{#if enabled}}Enabled{{/if}}|{{#if disabled}}Disabled{{/if}}|{{#ifnot disabled}}Not disabled{{/ifnot}}|{{#ifnot enabled}}Not enabled{{/ifnot}}";

        // Act
        var result = engine.Process(template, new Dictionary<string, object>
        {
            ["enabled"] = true,
            ["disabled"] = false
        });

        // Assert
        Assert.Equal("Enabled||Not disabled|", result);
    }

    [Theory]
    [InlineData("", false)]
    [InlineData("false", false)]
    [InlineData("FALSE", false)]
    [InlineData("0", false)]
    [InlineData("hello", true)]
    public void TemplateEngine_Process_ShouldEvaluateStringTruthiness(string value, bool shouldRender)
    {
        // Arrange
        var engine = new TemplateEngine();

        // Act
        var result = engine.Process("{{#if value}}Visible{{/if}}", new Dictionary<string, object>
        {
            ["value"] = value
        });

        // Assert
        Assert.Equal(shouldRender ? "Visible" : string.Empty, result);
    }

    [Theory]
    [InlineData("count == 5", "Equals")]
    [InlineData("count != 6", "NotEquals")]
    [InlineData("count > 4", "Greater")]
    [InlineData("count < 6", "Less")]
    [InlineData("count >= 5", "GreaterOrEqual")]
    [InlineData("count <= 5", "LessOrEqual")]
    public void TemplateEngine_Process_ShouldEvaluateNumericComparisons(string condition, string content)
    {
        // Arrange
        var engine = new TemplateEngine();

        // Act
        var result = engine.Process($"{{{{#if {condition}}}}}{content}{{{{/if}}}}", new Dictionary<string, object>
        {
            ["count"] = 5
        });

        // Assert
        Assert.Equal(content, result);
    }

    [Theory]
    [InlineData("status == 'beta'", "Active")]
    [InlineData("status != \"inactive\"", "NotInactive")]
    [InlineData("status > alpha", "AfterAlpha")]
    [InlineData("status < omega", "BeforeOmega")]
    [InlineData("status >= beta", "AtLeastBeta")]
    [InlineData("status <= beta", "AtMostBeta")]
    public void TemplateEngine_Process_ShouldEvaluateStringComparisons(string condition, string content)
    {
        // Arrange
        var engine = new TemplateEngine();

        // Act
        var result = engine.Process($"{{{{#if {condition}}}}}{content}{{{{/if}}}}", new Dictionary<string, object>
        {
            ["status"] = "beta"
        });

        // Assert
        Assert.Equal(content, result);
    }

    [Fact]
    public void TemplateEngine_Process_ShouldTreatMissingComparisonVariableAsFalse()
    {
        // Arrange
        var engine = new TemplateEngine();

        // Act
        var result = engine.Process("Before {{#if missing == 'yes'}}Visible{{/if}} After", new Dictionary<string, object>());

        // Assert
        Assert.Equal("Before  After", result);
    }

    [Fact]
    public void TemplateEngine_Process_ShouldRenderLoopsFromDictionaryListAndEmptyNullValues()
    {
        // Arrange
        var engine = new TemplateEngine();
        var data = new Dictionary<string, object>
        {
            ["items"] = new List<Dictionary<string, object>>
            {
                new() { ["name"] = "One", ["value"] = 1 },
                new() { ["name"] = "Two", ["value"] = null! }
            }
        };

        // Act
        var result = engine.Process("{{#each items}}{name}:{value}{{/each}}", data);

        // Assert
        Assert.Equal("One:1\nTwo:", result);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TemplateEngine_Process_ShouldRemoveLoopWhenSourceIsMissingOrNotAList(bool includeNonListSource)
    {
        // Arrange
        var engine = new TemplateEngine();
        var data = includeNonListSource
            ? new Dictionary<string, object> { ["items"] = "not a list" }
            : new Dictionary<string, object>();

        // Act
        var result = engine.Process("Before {{#each items}}{name}{{/each}} After", data);

        // Assert
        Assert.Equal("Before  After", result);
    }
}
