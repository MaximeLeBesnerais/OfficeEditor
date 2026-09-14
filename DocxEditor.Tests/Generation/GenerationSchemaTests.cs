using System.Text.Json;
using PptxEditor.Core.Generation.Schema;

namespace DocxEditor.Tests.Generation;

public class GenerationSchemaTests
{
    [Fact]
    public void SchemaJson_IsValidJsonWithExpectedContract()
    {
        using var document = JsonDocument.Parse(GenerationSchema.SchemaJson);
        var root = document.RootElement;

        Assert.Equal("https://json-schema.org/draft/2020-12/schema", root.GetProperty("$schema").GetString());
        Assert.Equal("object", root.GetProperty("type").GetString());
        Assert.False(root.GetProperty("additionalProperties").GetBoolean());

        var required = root.GetProperty("required").EnumerateArray().Select(e => e.GetString()).ToList();
        Assert.Contains("version", required);
        Assert.Contains("design", required);
        Assert.Contains("slides", required);

        Assert.Equal("2.0", root.GetProperty("properties").GetProperty("version").GetProperty("const").GetString());
        Assert.Equal(GenerationSchema.Version, "2.0");

        var defs = root.GetProperty("$defs");
        foreach (var def in new[] { "design", "layout", "size", "at", "fill", "gradient", "stroke", "radius", "shadow", "container", "text", "rect", "ellipse", "line", "image", "group", "component", "element" })
        {
            Assert.True(defs.TryGetProperty(def, out _), $"$defs is missing '{def}'.");
        }
    }

    [Fact]
    public void SchemaJson_DeclaresEveryElementTypeAndComponent()
    {
        using var document = JsonDocument.Parse(GenerationSchema.SchemaJson);
        var defs = document.RootElement.GetProperty("$defs");

        Assert.Equal("container", defs.GetProperty("container").GetProperty("properties").GetProperty("type").GetProperty("const").GetString());
        Assert.Equal("text", defs.GetProperty("text").GetProperty("properties").GetProperty("type").GetProperty("const").GetString());

        var lineTypes = defs.GetProperty("line").GetProperty("properties").GetProperty("type").GetProperty("enum")
            .EnumerateArray().Select(e => e.GetString()).ToList();
        Assert.Contains("line", lineTypes);
        Assert.Contains("connector", lineTypes);

        var components = defs.GetProperty("component").GetProperty("properties").GetProperty("type").GetProperty("enum")
            .EnumerateArray().Select(e => e.GetString()).ToList();
        Assert.Equal(
            new[] { "card", "kpi", "title_block", "bullet_list", "divider", "badge", "image_card", "table_block" },
            components);
    }

    [Fact]
    public void SchemaFile_MatchesEmbeddedConstant()
    {
        var schemaPath = FindSchemaFile();

        var fileContent = File.ReadAllText(schemaPath).TrimEnd();
        Assert.Equal(fileContent, GenerationSchema.SchemaJson.TrimEnd());
    }

    [Fact]
    public void SchemaJson_ContainerDef_AcceptsSlideNotes()
    {
        using var document = JsonDocument.Parse(GenerationSchema.SchemaJson);
        var containerProps = document.RootElement.GetProperty("$defs").GetProperty("container").GetProperty("properties");

        var notes = containerProps.GetProperty("notes");
        Assert.Equal("string", notes.GetProperty("type").GetString());
        Assert.Contains("not rendered in the preview", notes.GetProperty("description").GetString());
    }

    [Fact]
    public void SchemaJson_SlideSize_AcceptsPresetsAndCustomObject()
    {
        using var document = JsonDocument.Parse(GenerationSchema.SchemaJson);
        var slideSize = document.RootElement.GetProperty("properties").GetProperty("slideSize");

        Assert.Equal("16:9", slideSize.GetProperty("default").GetString());

        var oneOf = slideSize.GetProperty("oneOf").EnumerateArray().ToList();
        Assert.Equal(2, oneOf.Count);

        // Branch 0: the v1 preset strings.
        var presets = oneOf[0].GetProperty("enum").EnumerateArray().Select(e => e.GetString()).ToList();
        Assert.Equal(new[] { "16:9", "4:3" }, presets);

        // Branch 1: a custom point-dimension object.
        var custom = oneOf[1];
        Assert.Equal("object", custom.GetProperty("type").GetString());
        Assert.False(custom.GetProperty("additionalProperties").GetBoolean());
        var required = custom.GetProperty("required").EnumerateArray().Select(e => e.GetString()).ToList();
        Assert.Equal(new[] { "width", "height" }, required);
        var width = custom.GetProperty("properties").GetProperty("width");
        Assert.Equal("number", width.GetProperty("type").GetString());
        Assert.Equal(0, width.GetProperty("exclusiveMinimum").GetDouble());
        Assert.Equal(4032, width.GetProperty("maximum").GetDouble());
        var height = custom.GetProperty("properties").GetProperty("height");
        Assert.Equal("number", height.GetProperty("type").GetString());
        Assert.Equal(0, height.GetProperty("exclusiveMinimum").GetDouble());
        Assert.Equal(4032, height.GetProperty("maximum").GetDouble());
    }

    private static string FindSchemaFile()
    {
        var relative = Path.Combine("PptxEditor.Core", "Generation", "Schema", GenerationSchema.FileName);
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        for (var depth = 0; depth < 10 && directory is not null; depth++)
        {
            var candidate = Path.Combine(directory.FullName, relative);
            if (File.Exists(candidate))
            {
                return candidate;
            }
            directory = directory.Parent;
        }
        throw new FileNotFoundException($"Could not locate {relative} above {AppContext.BaseDirectory}.");
    }
}
