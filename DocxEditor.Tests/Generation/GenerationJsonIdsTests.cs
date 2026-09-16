using System.Text.Json;
using System.Text.Json.Nodes;
using OfficeEditor.Core.Generation;
using DocxEditor.Core.Generation;
using DocxEditor.Core.Generation.Schema;
using PptxEditor.Core.Generation;
using PptxEditor.Core.Generation.Schema;
using XlsxEditor.Core.Instructions;

namespace DocxEditor.Tests.Generation;

public sealed class GenerationJsonIdsTests
{
    private const string Deck = """
        {"version":"2.0","design":{"palette":{"ink":"#17245C"},"fonts":{"body":"Arial","display":"Arial"}},"slides":[
          {"type":"container","layout":{"mode":"column"},"children":[{"type":"text","text":"First","size":{"w":120,"h":40}},
            {"type":"text","id":"text-0","text":"Second","size":{"w":120,"h":40}}]}]}
        """;
    private const string Document = """
        {"version":"1.0","sections":[{"header":[{"type":"paragraph","text":"Header"}],"blocks":[
          {"type":"group","blocks":[{"type":"paragraph","text":"Hello"}]}],
          "positioned":[{"type":"textBox","text":"Box","x":20,"y":20,"width":120,"height":40}]}]}
        """;
    private const string Workbook = """
        {"version":"1.0","worksheets":[{"name":"Data","columns":[{"name":"Value","type":"string"}],
          "cells":[{"address":"A1","value":"Hello"},{"address":"A2","value":"World"}],
          "tables":[{"name":"DataTable","range":"A1:A2"}]}]}
        """;

    [Theory]
    [InlineData(Deck, "text-1", "text", "Updated")]
    [InlineData(Document, "paragraph-1", "text", "Updated")]
    [InlineData(Workbook, "cell-1", "value", "Updated")]
    public void Batch_RenamesThenEdits_AndRegenerates(string source, string id, string property, string value)
    {
        var operations = new JsonArray(
            new JsonObject { ["type"] = "rename", ["id"] = id, ["newId"] = "selected" },
            new JsonObject { ["type"] = "set", ["id"] = "selected", ["properties"] = new JsonObject { [property] = value } });
        var result = GenerationJsonEditor.Apply(source, operations.ToJsonString());
        Assert.Equal(value, GenerationJsonIds.GetElement(result, "selected").Properties[property]!.GetValue<string>());
        Assert.Equal(0, GenerationJsonIds.Normalize(result).AddedCount);
        Assert.Equal(result, GenerationJsonEditor.Apply(result, "[]"));
        Assert.Throws<KeyNotFoundException>(() => GenerationJsonIds.GetElement(result, id));
        if (source == Deck) Assert.True(new PptxGenerator().Generate(result).Success);
        else if (source == Document) Assert.NotEmpty(new DocxGenerator().GenerateToBytes(result).Content);
        else Assert.True(XlsxGenerator.Generate(result).IsValid);
    }

    [Theory]
    [InlineData("{}")]
    [InlineData("[null]")]
    [InlineData("[{\"type\":\"delete\",\"id\":\"text-1\"}]")]
    [InlineData("[{\"type\":\"set\",\"id\":\"missing\",\"properties\":{}}]")]
    [InlineData("[{\"type\":\"set\",\"id\":\"text-1\",\"properties\":null}]")]
    [InlineData("[{\"type\":\"rename\",\"id\":\"text-1\",\"newId\":\"text-0\"}]")]
    [InlineData("[{\"type\":\"rename\",\"id\":\"text-1\",\"newId\":\"new\",\"typo\":true}]")]
    [InlineData("[{\"type\":\"set\",\"id\":\"text-1\",\"properties\":{\"text\":\"a\",\"TEXT\":\"b\"}}]")]
    public void Batch_RejectsInvalidOperations(string operations)
    {
        Assert.Throws<JsonException>(() => GenerationJsonEditor.Apply(Deck, operations));
    }

    [Fact]
    public void Batch_RejectsAmbiguousSourceProperties()
    {
        Assert.Throws<JsonException>(() => GenerationJsonEditor.Apply(
            """{"version":"1.0","sections":[{"id":"one","ID":"two"}]}""", "[]"));
    }

    [Fact]
    public void Batch_LaterFailureDoesNotMutateSource()
    {
        var source = GenerationJsonIds.Normalize(Deck).Json;
        var error = Assert.Throws<JsonException>(() => GenerationJsonEditor.Apply(source,
            """[{"type":"rename","id":"text-1","newId":"heading"},{"type":"set","id":"text-1","properties":{"text":"changed"}}]"""));
        Assert.Contains("operations[1]", error.Message);
        Assert.Equal("First", GenerationJsonIds.GetElement(source, "text-1").Properties["text"]!.GetValue<string>());
    }

    [Fact]
    public void Defaults_ReserveAuthoredIdsBeforeAllocating_AndMatchPptxParser()
    {
        var result = GenerationJsonIds.Normalize(Deck);
        Assert.Equal(2, result.AddedCount);
        Assert.Equal(new[] { "slide-0", "text-1", "text-0" }, GenerationJsonIds.Inspect(result.Json).Select(e => e.Id));
        var parsed = new GenerationDocumentParser().Validate(Deck);
        Assert.True(parsed.IsValid, string.Join("\n", parsed.Errors));
        Assert.Equal(result.Json, parsed.NormalizedJson);
        Assert.Equal("slide-0", parsed.Document!.Slides[0].Id);
        Assert.Equal("text-1", parsed.Document.Slides[0].Children[0].Id);
        Assert.DoesNotContain("id", JsonNode.Parse(result.Json)!["design"]!.ToJsonString());
    }

    [Fact]
    public void NormalizedSource_IsByteStableOnSecondPass_AndReorderPreservesIdentity()
    {
        var normalized = GenerationJsonIds.Normalize(Deck).Json;
        var again = GenerationJsonIds.Normalize(normalized);
        Assert.Equal(0, again.AddedCount);
        Assert.Equal(normalized, again.Json);
        var root = JsonNode.Parse(normalized)!;
        var children = root["slides"]![0]!["children"]!.AsArray();
        var first = children[0];
        children.RemoveAt(0);
        children.Add(first);
        var reordered = GenerationJsonIds.Normalize(root.ToJsonString());
        Assert.Equal(0, reordered.AddedCount);
        Assert.Equal("First", GenerationJsonIds.GetElement(reordered.Json, "text-1").Properties["text"]!.GetValue<string>());
    }

    [Theory]
    [InlineData("")]
    [InlineData("Bad name")]
    [InlineData("1bad")]
    [InlineData("name\n")]
    public void InvalidIds_FailWithoutRepair(string id)
    {
        Assert.Throws<JsonException>(() => GenerationJsonIds.Rename(Deck, "text-0", id));
    }

    [Fact]
    public void DuplicateIds_FailWithBothPaths()
    {
        var error = Assert.Throws<JsonException>(() => GenerationJsonIds.Rename(Deck, "text-1", "text-0"));
        Assert.Contains("slides[0].children[0]", error.Message);
        Assert.Contains("slides[0].children[1]", error.Message);
        Assert.Throws<JsonException>(() => GenerationJsonIds.Normalize(Deck.Replace("\"text\":\"First\"", "\"id\":null,\"text\":\"First\"")));
    }

    [Fact]
    public void RenameAndPropertyEdit_TargetOnlyChosenElement()
    {
        var renamed = GenerationJsonIds.Rename(Deck, "text-1", "heading");
        var updated = GenerationJsonIds.SetProperties(renamed, "heading", new JsonObject { ["text"] = "Updated", ["bold"] = true });
        var selected = GenerationJsonIds.GetElement(updated, "heading");
        Assert.Equal("Updated", selected.Properties["text"]!.GetValue<string>());
        Assert.Equal("slide-0", selected.ParentId);
        Assert.Equal("Second", GenerationJsonIds.GetElement(updated, "text-0").Properties["text"]!.GetValue<string>());
        Assert.Throws<KeyNotFoundException>(() => GenerationJsonIds.GetElement(updated, "text-1"));
        // Inspection never exposes the source tree itself.
        selected.Properties["text"] = "Unrelated mutation";
        Assert.Equal("Updated", GenerationJsonIds.GetElement(updated, "heading").Properties["text"]!.GetValue<string>());
        var generated = new PptxGenerator().Generate(updated);
        Assert.True(generated.Success);
        Assert.Equal(updated, generated.NormalizedJson);
    }

    [Fact]
    public void NestedDocxElements_AllReceiveIds_AndRoundTripThroughGeneration()
    {
        var parsed = new DocxGenerationDocumentParser().Validate(Document);
        Assert.True(parsed.IsValid, string.Join("\n", parsed.Errors));
        Assert.NotNull(parsed.NormalizedJson);
        var elements = GenerationJsonIds.Inspect(parsed.NormalizedJson);
        Assert.Equal(new[] { "section-0", "paragraph-0", "group-0", "paragraph-1", "textbox-0" }, elements.Select(e => e.Id));
        var section = parsed.Document!.Sections[0];
        Assert.Equal("section-0", section.Id);
        Assert.Equal("paragraph-0", section.Header[0].Id);
        Assert.Equal("textbox-0", section.Positioned[0].Id);
        var changed = GenerationJsonIds.SetProperties(parsed.NormalizedJson, "textbox-0", new JsonObject { ["x"] = 100, ["text"] = "Moved" });
        var generated = new DocxGenerator().GenerateToBytes(changed);
        Assert.NotEmpty(generated.Content);
        Assert.Equal(changed, generated.NormalizedJson);
        Assert.Equal(0, GenerationJsonIds.Normalize(changed).AddedCount);
        Assert.False(new DocxGenerationDocumentParser().Validate(Document.Replace("\"type\":\"paragraph\"", "\"id\":\"duplicate\",\"type\":\"paragraph\"")).IsValid);
    }

    [Fact]
    public void XlsxIdentities_AreIndependentOfCellAddresses_AndGenerate()
    {
        var parsed = XlsxInstructionParser.ParseAndValidate(Workbook);
        Assert.True(parsed.IsValid, string.Join("\n", parsed.Validation.Errors));
        Assert.NotNull(parsed.NormalizedJson);
        var sheet = parsed.InstructionSet!.Worksheets[0];
        Assert.Equal("worksheet-0", sheet.Id);
        Assert.Equal("cell-0", sheet.Cells![0].Id);
        Assert.Equal("column-0", sheet.Columns![0].Id);
        Assert.Equal("table-0", sheet.Tables![0].Id);
        var updated = GenerationJsonIds.SetProperties(parsed.NormalizedJson, "cell-1", new JsonObject { ["address"] = "A3" });
        Assert.Equal("A3", GenerationJsonIds.GetElement(updated, "cell-1").Properties["address"]!.GetValue<string>());
        var generated = XlsxGenerator.Generate(updated, new XlsxGenerateOptions { ValidatePackage = true });
        Assert.True(generated.IsValid, string.Join("\n", generated.Validation.Errors));
        Assert.Equal(updated, generated.NormalizedJson);
        Assert.False(XlsxInstructionParser.ParseAndValidate(Workbook.Replace("\"address\":", "\"id\":\"duplicate\",\"address\":")).IsValid);
    }

    [Fact]
    public void FileNormalization_PersistsIds_AndDoesNotRewriteCompletedFiles()
    {
        var path = Path.Combine(Path.GetTempPath(), $"oe-ids-{Guid.NewGuid():N}.json");
        try
        {
            File.WriteAllText(path, Document);
            Assert.Equal(5, GenerationJsonIds.NormalizeFile(path).AddedCount);
            File.SetLastWriteTimeUtc(path, new DateTime(2020, 1, 1, 0, 0, 0, DateTimeKind.Utc));
            Assert.Equal(0, GenerationJsonIds.NormalizeFile(path).AddedCount);
            Assert.Equal(2020, File.GetLastWriteTimeUtc(path).Year);
            File.WriteAllText(path, Document.Replace("\"type\":\"paragraph\"", "\"id\":\"duplicate\",\"type\":\"paragraph\""));
            var invalid = File.ReadAllText(path);
            Assert.Throws<JsonException>(() => GenerationJsonIds.NormalizeFile(path));
            Assert.Equal(invalid, File.ReadAllText(path));
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Normalization_PreservesUnicodeNumbersAndWhitespace()
    {
        const string json = "{\r\n  \"version\": \"1.0\", \"sections\": [{\"blocks\": [{\"text\": \"Expérience 中文\", \"type\": \"paragraph\"}]}], \"extra\": 1.00e+2\r\n}\r\n";
        var result = GenerationJsonIds.Normalize(json);
        Assert.Equal(2, result.AddedCount);
        Assert.Contains("Expérience 中文", result.Json);
        Assert.Contains("1.00e+2", result.Json);
        Assert.EndsWith("\r\n", result.Json);
        Assert.Equal(json, result.Json.Replace("\"id\": \"section-0\", ", "").Replace("\"id\": \"paragraph-0\", ", ""));
    }

    [Fact]
    public void MalformedAndAmbiguousJson_DoesNotEscapeValidationContract()
    {
        Assert.False(GenerationJsonIds.IsGenerationDocument("{broken"));
        var ambiguous = Deck.Replace("\"slides\":", "\"Slides\":[],\"slides\":");
        Assert.False(new GenerationDocumentParser().Validate(ambiguous).IsValid);
    }

    [Theory]
    [InlineData("{\"version\":\"1.0\",\"instructions\":[]}")]
    [InlineData("{\"name\":\"web\",\"version\":\"1.0\"}")]
    [InlineData("{\"slides\":[]}")]
    public void UnrelatedJson_IsNotTreatedAsGeneration(string json)
    {
        Assert.False(GenerationJsonIds.IsGenerationDocument(json));
        Assert.Throws<JsonException>(() => GenerationJsonIds.Normalize(json));
    }
}
